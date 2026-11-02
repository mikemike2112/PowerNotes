using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using PowerNotes.Core.Layout;
using PowerNotes.Core.Model;

namespace PowerNotes.Core.Edit;

public enum TableSel { Cell, Row, Column, Table }
public enum TableFit { Content, Window, Fixed }

/// <summary>Comandi di tabella della scheda «Layout di tabella». Ogni comando è una sola voce di annullamento.</summary>
public sealed partial class EditorSession
{
    const string NotInTable = "Posiziona il cursore in una tabella.";

    static List<int> GridW(Table t) => (t.GetFirstChild<TableGrid>() ?? throw new InvalidOperationException("Tabella senza griglia.")).Elements<GridColumn>().Select(g => int.TryParse(g.Width?.Value, out var w) ? w : 0).ToList();

    static void SetGrid(Table t, List<int> w)
    {
        var g = t.GetFirstChild<TableGrid>()!;
        g.RemoveAllChildren();
        foreach (var x in w) g.Append(new GridColumn { Width = Math.Max(1, x).ToString() });
    }

    /// <summary>Larghezza di ogni cella = somma delle colonne di griglia coperte.</summary>
    static void SyncWidths(Table t)
    {
        var grid = GridW(t);
        foreach (var row in t.Elements<TableRow>())
            foreach (var (cell, col, span) in Cells(row))
            {
                var pr = cell.TableCellProperties ?? cell.PrependChild(new TableCellProperties());
                pr.TableCellWidth = new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = grid.Skip(col).Take(span).Sum().ToString() };
            }
    }

    (TableCell cell, TableRow row, Table tbl) InTbl() => CellOf(Caret.P) ?? throw new InvalidOperationException(NotInTable);

    /// <summary>Rettangolo selezionato (righe e colonne di griglia, estremi inclusi in r, esclusi in g1); con una sola cella è la cella del caret.</summary>
    (Table tbl, List<TableRow> rows, int r0, int r1, int g0, int g1) SelRect()
    {
        var b = InTbl();
        var a = CellOf(Anchor.P);
        if (a == null || a.Value.tbl != b.tbl) a = b;
        var rows = b.tbl.Elements<TableRow>().ToList();
        (int r, int c0, int c1) P((TableCell cell, TableRow row, Table tbl) x) { var e = Cells(x.row).First(y => y.Cell == x.cell); return (rows.IndexOf(x.row), e.Col, e.Col + e.Span); }
        var pa = P(a.Value); var pb = P(b);
        return (b.tbl, rows, Math.Min(pa.r, pb.r), Math.Max(pa.r, pb.r), Math.Min(pa.c0, pb.c0), Math.Max(pa.c1, pb.c1));
    }

    /// <summary>Nuove colonne di griglia al posto di una: le celle che la attraversano (in ogni riga) si allargano di k-1.</summary>
    static void SplitGridCol(Table tbl, int col, int k)
    {
        var grid = tbl.GetFirstChild<TableGrid>()!;
        var cols = grid.Elements<GridColumn>().ToList();
        int w = int.TryParse(cols[col].Width?.Value, out var gw) ? gw : 1440;
        int part = Math.Max(1, w / k);
        cols[col].Width = Math.Max(1, w - part * (k - 1)).ToString();
        OpenXmlElement last = cols[col];
        for (int i = 1; i < k; i++) { var ng = new GridColumn { Width = part.ToString() }; last.InsertAfterSelf(ng); last = ng; }
        foreach (var row in tbl.Elements<TableRow>())
            foreach (var (cell, c0, sp) in Cells(row))
                if (c0 <= col && col < c0 + sp)
                    (cell.TableCellProperties ?? cell.PrependChild(new TableCellProperties())).GridSpan = new GridSpan { Val = sp + k - 1 };
    }

    // ---------- Unione: Dividi celle / Dividi tabella ----------
    /// <summary>Divide la cella del caret in parts colonne (il contenuto resta nella prima). Aggiunge colonne di griglia se servono.</summary>
    public void SplitCell(int parts)
    {
        if (parts < 2 || parts > 12) throw new ArgumentOutOfRangeException(nameof(parts), "Celle da 2 a 12.");
        var c = InTbl();
        if (c.cell.TableCellProperties?.VerticalMerge != null) throw new InvalidOperationException("Dividi celle non è disponibile per celle con unione verticale.");
        Begin();
        var me = Cells(c.row).First(x => x.Cell == c.cell);
        if (me.Span < parts) SplitGridCol(c.tbl, me.Col, parts - me.Span + 1);
        me = Cells(c.row).First(x => x.Cell == c.cell);
        int baseSpan = me.Span / parts, rem = me.Span % parts;
        var cells = new List<TableCell> { c.cell };
        OpenXmlElement last = c.cell;
        var lastPara = c.cell.Elements<Paragraph>().LastOrDefault();
        for (int i = 1; i < parts; i++)
        {
            var nc = new TableCell();
            if (c.cell.TableCellProperties != null) { var pr = (TableCellProperties)c.cell.TableCellProperties.CloneNode(true); pr.GridSpan = null; nc.Append(pr); }
            var np = new Paragraph();
            if (lastPara?.ParagraphProperties != null) { var pp = (ParagraphProperties)lastPara.ParagraphProperties.CloneNode(true); pp.SectionProperties = null; np.Append(pp); }
            nc.Append(np);
            last.InsertAfterSelf(nc); last = nc; cells.Add(nc);
        }
        for (int i = 0; i < parts; i++)
        {
            int sp = baseSpan + (i < rem ? 1 : 0);
            (cells[i].TableCellProperties ?? cells[i].PrependChild(new TableCellProperties())).GridSpan = sp > 1 ? new GridSpan { Val = sp } : null;
        }
        SyncWidths(c.tbl);
        Changed();
    }

    /// <summary>Divide la tabella in due alla riga del caret (che apre la seconda); fra le due resta un paragrafo vuoto.</summary>
    public void SplitTable()
    {
        var c = InTbl();
        var rows = c.tbl.Elements<TableRow>().ToList();
        int r = rows.IndexOf(c.row);
        if (r <= 0) throw new InvalidOperationException("Posiziona il cursore in una riga che non sia la prima.");
        if (rows[r].Elements<TableCell>().Any(x => x.TableCellProperties?.VerticalMerge is { } vm && vm.Val?.Value != MergedCellValues.Restart))
            throw new InvalidOperationException("Non si può dividere la tabella dove ci sono celle unite verticalmente.");
        Begin();
        var nt = new Table();
        foreach (var ch in c.tbl.ChildElements.Where(e => e is TableProperties or TableGrid)) nt.Append(ch.CloneNode(true));
        foreach (var row in rows.Skip(r)) { row.Remove(); nt.Append(row); }
        var gap = new Paragraph();
        c.tbl.InsertAfterSelf(gap); gap.InsertAfterSelf(nt);
        Set(new Pos(gap, 0)); Changed();
    }

    // ---------- Dimensioni cella ----------
    /// <summary>Altezza stimata di una riga dal layout: la cella più alta (somma delle righe di testo).</summary>
    static double RowContentH(LayoutResult lay, TableRow row)
    {
        double best = 0;
        foreach (var cell in row.Elements<TableCell>())
            best = Math.Max(best, cell.Descendants<Paragraph>().Sum(p => lay.LinesOf(p).Sum(l => l.Height)));
        var th = row.TableRowProperties?.GetFirstChild<TableRowHeight>()?.Val?.Value;
        return Math.Max(best, th != null ? th.Value / 20.0 : 0);
    }

    static void SetRowHeight(TableRow row, uint twips)
    {
        var pr = row.TableRowProperties ?? row.PrependChild(new TableRowProperties());
        var th = pr.GetFirstChild<TableRowHeight>() ?? pr.AppendChild(new TableRowHeight());
        th.Val = twips; th.HeightType = HeightRuleValues.AtLeast;
    }

    /// <summary>Rende uguali le righe selezionate (tutte se ne è selezionata una sola): altezza minima = media delle altezze.</summary>
    public void DistributeRows()
    {
        var s = SelRect();
        int a = s.r0, b = s.r1;
        if (a == b) { a = 0; b = s.rows.Count - 1; }
        if (a == b) return;
        var lay = Layout;
        double total = 0;
        for (int i = a; i <= b; i++) total += RowContentH(lay, s.rows[i]);
        uint tw = (uint)Math.Clamp(Math.Ceiling(total / (b - a + 1) * 20), 20, 31680);
        Begin();
        for (int i = a; i <= b; i++) SetRowHeight(s.rows[i], tw);
        Changed();
    }

    /// <summary>Rende uguali le colonne selezionate (tutte se ne è selezionata una sola), a parità di larghezza totale.</summary>
    public void DistributeColumns()
    {
        var s = SelRect();
        var grid = GridW(s.tbl);
        int g0 = s.g0, g1 = Math.Min(s.g1, grid.Count);
        if (g1 - g0 <= 1) { g0 = 0; g1 = grid.Count; }
        int n = g1 - g0;
        if (n <= 1) return;
        int total = grid.Skip(g0).Take(n).Sum(), each = total / n;
        Begin();
        for (int i = g0; i < g1; i++) grid[i] = i == g1 - 1 ? total - each * (n - 1) : each;
        SetGrid(s.tbl, grid); SyncWidths(s.tbl);
        Changed();
    }

    /// <summary>Larghezza utile (twip) dove sta la tabella: area di testo della sezione, o la larghezza attuale se la tabella è dentro una cella.</summary>
    int AvailTwips(Table tbl)
    {
        if (tbl.Parent is TableCell) return GridW(tbl).Sum();
        var sec = SectionFor(tbl.Descendants<Paragraph>().First());
        var pm = sec.GetFirstChild<PageMargin>();
        double w = (sec.GetFirstChild<PageSize>()?.Width?.Value ?? 11906) - (pm?.Left?.Value ?? 1440) - (pm?.Right?.Value ?? 1440) - (pm?.Gutter?.Value ?? 0);
        if (sec.GetFirstChild<Columns>() is { ColumnCount.Value: > 1 } cols && int.TryParse(cols.Space?.Value, out var sp)) w = (w - sp * (cols.ColumnCount.Value - 1)) / cols.ColumnCount.Value;
        return (int)Math.Max(720, w);
    }

    static void ScaleGrid(List<int> g, int target)
    {
        int cur = Math.Max(1, g.Sum()); int acc = 0;
        for (int i = 0; i < g.Count; i++) { g[i] = i == g.Count - 1 ? Math.Max(1, target - acc) : Math.Max(1, (int)Math.Round((double)g[i] * target / cur)); acc += g[i]; }
    }

    /// <summary>Adatta: al contenuto, alla finestra (100% dell'area di testo) o larghezza fissa.</summary>
    public void AutoFit(TableFit mode)
    {
        var tbl = InTbl().tbl;
        var lay = Layout;
        Begin();
        var grid = GridW(tbl);
        var tp = tbl.GetFirstChild<TableProperties>() ?? tbl.PrependChild(new TableProperties());
        switch (mode)
        {
            case TableFit.Window:
                ScaleGrid(grid, AvailTwips(tbl));
                tp.TableWidth = new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" }; tp.TableLayout = null;
                break;
            case TableFit.Fixed:
                tp.TableLayout = new TableLayout { Type = TableLayoutValues.Fixed };
                tp.TableWidth = new TableWidth { Type = TableWidthUnitValues.Dxa, Width = grid.Sum().ToString() };
                break;
            default:
                var rows = tbl.Elements<TableRow>().ToList();
                for (int j = 0; j < grid.Count; j++)
                {
                    double w = 0; bool any = false;
                    foreach (var row in rows)
                        foreach (var (cell, col, span) in Cells(row))
                            if (col == j && span == 1)
                            {
                                any = true;
                                foreach (var p in cell.Elements<Paragraph>())
                                    foreach (var ln in lay.LinesOf(p))
                                    {
                                        var segs = ln.Segs.Where(x => !x.Synthetic).ToList();
                                        if (segs.Count > 0) w = Math.Max(w, segs.Max(x => x.X + x.Width) - segs.Min(x => x.X));
                                    }
                            }
                    if (any) grid[j] = (int)Math.Max(400, Math.Ceiling(w * 20) + 220);
                }
                tp.TableWidth = new TableWidth { Type = TableWidthUnitValues.Auto, Width = "0" }; tp.TableLayout = null;
                break;
        }
        SetGrid(tbl, grid); SyncWidths(tbl);
        Changed();
    }

    // ---------- Dati ----------
    /// <summary>Ordina le righe della tabella secondo la colonna del caret (numerico se tutti i valori sono numeri). La prima riga resta ferma se è un'intestazione ripetuta.</summary>
    public void SortTable(bool descending)
    {
        var c = InTbl();
        var tbl = c.tbl;
        var rows = tbl.Elements<TableRow>().ToList();
        if (rows.SelectMany(r => r.Elements<TableCell>()).Any(x => x.TableCellProperties?.GridSpan?.Val?.Value > 1 || x.TableCellProperties?.VerticalMerge != null))
            throw new InvalidOperationException("Ordina non è disponibile per tabelle con celle unite.");
        int gcol = Cells(c.row).First(x => x.Cell == c.cell).Col;
        var data = rows.Skip(rows[0].TableRowProperties?.GetFirstChild<TableHeader>() != null ? 1 : 0).ToList();
        if (data.Count < 2) return;
        string Key(TableRow r) => string.Join("\n", r.Elements<TableCell>().ElementAtOrDefault(gcol)?.Elements<Paragraph>().Select(ParaText.Text) ?? []).Trim();
        decimal? Num(string k) => decimal.TryParse(k.Replace(',', '.'), NumberStyles.Number, CultureInfo.InvariantCulture, out var d) ? d : null;
        var keys = data.ToDictionary(r => r, Key);
        bool numeric = keys.Values.Any(k => k.Length > 0) && keys.Values.All(k => k.Length == 0 || Num(k) != null);
        int Cmp(TableRow x, TableRow y) => numeric ? (Num(keys[x]) ?? decimal.MinValue).CompareTo(Num(keys[y]) ?? decimal.MinValue) : string.Compare(keys[x], keys[y], StringComparison.InvariantCultureIgnoreCase);
        var sorted = (descending ? data.OrderByDescending(r => r, Comparer<TableRow>.Create(Cmp)) : data.OrderBy(r => r, Comparer<TableRow>.Create(Cmp))).ToList();
        if (sorted.SequenceEqual(data)) return;
        Begin();
        foreach (var r in data) r.Remove();
        foreach (var r in sorted) tbl.Append(r);
        Changed();
    }

    /// <summary>Attiva/disattiva «Ripeti righe di intestazione» dalla prima riga fino alla riga selezionata. True se ora è attivo.</summary>
    public bool ToggleRepeatHeader()
    {
        var s = SelRect();
        var sel = s.rows.Take(s.r1 + 1).ToList();
        bool all = sel.All(r => r.TableRowProperties?.GetFirstChild<TableHeader>() != null);
        Begin();
        foreach (var r in sel)
        {
            if (all) r.TableRowProperties?.GetFirstChild<TableHeader>()?.Remove();
            else { var pr = r.TableRowProperties ?? r.PrependChild(new TableRowProperties()); if (pr.GetFirstChild<TableHeader>() == null) pr.Append(new TableHeader()); }
        }
        Changed();
        return !all;
    }

    public bool HasRepeatHeader => CellOf(Caret.P) is { } c && c.tbl.Elements<TableRow>().FirstOrDefault()?.TableRowProperties?.GetFirstChild<TableHeader>() != null;

    /// <summary>Converte la tabella in testo: una riga di paragrafo per riga di tabella, celle separate da tabulazione.</summary>
    public void ConvertTableToText()
    {
        var tbl = InTbl().tbl;
        if (tbl.Descendants<Table>().Any()) throw new InvalidOperationException("Converti in testo non è disponibile per tabelle che ne contengono altre.");
        Begin();
        var made = new List<Paragraph>();
        foreach (var row in tbl.Elements<TableRow>())
        {
            var np = new Paragraph();
            var first = row.Descendants<Paragraph>().FirstOrDefault();
            if (first?.ParagraphProperties != null) { var pp = (ParagraphProperties)first.ParagraphProperties.CloneNode(true); pp.SectionProperties = null; np.Append(pp); }
            bool firstCell = true;
            foreach (var cell in row.Elements<TableCell>())
            {
                if (!firstCell) np.Append(new Run(new TabChar()));
                firstCell = false;
                bool firstPara = true;
                foreach (var para in cell.Elements<Paragraph>())
                {
                    if (!firstPara) np.Append(new Run(new Break()));
                    firstPara = false;
                    foreach (var ch in para.ChildElements.Where(e => e is not ParagraphProperties).ToList()) { ch.Remove(); np.Append(ch); }
                }
            }
            made.Add(np);
        }
        foreach (var p in made) tbl.InsertBeforeSelf(p);
        tbl.Remove();
        Set(new Pos(made[0], 0)); Changed();
    }

    // ---------- Tabella: Seleziona / Proprietà ----------
    public void SelectTablePart(TableSel what)
    {
        var c = InTbl();
        var rows = c.tbl.Elements<TableRow>().ToList();
        var me = Cells(c.row).First(x => x.Cell == c.cell);
        TableCell first = c.cell, last = c.cell;
        switch (what)
        {
            case TableSel.Row: first = c.row.Elements<TableCell>().First(); last = c.row.Elements<TableCell>().Last(); break;
            case TableSel.Table: first = rows[0].Elements<TableCell>().First(); last = rows[^1].Elements<TableCell>().Last(); break;
            case TableSel.Column:
                TableCell? At(TableRow r) => Cells(r).FirstOrDefault(x => x.Col <= me.Col && me.Col < x.Col + x.Span).Cell;
                first = rows.Select(At).First(x => x != null)!; last = rows.Select(At).Last(x => x != null)!; break;
        }
        var fp = first.Descendants<Paragraph>().First(); var lp = last.Descendants<Paragraph>().Last();
        Anchor = new Pos(fp, 0); Caret = new Pos(lp, ParaText.Length(lp)); Pending = null;
    }

    public sealed record TableProps(string Align, int WidthPct, double RowHeightPt);

    public TableProps GetTableProps()
    {
        var c = InTbl();
        var tp = c.tbl.GetFirstChild<TableProperties>();
        var jc = tp?.TableJustification?.Val?.Value;
        string al = jc == TableRowAlignmentValues.Center ? "Centro" : jc == TableRowAlignmentValues.Right ? "Destra" : "Sinistra";
        var w = tp?.TableWidth;
        int pct = w?.Type?.Value == TableWidthUnitValues.Pct && int.TryParse(w.Width?.Value, out var pw) ? pw / 50 : 0;
        double h = c.row.TableRowProperties?.GetFirstChild<TableRowHeight>()?.Val?.Value / 20.0 ?? 0;
        return new TableProps(al, pct, h);
    }

    /// <summary>Allineamento (Sinistra/Centro/Destra), larghezza in % dell'area di testo (0 = automatica) e altezza minima delle righe (pt, 0 = automatica).</summary>
    public void SetTableProps(string align, int widthPct, double rowHeightPt)
    {
        if (widthPct < 0 || widthPct > 100) throw new ArgumentOutOfRangeException(nameof(widthPct), "Larghezza da 0 a 100%.");
        if (double.IsNaN(rowHeightPt) || rowHeightPt < 0 || rowHeightPt > 1584) throw new ArgumentOutOfRangeException(nameof(rowHeightPt), "Altezza riga da 0 a 1584 pt.");
        if (align is not ("Sinistra" or "Centro" or "Destra")) throw new ArgumentException("Allineamento non valido.", nameof(align));
        var tbl = InTbl().tbl;
        Begin();
        var tp = tbl.GetFirstChild<TableProperties>() ?? tbl.PrependChild(new TableProperties());
        tp.TableJustification = align == "Sinistra" ? null : new TableJustification { Val = align == "Centro" ? TableRowAlignmentValues.Center : TableRowAlignmentValues.Right };
        if (widthPct == 0) { if (tp.TableWidth?.Type?.Value == TableWidthUnitValues.Pct) tp.TableWidth = new TableWidth { Type = TableWidthUnitValues.Auto, Width = "0" }; }
        else
        {
            var g = GridW(tbl); ScaleGrid(g, AvailTwips(tbl) * widthPct / 100); SetGrid(tbl, g); SyncWidths(tbl);
            tp.TableWidth = new TableWidth { Type = TableWidthUnitValues.Pct, Width = (widthPct * 50).ToString() };
        }
        foreach (var r in tbl.Elements<TableRow>())
        {
            if (rowHeightPt > 0) SetRowHeight(r, (uint)Math.Round(rowHeightPt * 20));
            else r.TableRowProperties?.GetFirstChild<TableRowHeight>()?.Remove();
        }
        Changed();
    }
}
