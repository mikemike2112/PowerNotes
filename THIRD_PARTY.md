# Dipendenze di terze parti

Solo dipendenze con licenza MIT, Apache-2.0 o BSD. Licenza verificata sul campo `license` del file `.nuspec` di ogni pacchetto.

## Applicazione (runtime)

| Pacchetto | Versione | Licenza | Uso |
|---|---|---|---|
| DocumentFormat.OpenXml (+ .Framework) | 3.5.1 | MIT | Lettura/scrittura OOXML, `OpenXmlValidator` |
| PDFsharp | **6.2.4** | MIT | Scrittura del PDF (testo vettoriale, font incorporati, immagini) |
| Microsoft.Extensions.Logging (+ Abstractions, DependencyInjection.Abstractions) | 8.0.x | MIT | Dipendenza transitiva di PDFsharp |
| System.Security.Cryptography.Pkcs | 10.0.13 | MIT | Dipendenza transitiva di PDFsharp, alzata a una versione senza advisory noti |
| System.IO.Packaging | 10.0.x | MIT | Dipendenza transitiva di DocumentFormat.OpenXml |

> **Attenzione, versione di PDFsharp bloccata.** PDFsharp 6.5.0 (e successive) è pubblicato con licenza **AGPL-3.0-only**, incompatibile con il vincolo MIT/Apache/BSD del progetto. La versione 6.2.4 è MIT. Non aggiornare PDFsharp oltre la 6.2.x senza ricontrollare la licenza.

## Test e strumenti

| Pacchetto | Versione | Licenza |
|---|---|---|
| xunit, xunit.core, xunit.assert, xunit.extensibility.* | 2.9.3 | Apache-2.0 |
| xunit.runner.visualstudio | 3.1.4 | Apache-2.0 |
| Microsoft.NET.Test.Sdk | 17.14.1 | MIT |
| coverlet.collector | 6.0.4 | MIT |

`tools/CompatTool` usa solo API del sistema (Windows.Data.Pdf, WPF) più il progetto Core.

## Font

L'applicazione **non include font**: misura e disegna con i font installati nel PC (cartella Windows e cache dei font cloud di Office, in sola lettura). Nessun nome, logo, icona o risorsa Microsoft è incluso nel repository. L'icona è generata da `tools/make-icon.ps1`.
