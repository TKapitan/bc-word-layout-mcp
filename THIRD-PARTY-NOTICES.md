# Third-party notices

`bc-word-layout-mcp` is licensed under the [MIT License](LICENSE). It incorporates the third-party
components below. License expressions were verified against each package's NuGet metadata on
2026-08-01 (versions as pinned by the project files on that date); all are compatible with
distribution under this repository's MIT license. Apache-2.0 and BSD-licensed components require
their notices to travel with distributions — this file is that notice.

## Runtime dependencies (redistributed in self-contained builds)

| Component | Version | License | Project |
|---|---|---|---|
| ModelContextProtocol (+ `.Core`) | 1.4.1 | **Apache-2.0** | <https://github.com/modelcontextprotocol/csharp-sdk> |
| DocumentFormat.OpenXml (+ `.Framework`) | 3.5.1 | MIT | <https://github.com/dotnet/Open-XML-SDK> |
| Microsoft.Extensions.Hosting (+ transitive `Microsoft.Extensions.*`, `System.Diagnostics.EventLog`, `System.IO.Packaging`) | 10.0.x | MIT | <https://github.com/dotnet/runtime> |
| Microsoft.Extensions.AI.Abstractions | 10.5.2 | MIT | <https://github.com/dotnet/extensions> |
| PDFtoImage | 5.3.0 | MIT | <https://github.com/sungaila/PDFtoImage> |
| SkiaSharp (+ `SkiaSharp.NativeAssets.*`) | 4.150.1 | MIT | <https://github.com/mono/SkiaSharp> |
| bblanchon.PDFium.Win32 | 152.0.7961 | **Apache-2.0** (packaging) | <https://github.com/bblanchon/pdfium-binaries> |
| .NET runtime (self-contained publish payload) | 10.0.x | MIT | <https://github.com/dotnet/runtime> |

### Bundled native libraries

- **Skia** (the native graphics library inside `SkiaSharp.NativeAssets.*`): BSD-3-Clause,
  Copyright (c) Google LLC — <https://github.com/google/skia/blob/main/LICENSE>.
- **PDFium** (the native PDF library packaged by `bblanchon.PDFium.Win32`): BSD-3-Clause,
  Copyright The PDFium Authors — <https://pdfium.googlesource.com/pdfium/+/main/LICENSE>.

## Test / development-only dependencies (not redistributed)

| Component | Version | License |
|---|---|---|
| xunit, xunit.runner.visualstudio | 2.5.3 | Apache-2.0 |
| coverlet.collector | 6.0.0 | MIT |
| Microsoft.NET.Test.Sdk | 17.14.1 | MIT |

## Corpus layouts

Nine of the ten Word layout files in `tests/corpus/` are captures of report layouts Microsoft
publishes as open source in [`microsoft/BCApps`](https://github.com/microsoft/BCApps) under the
**MIT License, Copyright (c) Microsoft Corporation**. The per-file source mapping is recorded in
[`tests/corpus/PROVENANCE.md`](tests/corpus/PROVENANCE.md); the tenth file is this project's own
material, published under the repository license.

## Optional external programs (not bundled)

`preview_layout` can drive **Microsoft Word** (COM automation) or **LibreOffice** (`soffice` CLI)
for PDF conversion if — and only if — they are already installed on the machine. Neither is
included in, nor required by, any distribution of this project.
