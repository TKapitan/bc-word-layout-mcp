# Third-party notices

`bc-word-layout-mcp` is licensed under the [MIT License](LICENSE). It incorporates the third-party
components below. Their license expressions were checked against each package's NuGet metadata, most
recently on 2026-10-17; all are compatible with distribution under this repository's MIT license.
Apache-2.0 and BSD-licensed components require their notices to travel with distributions — this file
is that notice, and it ships next to `LICENSE` in every NuGet package and release zip.

The tables name components, not versions: the versions in use are the ones the project files pin, and a
version bump does not change a license. A test fails the build when a package the server ships is
missing from the runtime table, so a new dependency cannot ship without its notice.

## Runtime dependencies (redistributed in the NuGet tool packages and the release zips)

| Component | License | Project |
|---|---|---|
| ModelContextProtocol (+ `.Core`) | **Apache-2.0** | <https://github.com/modelcontextprotocol/csharp-sdk> |
| DocumentFormat.OpenXml (+ `.Framework`, which brings `System.IO.Packaging`) | MIT | <https://github.com/dotnet/Open-XML-SDK> |
| Microsoft.Extensions.Hosting (+ transitive `Microsoft.Extensions.*`, `System.Diagnostics.EventLog`) | MIT | <https://github.com/dotnet/runtime> |
| Microsoft.Extensions.AI.Abstractions | MIT | <https://github.com/dotnet/extensions> |
| PDFtoImage | MIT | <https://github.com/sungaila/PDFtoImage> |
| SkiaSharp (+ `SkiaSharp.NativeAssets.*`) | MIT | <https://github.com/mono/SkiaSharp> |
| bblanchon.PDFium.Win32 | **Apache-2.0** (packaging) | <https://github.com/bblanchon/pdfium-binaries> |
| .NET runtime (the release zips' self-contained publish payload) | MIT | <https://github.com/dotnet/runtime> |

### Bundled native libraries

- **Skia** (the native graphics library inside `SkiaSharp.NativeAssets.*`): BSD-3-Clause,
  Copyright (c) Google LLC — <https://github.com/google/skia/blob/main/LICENSE>.
- **PDFium** (the native PDF library packaged by `bblanchon.PDFium.Win32`): BSD-3-Clause,
  Copyright The PDFium Authors — <https://pdfium.googlesource.com/pdfium/+/main/LICENSE>.

## Test / development-only dependencies (not redistributed)

| Component | License |
|---|---|
| xunit | Apache-2.0 |
| xunit.runner.visualstudio | Apache-2.0 |
| coverlet.collector | MIT |
| Microsoft.NET.Test.Sdk | MIT |

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
