# Security Policy

## Supported versions

Only the latest released version receives security fixes.

| Version | Supported |
|---|---|
| the latest release | ✅ |
| anything older | ❌ |

## Reporting a vulnerability

**Please do not open a public issue for a security problem.**

Use GitHub's private vulnerability reporting:
**[Security → Report a vulnerability](https://github.com/TKapitan/bc-word-layout-mcp/security/advisories/new)**.
The report reaches the maintainer privately. You should get an acknowledgement within 7 days; fixes
are prioritized by severity and released as a patch version.

## Threat model, in one paragraph

This is a **local stdio tool**: no network listener, no credentials, no telemetry. It reads and
writes files **with the invoking user's own permissions** at whatever paths the calling agent passes
— so point it only at repositories and layouts you trust, exactly as you would any local build tool.
The primary adversarial input is a **hostile `.docx` layout or schema file** (e.g. from a cloned
repo). The mitigations that exist for that today:

- **Resource caps before any parse.** Every `.docx` is opened through one gate (`LayoutPackage`)
  that streams each package part through a non-recursive reader *before* the OpenXml SDK parses
  it, capping XML nesting depth, per-part and total XML size, and part count. The SDK's own part
  loads are recursive, and a few-KB file nested deeply enough would otherwise overflow the stack,
  which no code can catch. With the gate it fails one call instead. `ResourceLimits` also caps
  the parts and recursive walks this repo's own code handles.
- **A render copy with no way out.** `preview_layout` only ever hands a converter a merged
  *copy*, and before it does, everything in that copy that could make Word or LibreOffice reach
  out to a UNC path or URL is removed:
  - external relationships (`attachedTemplate`, linked images/OLE, subDocument/frame references)
  - every field that isn't computed from the document itself: `INCLUDEPICTURE`, `INCLUDETEXT`,
    `LINK`, `DDE` and the like are unlinked to their cached result, including fields whose type
    is hidden behind a nested field
  - mail-merge, update-fields-on-open and schema-library settings
  - `altChunk` imports
  - VML location attributes

  Plain click-to-follow hyperlinks are kept.
- **No in-place writes.** Every mutating tool writes to a staged copy, validates it, and commits
  with an atomic rename. The original file is never opened writable.

The design rationale lives in [`docs/SOLUTION-DESIGN.md`](docs/SOLUTION-DESIGN.md) §6 and
[`docs/adr/`](docs/adr/).

Residual surfaces, accepted deliberately and worth knowing about:

- **PDFium.** `render_preview_pages` hands the PDF at the supplied path to the bundled PDFium, so
  pointing it at a hostile PDF exercises PDFium's parser. Output is bounded: at most 10 pages per
  call, at most 300 DPI, and at most 5000 px per side and 20 MP per page, with a larger page
  rendered smaller and flagged `downscaled`. The native libraries are kept current via Dependabot.
- **The converter's own behaviour.** The hardening covers every reference *the package carries*.
  It does not change what Word or LibreOffice do of their own accord, such as Office downloading
  a Microsoft cloud font a layout names. That traffic goes to the vendor, not to a host the layout
  chooses.

Reports about these mitigations being insufficient against a crafted file are exactly the kind of
report we want.
