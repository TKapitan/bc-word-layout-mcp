# Security Policy

## Supported versions

Only the latest released version receives security fixes.

| Version | Supported |
|---|---|
| latest 1.x release | ✅ |
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
repo). The mitigations that exist for that today: `ResourceLimits` caps XML part size, part count
and recursion depth (zip-bomb / deep-nesting protection); `ExternalRelationshipStripper` removes
external relationships (`attachedTemplate`, linked images/OLE, mail-merge sources) from every merged
copy before any converter opens it, so a poisoned layout cannot make Word reach out to a UNC path or
URL; and every mutating tool writes to a staged copy, validates, and commits with an atomic rename —
the original file is never opened writable. The design rationale lives in
[`docs/SOLUTION-DESIGN.md`](docs/SOLUTION-DESIGN.md) §6 and [`docs/adr/`](docs/adr/).

Two residual surfaces are accepted deliberately and worth knowing about: the OpenXml SDK's own
parse of `document.xml` and header/footer parts is SDK-internal (our byte/depth caps cover every
part **this repo's code** loads — a pathologically large main part costs the invoking user their
own memory, nothing more); and `render_preview_pages` hands the PDF at the supplied path to the
bundled PDFium, so pointing it at a hostile PDF exercises PDFium's parser (output size is clamped
via the DPI/page caps; the native libraries are kept current via Dependabot).

Reports about these mitigations being insufficient against a crafted file are exactly the kind of
report we want.
