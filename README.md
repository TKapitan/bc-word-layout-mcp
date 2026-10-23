# bc-word-layout-mcp

An MCP (Model Context Protocol) server that lets AI agents **create, edit, validate, and preview
Microsoft Word report layouts for Microsoft Dynamics 365 Business Central** through deterministic,
typed tools — instead of free-hand OOXML editing, which is unreliable in a file format where a
single misplaced element silently corrupts the document.

**28 typed MCP tools** over stdio: inspect any layout, validate it (structurally, and via a real
dry-run merge), render an offline mock PDF preview the agent can look at, and author edits — fields,
labels, repeater tables, nested detail rows, table structure, full lifecycle refresh. One uniform
response envelope; nothing throws across the MCP boundary; every mutating edit is validated *before*
it is saved and rejected outright if it would corrupt the file.

- [`skills/al-word-layout/SKILL.md`](skills/al-word-layout/SKILL.md) — the companion skill: the
  intended agent workflow (inspect → edit → validate → preview → sandbox verify).
- [`skills/al-word-layout-design/SKILL.md`](skills/al-word-layout-design/SKILL.md) — the design
  skill: what a new layout should *look like* — archetype skeletons and observed BC document
  conventions, each tagged with the stock corpus layout or BC-verified build it comes from.
- [`skills/al-word-layout-from-design/SKILL.md`](skills/al-word-layout-from-design/SKILL.md) — from a
  finished Word design plus the report's empty layout to a bound layout: what to ask a consultant for,
  mapping sample text to dataset fields, binding it, checking it.
- [`skills/al-word-layout-composite/SKILL.md`](skills/al-word-layout-composite/SKILL.md) — composable
  layouts: what the body, header/footer and theme files contain, and the style contract between them.
- [`docs/SOLUTION-DESIGN.md`](docs/SOLUTION-DESIGN.md) — architecture and design rationale.
- [`docs/adr/`](docs/adr/) — the deliberate design decisions, as ADRs.

## Install

The server reads and writes layout files **with your own permissions, at paths the calling agent
supplies** — point it only at repositories and layouts you trust, as you would any local build tool.

### Option A — NuGet package via `dnx` (needs the .NET 10 SDK)

The .NET 10 SDK ships [`dnx`](https://learn.microsoft.com/dotnet/core/tools/dnx), which fetches and
runs the packaged server on demand — no install step, version-pinnable.

**VS Code** (`.vscode/mcp.json`, or user-level `mcp.json`):

```json
{
  "servers": {
    "bc-word-layout": {
      "type": "stdio",
      "command": "dnx",
      "args": ["BcWordLayout.Mcp", "--yes"]
    }
  }
}
```

**Claude Code** (`.mcp.json` in a project, or user scope via the one-liner):

```json
{
  "mcpServers": {
    "bc-word-layout": {
      "command": "dnx",
      "args": ["BcWordLayout.Mcp", "--yes"]
    }
  }
}
```

```pwsh
claude mcp add bc-word-layout -- dnx BcWordLayout.Mcp --yes
```

Unpinned, `dnx` resolves the latest stable release at each launch, so updates arrive without config
edits. Append an exact version — `BcWordLayout.Mcp@1.2.3` — to pin for reproducibility; a pinned
config stays on that version until you edit it. Pinning is also the only way to run a pre-release,
which latest-stable resolution skips.

Prefer a pinned global install over on-demand fetch? `dotnet tool install -g BcWordLayout.Mcp`
installs the same server as the `bc-word-layout-mcp` command; use that as the `command` in the
blocks above (no `args`).

### Option B — self-contained binaries (no .NET required)

Each [GitHub Release](https://github.com/TKapitan/bc-word-layout-mcp/releases) carries `win-x64` and
`win-arm64` zips (the whole publish folder — the exe does not run alone) plus an install script that
unpacks to a stable location, verifies the checksum, and prints the ready-to-paste config block:

```pwsh
irm https://github.com/TKapitan/bc-word-layout-mcp/releases/latest/download/install.ps1 | iex
```

For a specific version, pre-releases included, run that release's own script with `-Version`:

```pwsh
& ([scriptblock]::Create((irm https://github.com/TKapitan/bc-word-layout-mcp/releases/download/v1.2.3/install.ps1))) -Version 1.2.3
```

The documented config then points at the fixed path
`%LOCALAPPDATA%\bc-word-layout-mcp\current\BcWordLayout.McpHost.exe`.

### Option C — from source (contributors)

The repo-root [`.mcp.json`](.mcp.json) launches the server from your working tree with
`dotnet run --project src/BcWordLayout.McpHost --artifacts-path .mcp-run`, so there's no publish step
while iterating. The separate `.mcp-run/` build folder matters on Windows: a running server locks
the DLLs it loaded, and building into the normal `bin/` folders would stop `dotnet build` and
`dotnet test` while any session is open.

Claude Code starts every `.mcp.json` server when a session opens, whether or not it's used. So the
committed [`.claude/settings.json`](.claude/settings.json) lists this one in `disabledMcpjsonServers`,
and sessions in the repo don't start it. To use your build as a server, add it under your own local
scope, which that setting doesn't cover:

```pwsh
claude mcp add --scope local bc-word-layout-dev -- dotnet run --project src/BcWordLayout.McpHost --artifacts-path .mcp-run
```

Remove it again with `claude mcp remove --scope local bc-word-layout-dev`. Either way, the server
runs the code as built when the session started; restart the session to pick up later edits. See
[Building from source](#building-from-source).

### The companion skills and plugin

Four skills ride along with the server: `al-word-layout` teaches an agent the intended workflow and
the supported matrix, `al-word-layout-design` teaches what a from-scratch layout should look
like (archetype skeletons and observed BC document conventions), `al-word-layout-from-design`
turns a finished Word design plus the report's empty layout into a bound layout — so a functional
consultant can get a working layout without touching OOXML or AL — and `al-word-layout-composite`
covers composable layouts: the report body, the reusable header/footer layout and the theme. The server is fully usable without
them, but agents drive the tools noticeably better with them. The plugin installs the skills and
the server entry together, and this repository is a plugin marketplace for **both** ecosystems:

- **Claude Code:** `/plugin marketplace add TKapitan/bc-word-layout-mcp`, then install the
  `bc-word-layout` plugin it offers.
- **VS Code / GitHub Copilot (agent plugins, preview):** add
  `"chat.plugins.marketplaces": ["TKapitan/bc-word-layout-mcp"]` to your **user** `settings.json`,
  then install the `bc-word-layout` plugin from the marketplace. The same mechanism serves
  Copilot CLI.
- **Manual copy (the floor — no plugin system involved):** copy the `skills/` folders
  (`al-word-layout/`, `al-word-layout-design/`, `al-word-layout-from-design/`, `al-word-layout-composite/`) into your project's `.claude/skills/` or `.github/skills/` (or
  user-level `~/.claude/skills/`). `SKILL.md` agent skills are an open standard — Claude Code,
  VS Code, and Copilot CLI all scan these locations natively; the server is then registered
  separately per Option A/B above.

### Configuration (optional)

Label classification works out of the box for both shapes found in real BC layouts: columns whose
name ends in `Lbl`/`_Lbl` (BC's documented convention), plus every direct column of a data item
named `Labels` — a dedicated labels data item, common in older/converted reports whose columns are
suffixed `Caption`/`Label` or not at all. The second rule is self-scoping: it only affects layouts
that actually contain a `Labels` data item, so no per-report configuration is needed. Two
environment variables override this default convention, read once at server startup; an invalid
value is ignored (logged to stderr) and the default kept:

- `BCWL_LABEL_SUFFIXES` — comma-separated label suffixes, e.g. `Lbl,Caption` (default: `Lbl`).
- `BCWL_LABELS_DATA_ITEM` — retargets the labels data-item rule to a different data-item name, or
  disables it entirely with the special value `-` (default: `Labels`).

```json
{
  "mcpServers": {
    "bc-word-layout": {
      "command": "dnx",
      "args": ["BcWordLayout.Mcp", "--yes"],
      "env": {
        "BCWL_LABEL_SUFFIXES": "Lbl,Caption",
        "BCWL_LABELS_DATA_ITEM": "ReportLabels"
      }
    }
  }
}
```

## Supported platforms

**Windows (x64 / arm64) is the supported platform.** PDF preview conversion uses Microsoft Word via
COM automation when Word is installed. A LibreOffice converter (`soffice` CLI) exists in the code as
the fallback and is what non-Windows use would rely on — but it is **untested and unverified end to
end**, so Linux and macOS are currently *unsupported*, not "supported via LibreOffice" (the
non-Windows code paths compile and their unit tests run in CI as a rot-guard; that is a build claim,
not a support claim).

**Neither Word nor LibreOffice is required** to use the server. Without a converter,
`preview_layout` still merges the sample data and returns the merged `.docx` path — it reports
`conversionOk: false` with a `conversionError` explaining why, instead of failing the call (and
removes the PDF an earlier preview left, so no stale render remains next to the new merge). Every
other tool is conversion-free and works everywhere the server runs.

## Tools

| Tool | Purpose |
|---|---|
| `get_layout_info` | Inspect a layout: report name/id, dataset namespace, `storeItemID`, the full control inventory (kind, alias, tag, `xPath`, part, parent repeater), per-part details (kind, header/footer role `default`/`first`/`even`, and which part a `partName`-less header/footer edit targets) plus a `hasTitlePage` flag, every table's rows (one per `w:tr`, so each row of a multi-row repeater item is addressable, tagged with the repeater it repeats with) and cells — each cell with both its physical `colIndex` (what the cell tools take) and its `gridColumn`/`gridSpan` (what the column tools take; they differ on any spanned row), each row with its `gridBefore`/`gridAfter` — and a quick validation summary. |
| `list_dataset_fields` | List the dataset hierarchy (data items + columns) from a `.docx` layout or a standalone schema `.xml`; flags labels and, for a layout, bound/unbound columns. |
| `validate_layout` | Validate a layout: `level=quick` (structural + binding checks, including a `repeater-downgraded` error for a repeater a Word save turned into a plain content control, a `top-level-repeater-group` error for a block-level repeater over the top-level data item without the `w:group` wrapper Business Central needs to print it, a warning when a layout containing repeaters declares a Word compatibility mode below 15, and a `duplicate-control-id` warning when two content controls share a `w:id`) or `level=full` (adds a dry-run merge that surfaces every unresolved binding as a finding). |
| `preview_layout` | Mock preview: merges deterministic (or supplied) sample data into a working copy and converts it to PDF via Word COM or LibreOffice. |
| `render_preview_pages` | Render pages of a preview PDF (the `pdfPath` from `preview_layout`) as PNG images returned inline as MCP image content blocks, so the calling agent can visually inspect the preview (default 3 pages, cap 10 per call; page onward via `firstPage`; each page capped at 5000 px a side and 20 MP, a larger one rendered smaller and flagged `downscaled`). Pass `outputDir` to also write each page as a `.png` a human can open — optionally with `inlineImages=false` to skip the token cost of the inline copies. |
| `create_layout` | Create a new blank layout `.docx` from a schema, optionally starting from a branded/styled (unbound) template. A blank build ships deterministic default typography (Calibri 11 pt via a scaffolded styles part) so it renders with the same font in Word and in BC, plus a settings part declaring `compatibilityMode` 15 so its repeaters survive being opened and saved in Word. |
| `insert_field` | Insert a plain-text FIELD content control bound to a dataset path, in the body, a header, or a footer. |
| `insert_label` | Insert a plain-text LABEL content control bound to a label-suffixed (`Lbl`/`_Lbl`) dataset path. |
| `insert_text` | Insert plain STATIC text — a literal run, not a content control: the separator space, colon or `" / "` that glues two inline controls together (without it, chained controls render as `Document NoDOCU-0150`). Creates no control, so it has no `controlId` and cannot be targeted by `remove_control`. |
| `insert_page_number` | Insert Word `PAGE`/`NUMPAGES` field codes — the "Page X / Y" chrome every stock BC document header carries (by default the full `X / Y` shape; `includeTotal=false` for the current page alone). Field codes are plain runs, so no `controlId`. |
| `insert_repeater_table` | Insert a complete repeater table (label/static header row + one bound data row) for a repeating data item, in the body — the flagship editing tool. Draws the BC-native look by default (no grid, one rule under the header row; `borders` also offers none or a full grid) and can reference a table style the layout defines; optional per-column widths and alignments, pinned with a fixed table layout so Word renders the grid as specified. |
| `insert_repeater_row` | Add a nested DETAIL ROW repeater inside an existing repeater's item — the standard BC shape for per-line detail (assembly components, serial/lot nos): a row under each line, aligned to the parent grid via spans, repeating once per parent row. |
| `insert_picture` | Insert a PICTURE content control bound to a picture dataset path (e.g. `/Header/CompanyPicture`) — the logo placeholder a from-scratch layout needs; BC fills it at render time. |
| `insert_table` | Insert a plain (unbound) table, in the body, a header or a footer — the building block for non-repeating layout sections (address columns, label/value header grids, right-anchored totals blocks, the stock header and footer tables). Borderless by default, with the same `borders` and `tableStyle` options as `insert_repeater_table`. Column widths are pinned with a fixed table layout, so Word keeps the grid you asked for. Returns the new table's index for `set_cell_text`/`insert_field` to fill. |
| `wrap_in_repeater` | Wrap a run of existing body blocks in a block-level repeater over a data item, so they print once per record: a page per customer in a statement, a paragraph per work-description line. Over a top-level item it adds the `w:group` Business Central needs. |
| `remove_control` | Remove a content control (field/label/repeater/picture/unbound) by its `w:id`, searching the body and every header/footer; `keepText=true` on a block-level repeater unwraps it and keeps its content. |
| `set_cell_text` | Set (replace) the PLAIN TEXT of a table cell, addressed by table/row/column index — e.g. re-label a line-items column header. Rejects cells that hold a bound control. |
| `set_text_format` | Restyle text that is ALREADY in the layout without changing it: a control by id, a table cell, a whole row, or one exact piece of text, with the same `format` object the insert tools take (bold, italic, underline, size, font, colour, character style) and an optional alignment. |
| `clear_cell_text` | Remove all PLAIN TEXT from a table cell, leaving a valid empty cell (the cell/column is preserved) — e.g. blank an unwanted column-header label. |
| `insert_subtotal_row` | Append a STATIC per-group row at the end of an existing repeater's item — the stock BC shape for group subtotal (and spacer) rows: it renders once per PARENT row, after the nested detail rows, never per detail row. |
| `insert_table_row` | Insert ONE STATIC (non-repeating) row into an existing table — the stock BC shape for the totals block INSIDE the line-items table (spacer cells, a spanned bound caption cell, a right-aligned bound amount cell). Cells bind FULL dataset paths; the row renders exactly once, never per data row. |
| `set_column_widths` | Set a table's column widths (one twip value per grid column); `gridSpan`-aware — each cell is resized to the sum of the columns it spans. |
| `insert_column` | Add a new column at any grid position (append by default) — a bound field/label cell (as `insert_field`/`insert_label`) or a plain-text column. Adds the `w:gridCol` and one cell per row; a spanned cell in a row with no content there is widened instead. The table keeps its total width: the new column's width comes out of the existing columns. |
| `remove_column` | Remove a whole grid column — addressed by the `gridColumn` `get_layout_info` reports for its cells, not their `colIndex`: drops each row's covering cell (a bound cell too) or decrements its `gridSpan`, and removes the `w:gridCol`. The summary says per row which of the two happened. |
| `merge_cells` | Horizontally merge a run of adjacent cells in one row into a single `gridSpan` cell (keeps the first cell's content). |
| `split_cells` | Horizontally split one spanned cell back into single-column cells. |
| `set_cell_format` | Set the appearance of a whole row's cells, or one cell: the per-cell rules a BC document's look is actually made of (a line above a totals row, an underline on one cell) and/or a fill colour (a shaded header row). |
| `refresh_xml_part` | Update a layout's BC dataset XML part to a new schema in place; remaps/orphans bindings, re-points any binding left naming a foreign report's namespace (which BC refuses to upload), and reports newly-unbound fields. Also the repair for a dataset part that no longer parses (`validate_layout` names the part and the parser error). |

### Response envelope

Every tool returns a uniform envelope:

```json
{ "ok": true, "data": { "...": "tool-specific payload" } }
```

Members that would be `null` are omitted rather than sent as `null`: a success has no `error` key, a
failure has no `data` key, and an optional field inside `data` is simply absent.

One tool extends (never replaces) this shape: `render_preview_pages` returns the same JSON envelope as
its result's **first text content block**, followed by one MCP **image content block** per rendered page
— the only way a tool result can carry images over MCP. Its envelope JSON is serialized with the same
options the SDK applies to every other tool, so parsers see an identical shape either way.

On failure, `error` is populated with `{ code, message, hint }` — `hint` is always
a non-empty, agent-actionable next step (which argument to fix and its valid values, or which
inspection tool to call first). Nothing throws across the MCP boundary. The 22 mutating tools
(`insert_field`, `insert_label`, `insert_text`, `insert_page_number`, `insert_picture`, `insert_table`,
`insert_repeater_table`, `insert_repeater_row`, `insert_subtotal_row`, `insert_table_row`,
`wrap_in_repeater`, `remove_control`, `set_cell_text`, `set_text_format`, `clear_cell_text`, `set_column_widths`, `insert_column`,
`remove_column`, `merge_cells`, `split_cells`, `set_cell_format`, `refresh_xml_part`) additionally run
`OpenXmlValidator` *before* saving — a write that would introduce a **new** structural error is
rejected (`edit_would_corrupt`) and the file on disk is left untouched. Every one of them is also
checked by a grid-consistency guard that rejects (`edit_would_break_table`, file untouched) any edit
leaving a row's cells inconsistent with its `w:tblGrid` — a corruption class `OpenXmlValidator`
accepts silently. Every successful mutating response also carries a fresh post-edit
`quickValidation` summary. (`create_layout` writes a new file instead: it checks the layout it built
and returns `quickValidation` too, but a failure of that check is a server bug, reported as
`internal_error`, never `edit_would_corrupt`.)

**Envelope boundary: malformed calls get the envelope too.** The SDK looks the tool up and binds the
JSON arguments onto its typed parameters before any tool body runs. The server's call-tool filter wraps
that step, so a call that never reaches a body still returns `isError: true` with the same
`{ok:false,error:{code,message,hint}}` shape:

- **Unknown tool name:** `invalid_argument` naming the tool; the hint points at `tools/list`.
- **Missing required argument:** `invalid_argument` naming it; the hint lists the tool's required
  parameters.
- **Wrong JSON type** (`"abc"` or `1.5` for an integer, `"true"` for a boolean, `null` for a number or
  boolean): `invalid_argument` naming the parameter, the type it needs, and what the call sent.
- **Unknown argument name:** refused with `invalid_argument` rather than ignored, naming the closest
  real parameter (`layoutpart` → `layoutPart`; names are case-sensitive) and listing them all. An
  ignored misspelling would silently fall back to the parameter's default.
- **Relative path:** every path parameter must be absolute; a relative one (or a drive- or
  root-relative one like `C:Report.docx`) is refused with `invalid_argument` naming the parameter,
  instead of resolving against the server's working directory.
- **Explicit `null` for a string:** an optional enum-like string (`layoutPart`, `borders`, `borderStyle`,
  `converter`, `level`) takes its documented default, exactly as if it were omitted; a required one
  (`locationType`, `mode`) is refused with `invalid_argument`.

Only a request too malformed for the SDK to parse as a tool call at all (no tool `name`, or `arguments`
that is not a JSON object) fails earlier, as a JSON-RPC error with no tool result; MCP client libraries
do not produce that shape.

## Mock preview: not a substitute for a sandbox render

`preview_layout` merges **deterministic sample data** (or a real dataset via `dataOverridesPath` —
both the layout's own data-store part shape and the report UI's *Send to → XML* export are accepted,
and the export's `decimalformatter` columns are formatted the way BC itself renders them) into
a working copy of the layout and converts it to PDF **entirely offline** — never through the real
Business Central report engine. That's good enough to catch structural and binding mistakes on every
edit, but captions, fonts, and exact pagination can differ from a genuine BC render. Every
`preview_layout` response repeats this as a `disclaimer` string. A dataset printed for several
documents previews only the first, and a `multiple-top-level-records` warning says so: BC renders the
report body once per top-level record, the mock does not.

**Final sign-off on any layout is always a real Business Central sandbox render** (AL build → publish
to dev sandbox → run the report), never this tool's PDF alone.

**What has been validated against a real Business Central sandbox:** a pack of layouts made with
these tools was uploaded to real BC sandboxes and rendered by the real report engine — validated
against **BC 28.0 and 28.3** (Cronus AU) in two rounds. Round 1 had five items, one of them built
from scratch; round 2 had seven: three from-scratch builds (P01, P06, P07) and four edits of stock
layouts. [`docs/FIDELITY-CHECKLIST.md`](docs/FIDELITY-CHECKLIST.md) records both runs and what they
measured. Every item passed: BC accepts what these tools author, bindings resolve, repeaters
expand, and the mock preview's structure matched the real render on the dimensions checked. That
is what it proves — and no more: seven layouts, one BC version family, one localization. The mock preview remains a mock; the per-release
fidelity procedure and its open manual dimensions live in the same checklist.

## Public contract and versioning

This project follows [semver](https://semver.org). From v1.0.0 on, the **public API is: tool names,
tool parameters (names, types, required-ness), the `{ok, data, error}` envelope shape, and the error
codes.** Removing or changing any of those is a breaking change and lands only in a major version;
new tools, new optional parameters, new `data` fields, and new error codes are additive (minor).
Error `message`/`hint` texts are not part of the contract — never parse them.

- **Business Central versions:** validated against **BC 28.0 and 28.3** (see above). Word report layouts are
  a long-stable BC format, so the tools are *expected* to work with any BC version that supports
  Word layouts — reports of issues on other versions are welcome — but the compatibility commitment
  extends only to versions a validation has actually covered.
- **MCP SDK cadence:** the server pins an exact version of the official `ModelContextProtocol` C#
  SDK (the host's project file names it) and upgrades it in **minor** releases, with any behavioral consequence called out in the
  changelog. The envelope boundary above is pinned by a host-level test that runs the real server
  pipeline over the SDK's own transport, so an upgrade that changes it fails the suite.

## Not in v1 / deferred

- **Hide-if-empty BC add-in conditional controls** (Hide Field if Zero, Hide Empty Table, Hide Empty
  Table Row, layout comment) — no `set_hide_if_empty` tool. The row and table variants are captured
  in two corpus layouts and wait on a sandbox check of what BC treats as empty ([#123](https://github.com/TKapitan/bc-word-layout-mcp/issues/123));
  Hide Field if Zero and layout comments wait on capturing the BC Word add-in's OOXML ([#8](https://github.com/TKapitan/bc-word-layout-mcp/issues/8),
  [ADR-0005](docs/adr/0005-emit-only-observed-ooxml.md)).
- **Repeater tables in headers/footers** — `insert_repeater_table` only creates them in the document
  body; a pre-existing/hand-authored one is still detected by `validate_layout` with a warning rather
  than silently accepted. (NESTED repeaters in the body are supported: insert the inner table into a
  cell of the outer table's `dataRowIndex` row.)
- **Real XLF caption translations** — not planned. With generated sample data, `preview_layout` fills
  labels from the dataset's own label columns, falling back to humanized element names, so captions read
  differently from a real BC render. Pass a real exported dataset via `dataOverridesPath` when caption
  text matters: the dataset carries BC's own captions, and they then match the real render.
- **Table-structure changes** — largely supported (`insert_column`, `remove_column`,
  `set_column_widths`, horizontal `merge_cells`/`split_cells` — all `gridSpan`-aware,
  grid-column-addressed, guarded against desyncing a row from its `w:tblGrid`; rows skipping grid
  columns via `w:gridBefore`/`w:gridAfter` are handled). Still deferred and refused with a clear
  message, pending real-add-in OOXML capture: any table using **vertical merges** (`w:vMerge`) —
  except `set_cell_format`, which is cosmetic and accepts them. See
  [issue #9](https://github.com/TKapitan/bc-word-layout-mcp/issues/9).
- **Cosmetic formatting beyond text and cells** (margins, page setup, styles/branding) — no dedicated
  tool. Text has the `format` object (bold, italic, underline, size, font, colour, character style) on
  every tool that writes it, plus `set_text_format` to restyle text already in the layout; cells have
  the alignment arguments and `set_cell_format`'s rules and fills; and a blank `create_layout` ships
  deterministic default typography (Calibri 11 pt via its scaffolded styles part). For the rest,
  restyle via `create_layout`'s `templatePath` or hand-edit the OOXML directly, then run
  `validate_layout` to confirm the edit didn't break structure or bindings. Restyling in **Word** is
  safe for a layout this tool created — it declares `compatibilityMode` 15, so Word preserves the
  repeating sections; a layout that declares a lower mode (an older build, or a `templatePath` shell
  that brought one) loses them on save, which is what `validate_layout`'s `compatibility-mode` warning
  is for.
- **Generating the dataset schema itself** — deliberate, not a gap
  ([ADR-0006](docs/adr/0006-schema-transplanted-never-synthesized.md)): `create_layout` and
  `refresh_xml_part` take their schema from a BC-produced artifact (a compiler-generated layout, an
  exported stock layout, or an exported schema `.xml`) and transplant its part content byte-for-byte,
  so "the layout's XML part == the compiler's XML part" holds by construction. Reimplementing the AL
  compiler's dataset-to-schema mapping (element-name derivation, `IncludeCaption` → `_Lbl` columns,
  namespace construction) would reintroduce exactly the drift class the transplant rule eliminates.
  Practical consequence: a brand-new layout starts with one AL build (the compiler creates the
  referenced layout `.docx`, dataset part included) or one export from BC — the skill's §1 documents
  the workflow.
- **RDL layouts, Excel layouts** — out of scope entirely; this server is Word-only.
- **Any BC-connected upload** of a layout to a tenant — no public API exists for this; it is UI-only in
  Business Central, a deliberate scope decision rather than a gap to script around.

## Building from source

Prerequisite: the **.NET 10 SDK** (and the nuget.org package source configured).

```pwsh
dotnet build     # warnings fail the build (TreatWarningsAsErrors + .NET analyzers)
dotnet test
```

The suite includes Word COM PDF-conversion tests and OOXML snapshot tests against the corpus of 10
real BC layouts (`tests/corpus/` — see [`PROVENANCE.md`](tests/corpus/PROVENANCE.md)). The
PDF-converter tests are dependency-agnostic: they pass whether or not Word/LibreOffice is installed
(an unavailable converter is itself a tested, expected outcome, not a skip).

### Publish a self-contained executable

```pwsh
dotnet publish src/BcWordLayout.McpHost -c Release -r win-x64 --self-contained
```

Produces `BcWordLayout.McpHost.exe` (~160 KB) in
`src/BcWordLayout.McpHost/bin/Release/net10.0/win-x64/publish/` alongside the rest of a
self-contained deployment (238 files, ~109 MB — the .NET 10 runtime plus every dependency, so the
target machine needs no .NET installed; native debug symbols are stripped at publish). **Ship the
whole `publish/` folder together**, not just the `.exe` — it depends on the sibling `.dll`s.
`scripts/build-release-zips.ps1` builds the exact per-RID zips (+ checksums) a GitHub Release
carries, and `python tools/release/verify_server.py -- <command>` smoke-tests any install channel
(handshake + `tools/list`).

### Pack the NuGet packages

```pwsh
dotnet pack src/BcWordLayout.McpHost -c Release
```

Produces the top-level `BcWordLayout.Mcp` package (with the embedded `.mcp/server.json` manifest
NuGet.org renders as install config) plus one framework-dependent package per Windows RID — all of
them must be pushed together on release. `PackagingMetadataTests` pins the manifests' versions to
`Directory.Build.props`, so a version bump that misses one file fails the suite.

### Manual smoke test

The server speaks JSON-RPC over stdio (newline-delimited). Logs go to **stderr**; the protocol uses
**stdout**.

```pwsh
dotnet run --project src/BcWordLayout.McpHost
```

Then send, one JSON object per line:

```json
{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"smoke","version":"1.0"}}}
{"jsonrpc":"2.0","method":"notifications/initialized"}
{"jsonrpc":"2.0","id":2,"method":"tools/list","params":{}}
```

Or write the requests to a file of your own and pipe it (`Get-Content requests.jsonl | dotnet run
--project src/BcWordLayout.McpHost`), and try a real call:

```json
{"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"get_layout_info","arguments":{"layoutPath":"C:\\path\\to\\bc-word-layout-mcp\\tests\\corpus\\SalesInvoiceForSubscriptionBilling.docx"}}}
```

## Project layout

| Project | Role |
|---|---|
| `src/BcWordLayout.Domain` | Core: dataset models, `SchemaProvider` (schema parsing), `LayoutReader` (control inventory), `LayoutValidator` (quick validation), `LayoutEditor`/`SdtFactory`/`Location`+`LocationResolver` (edits), `LayoutBuilder` (`create_layout`), `LayoutRefresher` (`refresh_xml_part`). |
| `src/BcWordLayout.Merge` | `SampleDataGenerator` (seeded, type-aware fakes), `MergeEngine` (binding fill, repeater expansion incl. nested/re-anchored XPaths, picture placeholders), `FullValidator` (the dry-run merge behind `validate_layout level=full`). |
| `src/BcWordLayout.Render` | `IPdfConverter` + `WordComConverter` (primary) / `LibreOfficeConverter` (fallback) + `PdfConverterFactory` (auto-selection) + PDF output sanity checks. |
| `src/BcWordLayout.McpHost` | The MCP stdio server (console app): the 28 tool definitions (`Tools/ReadTools.cs`, `Tools/EditTools.cs`, `Tools/TableTools.cs`, `Tools/LifecycleTools.cs`, sharing `Tools/ToolGuards.cs`) and the `{ok,data,error}` envelope (`ToolContracts.cs`). |
| `tests/BcWordLayout.Tests` | xUnit tests: the 10 real corpus layouts in `tests/corpus/` plus synthetic fixtures — reader/validator/merge/editor/refresher/converter coverage, snapshot tests, MCP host tool tests, packaging-manifest guards, and the fidelity harness. |

## Contributing and security

- [`CONTRIBUTING.md`](CONTRIBUTING.md) — dev setup, the design constraints PRs need to fit
  ([`docs/adr/`](docs/adr/)), corpus rules, PR checklist.
- [`SECURITY.md`](SECURITY.md) — private vulnerability reporting and the threat model.
- [`CODE_OF_CONDUCT.md`](CODE_OF_CONDUCT.md).

## Further reading

- [`skills/al-word-layout/SKILL.md`](skills/al-word-layout/SKILL.md) — the companion skill: intended
  agent workflow, full tool reference, v1 supported matrix, anti-patterns, and error handling.
- [`skills/al-word-layout-design/SKILL.md`](skills/al-word-layout-design/SKILL.md) — the design
  skill: the two archetype skeletons (trading document, list report), per-element BC conventions
  (address blocks, caption placement, line-table alignment, totals, chrome, typography), and the
  stock idioms the tools cannot build.
- [`skills/al-word-layout-from-design/SKILL.md`](skills/al-word-layout-from-design/SKILL.md) — the
  design-to-layout skill: inputs to request, the sample-text-to-field map and its rules, the binding
  recipe for a styled design, mock-render traps, and the consultant handover.
- [`skills/al-word-layout-composite/SKILL.md`](skills/al-word-layout-composite/SKILL.md) — the
  composable-layout skill: body / header-footer / theme rules, the theme–body style contract, splitting
  one design into the three parts, and what made a theme fail to apply.
- [`docs/SOLUTION-DESIGN.md`](docs/SOLUTION-DESIGN.md) — architecture, tool surface, design risks.
- [GitHub issues](https://github.com/TKapitan/bc-word-layout-mcp/issues) — the backlog: what is
  deferred, why, and what each remaining item is blocked on. Labels classify each issue by type,
  impact, and blockage — the taxonomy is defined in
  [`docs/BRANCHING.md`](docs/BRANCHING.md#issue-classification).
- [`docs/FIDELITY-CHECKLIST.md`](docs/FIDELITY-CHECKLIST.md) — per-release mock-preview fidelity
  regression procedure, and the record of the BC-sandbox validation.
- [`CHANGELOG.md`](CHANGELOG.md) — release notes.

## License

[MIT](LICENSE) © Tom Kapitan. Third-party components and the corpus layouts' provenance:
[`THIRD-PARTY-NOTICES.md`](THIRD-PARTY-NOTICES.md).
