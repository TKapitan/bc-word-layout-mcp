# bc-word-layout-mcp

[![CI](https://github.com/TKapitan/bc-word-layout-mcp/actions/workflows/ci.yml/badge.svg?branch=main)](https://github.com/TKapitan/bc-word-layout-mcp/actions/workflows/ci.yml)
[![GitHub release](https://img.shields.io/github/v/release/TKapitan/bc-word-layout-mcp)](https://github.com/TKapitan/bc-word-layout-mcp/releases/latest)
[![NuGet](https://img.shields.io/nuget/v/BcWordLayout.Mcp)](https://www.nuget.org/packages/BcWordLayout.Mcp)
[![NuGet downloads](https://img.shields.io/nuget/dt/BcWordLayout.Mcp)](https://www.nuget.org/packages/BcWordLayout.Mcp)
[![.NET 10](https://img.shields.io/badge/.NET-10-512BD4)](https://dotnet.microsoft.com/download/dotnet/10.0)
[![License: MIT](https://img.shields.io/github/license/TKapitan/bc-word-layout-mcp)](LICENSE)

An MCP (Model Context Protocol) server that lets AI agents **create, edit, validate, and preview
Microsoft Word report layouts for Microsoft Dynamics 365 Business Central** through deterministic,
typed tools, instead of editing the OOXML by hand, where one misplaced element silently corrupts the
document.

- **30 typed MCP tools** over stdio: inspect a layout and its dataset, author fields, labels, pictures,
  repeater tables and table structure, format text and cells, and refresh a layout after its dataset
  changes.
- **Every edit is validated before it is saved**, and refused with the file untouched if it would
  corrupt the layout.
- **An offline mock PDF preview** the agent can look at, and a dry-run merge that finds unresolved
  bindings.
- **One response shape** for every tool, with an actionable hint on every error.

## Install

The server reads and writes layout files **with your permissions, at the paths the agent gives it**.
Point it only at repositories and layouts you trust, as you would any local build tool.

### Plugin: the server and its skills (recommended)

The plugin installs the server together with four skills that teach an agent how to use it (see
[Skills](#skills)). This repository is a plugin marketplace:

- **Claude Code:** `/plugin marketplace add TKapitan/bc-word-layout-mcp`, then install the
  `bc-word-layout` plugin.
- **VS Code / GitHub Copilot** (agent plugins, preview; Copilot CLI too): add
  `"chat.plugins.marketplaces": ["TKapitan/bc-word-layout-mcp"]` to your user `settings.json`, then
  install the `bc-word-layout` plugin.

### NuGet package via `dnx` (needs the .NET 10 SDK)

[`dnx`](https://learn.microsoft.com/dotnet/core/tools/dnx), part of the .NET 10 SDK, fetches and runs
the server on demand.

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

**Claude Code** (`.mcp.json` in a project):

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

or for your user: `claude mcp add bc-word-layout -- dnx BcWordLayout.Mcp --yes`.

Without a version, `dnx` runs the latest stable release. Pin one with `BcWordLayout.Mcp@1.2.3`; a
pre-release runs only when pinned. `dotnet tool install -g BcWordLayout.Mcp` installs the same server
as the `bc-word-layout-mcp` command instead.

### Self-contained Windows binaries (no .NET needed)

Each [GitHub Release](https://github.com/TKapitan/bc-word-layout-mcp/releases) has `win-x64` and
`win-arm64` zips and an installer script that unpacks the latest one to
`%LOCALAPPDATA%\bc-word-layout-mcp\current\`, verifies its checksum, and prints the config block to
paste:

```pwsh
irm https://github.com/TKapitan/bc-word-layout-mcp/releases/latest/download/install.ps1 | iex
```

For a specific version, run that release's own script with `-Version`:

```pwsh
& ([scriptblock]::Create((irm https://github.com/TKapitan/bc-word-layout-mcp/releases/download/v1.2.3/install.ps1))) -Version 1.2.3
```

To run the server from a clone of this repository, see [`CONTRIBUTING.md`](CONTRIBUTING.md).

## Skills

The server works on its own, but agents use it noticeably better with its skills. The plugin installs
them; without it, copy the folders under `skills/` into your project's `.claude/skills/` or
`.github/skills/` (or your user-level `~/.claude/skills/`).

- [`al-word-layout`](skills/al-word-layout/SKILL.md): the intended workflow (inspect → edit → validate
  → preview → sandbox verify), the full tool reference, and the error codes.
- [`al-word-layout-design`](skills/al-word-layout-design/SKILL.md): what a new layout should look like,
  from Business Central's own document conventions.
- [`al-word-layout-from-design`](skills/al-word-layout-from-design/SKILL.md): from a finished Word design
  and the report's empty layout to a bound layout.
- [`al-word-layout-composite`](skills/al-word-layout-composite/SKILL.md): composable layouts, meaning the
  report body, a reusable header/footer layout, and a report theme.

## Requirements

- **Windows (x64 or arm64).** Linux and macOS are not supported yet.
- **Microsoft Word is optional.** `preview_layout` uses it to turn the merged mock into a PDF. Without
  it, the preview still returns the merged `.docx` and says why there is no PDF; every other tool works
  without Word.

## Configuration

### Target Business Central version

Some layout features need a newer Business Central. Each call works out the **target**, the minimum BC
version the layout must work on; the first match wins:

1. **The nearest AL `app.json` above the layout** (above `outputPath` for `create_layout`): its
   `application`, e.g. `29.0.0.0`. The search stops at a folder containing `.git`.
2. **`BCWL_BC_VERSION`** (e.g. `29` or `29.1`) in the server's `env` block, for layouts outside an AL
   project.
3. Otherwise the target is **unknown**.

`get_layout_info` and `validate_layout` report the target they used. A tool that would write a feature
from the table below refuses with `bc_version_unsupported` when the target is lower or unknown, and
`validate_layout` reports such a feature a layout already has. The gated features are:
- theme colours such as `accent1` (in `format.color`, `fill` and `borderColor`);
- `paragraphStyle`;
- `set_layout_theme`.

Hex colours are never gated.

| Feature | Kind | Since | Covers | Evidence |
|---|---|---|---|---|
| `report-themes` | capability | 29.0 | report theme `.dotx` files, header/footer layouts bound to `Tenant_Report_Defaults` | 2026 release wave 2; a composite body + header/footer + theme set printed on 2026-10-17 |
| `themable-formatting` | capability | 29.0 | theme-slot colours for text, cell fills and borders, paragraph styles, and the layout's own theme colours and fonts | decided with report themes, 2026-10-27 (#152) |

### Label columns

The tools treat a dataset column as a label when its name ends in `Lbl`/`_Lbl`, and every column of a
data item named `Labels`. Two environment variables change that:

- `BCWL_LABEL_SUFFIXES`: comma-separated label suffixes, e.g. `Lbl,Caption` (default `Lbl`).
- `BCWL_LABELS_DATA_ITEM`: the labels data item's name, or `-` to turn that rule off (default `Labels`).

```json
{
  "mcpServers": {
    "bc-word-layout": {
      "command": "dnx",
      "args": ["BcWordLayout.Mcp", "--yes"],
      "env": { "BCWL_BC_VERSION": "29", "BCWL_LABEL_SUFFIXES": "Lbl,Caption" }
    }
  }
}
```

## Tools

| Tool | Purpose |
|---|---|
| `get_layout_info` | Inspect a layout: report identity, every content control, header/footer parts, every table's rows and cells (with the coordinates the cell and column tools take), the target BC version, and its look (theme, colours, styles). |
| `list_dataset_fields` | List the dataset (data items and columns) from a layout or a schema `.xml`, marking labels and, for a layout, which columns are bound. |
| `validate_layout` | Check structure and bindings (`level=quick`), or also dry-run merge sample data so every unresolved binding becomes a finding (`level=full`). |
| `preview_layout` | Merge sample (or real) data into a copy of the layout and convert it to PDF: a mock preview. |
| `render_preview_pages` | Return pages of a preview PDF as images the agent can look at, optionally also saved as `.png` files. |
| `create_layout` | Create a new layout from a report schema, blank or over a branded template; `theme: "office"` adds the stock Office theme. |
| `insert_field` | Insert a field bound to a dataset column, in the body, a header or a footer. |
| `insert_label` | Insert a label bound to a caption column. |
| `insert_text` | Insert static text, such as the separator between two fields. |
| `insert_page_number` | Insert "Page X / Y" (or just X). |
| `insert_page_conditional_text` | Insert text that prints on the last page only, every page but the last, or the first page only. |
| `insert_repeater_table` | Insert a line-items table that repeats per record, in the Business Central look by default. |
| `insert_repeater_row` | Add a detail row under each line of an existing repeater (assembly components, serial numbers). |
| `insert_picture` | Insert a picture bound to a dataset column, such as the company logo. |
| `insert_table` | Insert a plain table, in the body, a header or a footer: address blocks, caption/value grids, totals blocks. |
| `wrap_in_repeater` | Make existing content repeat per record, e.g. one page per customer in a statement. |
| `remove_control` | Remove a content control by its id, optionally keeping its content. |
| `set_cell_text` | Replace the plain text of a table cell. |
| `set_text_format` | Restyle text already in the layout: bold, italic, underline, size, font, colour, styles, alignment. |
| `set_layout_theme` | Change the layout's own theme colours and fonts; everything coloured through the theme follows. |
| `clear_cell_text` | Remove the plain text from a table cell. |
| `insert_subtotal_row` | Add a row after each group of a repeater, for subtotals or spacing. |
| `insert_table_row` | Insert a row that prints once, such as a totals row inside the lines table. |
| `set_column_widths` | Set a table's column widths. |
| `insert_column` | Add a column at any position, bound or plain; the table keeps its width. |
| `remove_column` | Remove a column. |
| `merge_cells` | Merge adjacent cells of one row. |
| `split_cells` | Split a merged cell back into single cells. |
| `set_cell_format` | Draw cell borders (rules) in any colour and fill cells. |
| `refresh_xml_part` | Update a layout to a changed report dataset, reporting bindings that no longer resolve and new fields. |

The [`al-word-layout`](skills/al-word-layout/SKILL.md) skill documents every parameter.

## Responses and errors

Every tool returns the same envelope:

```json
{ "ok": true, "data": { "...": "tool-specific payload" } }
```

On failure, `ok` is false and `error` carries `{ code, message, hint }`. The `hint` always names the next
step, such as which argument to fix or which tool to call first. Malformed calls (an unknown tool or
argument, a wrong type, a relative path) get the same envelope. `render_preview_pages` adds one image
per page after the envelope.

The 24 mutating tools (`insert_field`, `insert_label`, `insert_text`, `insert_page_number`,
`insert_page_conditional_text`, `insert_picture`, `insert_table`, `insert_repeater_table`,
`insert_repeater_row`, `insert_subtotal_row`, `insert_table_row`, `wrap_in_repeater`, `remove_control`,
`set_cell_text`, `set_text_format`, `set_layout_theme`, `clear_cell_text`, `set_column_widths`,
`insert_column`, `remove_column`, `merge_cells`, `split_cells`, `set_cell_format`, `refresh_xml_part`)
validate every edit before saving it:
- An edit that would corrupt the file or break a table is refused (`edit_would_corrupt`,
  `edit_would_break_table`), and the file on disk is left untouched.
- A successful edit returns a fresh `quickValidation` summary.
- Every save removes a link to an external Word template, which Business Central refuses to upload.

## Preview and Business Central

`preview_layout` merges **sample data** (or a real dataset you export from BC, via `dataOverridesPath`)
into a copy of the layout and converts it to PDF **offline**. It never runs the real Business Central
report engine.
- **What it is good for:** catching structural and binding mistakes on every edit.
- **What can differ from BC:** captions, fonts and exact pagination.
- **Data it can't reproduce:** sample data fills every field, and the preview ignores the Word add-in's
  hide-if-empty controls.

**Final sign-off on a layout is always a Business Central sandbox render.**

Layouts made with these tools have been printed in **BC 28** and **BC 29** sandboxes (Cronus AU): BC
accepts them, bindings resolve, repeaters expand, and the preview's structure matched the real render.
[`docs/FIDELITY-CHECKLIST.md`](docs/FIDELITY-CHECKLIST.md) records every round.

## Not supported

- **The Word add-in's conditional controls** (Hide Table Row If Empty, Hide Field If Zero, Hide Table
  Column If Empty, layout comments): no tool writes them yet. Their Business Central behaviour is
  measured in [#123](https://github.com/TKapitan/bc-word-layout-mcp/issues/123) and
  [#162](https://github.com/TKapitan/bc-word-layout-mcp/issues/162); until then, edit them by hand.
- **Repeaters in headers and footers.** Nothing shows that BC prints them, so the tools create
  repeaters only in the body.
- **Margins, page setup, and defining new styles.** Start from a branded template (`create_layout`'s
  `templatePath`), or edit in Word and run `validate_layout` afterwards.
- **Report theme `.dotx` files:** `set_layout_theme` changes a layout's own theme only. A report theme
  selected in BC replaces the layout's own theme when it prints.
- **Creating a vertical cell merge.** Tables that already have merges can be edited.
- **Generating the dataset schema.** It always comes from Business Central: a layout the AL compiler
  generated, an exported layout, or an exported schema `.xml`
  ([ADR-0006](docs/adr/0006-schema-transplanted-never-synthesized.md)). A new layout starts with one AL
  build or one export.
- **Translated captions in the preview.** Pass a real exported dataset for real captions.
- **RDL and Excel layouts, and uploading a layout to Business Central.**

## Public contract and versioning

[Semantic versioning](https://semver.org). The public API is the tool names, their parameters, the
`{ok, data, error}` envelope and the error codes. Removing or changing any of them happens only in a
major release; new tools, optional parameters, fields and error codes come in minor ones. Error
`message` and `hint` texts can change at any time, so never parse them.

- **Business Central:** validated on BC 28 and BC 29. Word report layouts are a long-stable format, so
  other versions are expected to work; reports are welcome.
- **MCP SDK:** an exact version of the official C# SDK, upgraded in minor releases, with any change in
  behaviour noted in the [changelog](CHANGELOG.md).

## More

- [`CHANGELOG.md`](CHANGELOG.md): release notes.
- [GitHub issues](https://github.com/TKapitan/bc-word-layout-mcp/issues): what is planned or deferred,
  and why.
- [`docs/SOLUTION-DESIGN.md`](docs/SOLUTION-DESIGN.md) and [`docs/adr/`](docs/adr/): architecture and
  design decisions.
- [`CONTRIBUTING.md`](CONTRIBUTING.md): building from source and contributing.
- [`SECURITY.md`](SECURITY.md): reporting a vulnerability.

## License

[MIT](LICENSE) © Tom Kapitan. Third-party components and the provenance of the corpus layouts:
[`THIRD-PARTY-NOTICES.md`](THIRD-PARTY-NOTICES.md).
