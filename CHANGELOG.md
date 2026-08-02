# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2026-08-01

Initial release. Everything below is new; there is no prior public version to compare against.

An MCP server that lets an agent create, edit, validate and preview Microsoft Word report layouts for
Business Central through deterministic, typed tools instead of free-hand OOXML editing.

### Added

- **23 MCP tools** over stdio, in four groups — see [`README.md`](README.md) for the full reference:
  - *Inspect* — `get_layout_info`, `list_dataset_fields`.
  - *Validate* — `validate_layout` (`quick` structural/binding checks, or `full` with a dry-run merge).
  - *Preview* — `preview_layout` (sample-data merge + PDF via Word COM or LibreOffice),
    `render_preview_pages` (PDF pages returned inline as images so the agent can look at its own work).
  - *Author and edit* — `create_layout`, `insert_field`, `insert_label`, `insert_text`, `insert_picture`,
    `insert_table`, `insert_repeater_table`, `insert_repeater_row`, `remove_control`, `set_cell_text`,
    `clear_cell_text`, `set_cell_borders`, `set_column_widths`, `insert_column`, `remove_column`,
    `merge_cells`, `split_cells`, `refresh_xml_part`.
- **A uniform `{ok, data, error}` envelope.** Nothing throws across the MCP boundary; every failure carries
  a `code`, a `message`, and a `hint` naming the argument to fix or the inspection tool to call first.
- **Write safety on every mutating tool.** Each edit is written to a staged copy, validated with
  `OpenXmlValidator` before saving, and rejected outright if it would introduce a new structural error —
  the file on disk is left untouched. Table edits additionally pass a grid-consistency guard that catches
  rows desyncing from their `w:tblGrid`, a corruption class `OpenXmlValidator` accepts silently.
- **Support for the shapes real BC layouts actually have**, rather than a tidy subset: pervasively
  `gridSpan`-spanned and ragged tables addressed by grid column, rows with skipped grid columns
  (`w:gridBefore`/`w:gridAfter`), repeater nesting several levels deep with per-level XPath re-anchoring,
  both the legacy `w:dataBinding` and the `w15:dataBinding` field-control shapes, layouts whose dataset
  part is UTF-8 or UTF-16 and may carry no store item at all, and multi-section documents where "the
  header" means the first section's default header rather than the first part in the package.
- **Label classification that handles both real-world shapes out of the box** — columns named
  `*Lbl`/`*_Lbl` (BC's documented convention) plus every direct column of a data item named `Labels`
  (the dedicated labels data item common in older/converted reports; the rule is self-scoping, so
  layouts without such an item are unaffected). Overridable per host via `BCWL_LABEL_SUFFIXES` and
  `BCWL_LABELS_DATA_ITEM` (a different data-item name retargets the rule; `-` disables it).
- **A companion skill** ([`skills/al-word-layout`](skills/al-word-layout/SKILL.md)) documenting the
  intended workflow, the supported matrix, and the anti-patterns.
- **Three install channels**: the `BcWordLayout.Mcp` NuGet MCP-server package (top-level manifest
  plus per-RID `win-x64`/`win-arm64` tool packages) launched on demand via `dnx`; self-contained
  GitHub-Release zips with a checksum-verifying `install.ps1` that unpacks to a stable path; and a
  plugin installing the server and the `al-word-layout` skill in one step — in **both Claude Code
  and VS Code / GitHub Copilot** (this repository is a plugin marketplace for each; the skill itself
  is an open-standard `SKILL.md` both consume natively).

### Known limitations

- **The preview is a MOCK.** It merges sample data offline and converts with Word or LibreOffice — never
  through the real Business Central report engine. It is good enough to catch structural and binding
  mistakes on every edit; it is not a substitute for a real sandbox render, which remains the sign-off
  step. See [`docs/FIDELITY-CHECKLIST.md`](docs/FIDELITY-CHECKLIST.md) for what is guarded automatically
  and what still needs a human eye.
- **BC-sandbox validation covers five layouts on one BC version family.** A five-layout authoring
  pack built with these tools was validated against real BC sandboxes (BC 28.0 and 28.3, Cronus AU)
  — all five passed ([`docs/FIDELITY-CHECKLIST.md`](docs/FIDELITY-CHECKLIST.md) records the run).
  That establishes that BC accepts and renders what the tools author; it does not make the mock
  preview authoritative, and other BC versions/localizations carry no validation yet.
- **The LibreOffice conversion path is unverified end to end**
  ([#7](https://github.com/TKapitan/bc-word-layout-mcp/issues/7)); the Word COM path is the
  tested one.
- **BC add-in conditional controls are not supported** — *Hide Field if Zero*, *Hide Empty Table*, *Hide
  Empty Table Row*, layout comments ([#8](https://github.com/TKapitan/bc-word-layout-mcp/issues/8)).
- **Tables using vertical merges (`w:vMerge`) are refused** by every column operation
  ([#9](https://github.com/TKapitan/bc-word-layout-mcp/issues/9)).
- **Creating a repeater in a header or footer is refused**; reading and merging a pre-existing one works,
  with a warning ([#10](https://github.com/TKapitan/bc-word-layout-mcp/issues/10)).
- **Preview captions match only when you supply real data** — with generated sample data, labels come
  from the dataset's own label columns or a humanized element name, so they read differently from a real
  BC render; pass a real exported dataset via `dataOverridesPath` and they match. Reading captions from
  the AL project's XLF files is not planned.

See the [issue tracker](https://github.com/TKapitan/bc-word-layout-mcp/issues) for the full list of
what is deferred and what each item is blocked on.
