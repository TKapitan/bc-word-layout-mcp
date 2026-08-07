# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- **New tool `insert_page_number`** (GitHub issue #29): emits the Word `PAGE`/`NUMPAGES` field-code
  construct every stock BC document header carries — by default the full "X / Y" shape (`PAGE`
  field, literal `" / "`, `NUMPAGES` field), or the bare `PAGE` field with `includeTotal=false` —
  exactly as captured in the four corpus layouts that carry it (StandardSalesQuote,
  StandardPurchaseOrder, StandardSalesInvoiceVatSpec, SalespersonCommission), including the
  instruction spacing and the `w:noProof` cached-result run. Field codes are plain runs (no content
  control), so like `insert_text` the response's `controlId` is `0`. Supports the same locations,
  header/footer targeting, on-demand part scaffolding, and optional `bold`/`fontSizePoints` as
  `insert_text`. The stock idiom's leading caption stays a composed `insert_label` (`Page_Lbl`) +
  `insert_text` separator, so it remains a translatable dataset binding.

## [1.0.0] - 2026-08-03

Initial release — the first stable version. The only earlier tags are the `v1.0.0-rc.1`/`v1.0.0-rc.2`
pre-releases from the release-candidate phase; everything below is new, and there is no prior stable
release to compare against.

An MCP server that lets an agent create, edit, validate and preview Microsoft Word report layouts for
Business Central through deterministic, typed tools instead of free-hand OOXML editing.

### Added

- **23 MCP tools** over stdio, in four groups — see [`README.md`](README.md) for the full reference:
  - *Inspect* — `get_layout_info`, `list_dataset_fields`.
  - *Validate* — `validate_layout` (`quick` structural/binding checks, or `full` with a dry-run merge).
    The `quick` level includes a `table-style-resolves` warning: any `w:tblStyle` whose style id the
    layout's own styles part does not define is flagged — a dangling reference renders fine but
    silently does nothing, which is exactly the trap when a `tableStyle` parameter or hand-authored
    reference is misspelled or the layout carries no styles part at all.
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
- **Two companion skills**:
  - [`al-word-layout`](skills/al-word-layout/SKILL.md) — the intended workflow (inspect → edit →
    validate → preview), the supported matrix, and the anti-patterns.
  - [`al-word-layout-design`](skills/al-word-layout-design/SKILL.md)
    ([#27](https://github.com/TKapitan/bc-word-layout-mcp/issues/27)) — what a from-scratch BC Word
    layout should *look like*, closing the one quality dimension validation cannot judge: a
    structurally valid, fully bound, ugly document passes `validate_layout level=full` with zero
    findings. Two archetype skeletons (trading document, list/analysis report) mapped step-by-step to
    tool calls, per-element conventions (address blocks, caption placement, line-table alignment and
    borders, totals, header/footer chrome, typography, twip widths), and the stock idioms the tools
    cannot build, with the supported route for each. Every convention is tagged with its evidence —
    observed in a named stock corpus layout, or BC-verified in a from-scratch sandbox build — never
    invented; a sandbox round or corpus addition that refutes a convention updates the skill.
- **The schema-source rule, documented as the workflow entry point** — skill §1 ("Where the schema
  comes from") plus a matching anti-pattern, a README scope bullet, SOLUTION-DESIGN §6.8, and
  [ADR-0006](docs/adr/0006-schema-transplanted-never-synthesized.md): the server transplants the BC
  dataset part byte-for-byte from a BC-produced artifact and never synthesizes it from AL source, so
  a brand-new layout starts with one AL build (the compiler creates the referenced `.docx`, dataset
  part included) or one stock-layout/schema export from BC.
- **Three install channels**: the `BcWordLayout.Mcp` NuGet MCP-server package (top-level manifest
  plus per-RID `win-x64`/`win-arm64` tool packages) launched on demand via `dnx`; self-contained
  GitHub-Release zips with a checksum-verifying `install.ps1` that unpacks to a stable path; and a
  plugin installing the server and both companion skills in one step — in **both Claude Code
  and VS Code / GitHub Copilot** (this repository is a plugin marketplace for each; the skills
  themselves are open-standard `SKILL.md` files both consume natively).

### Fixed

Found and fixed during the release-candidate phase (the 2026-08-01 and 2026-08-02 BC-sandbox
comparison rounds against the `v1.0.0-rc.*` builds) — no stable release ever shipped these:

- **The mock preview no longer loses repeated table header rows (`w:tblHeader`)**
  ([#19](https://github.com/TKapitan/bc-word-layout-mcp/issues/19)). `preview_layout`'s render copy
  used to keep every repeater row's content-control shell — binding-stripped, but still a row-level
  `w:sdt` — and Word fragments a table at those shells: the header row ended up alone in a one-row
  table fragment that can never break across a page, so the header repetition BC renders on every
  page of a multi-page table silently never triggered in the mock (found in the 2026-08-02 sandbox
  round 2, item P07; root cause proven by a Word COM probe — see the issue). The flatten-for-render
  step now unwraps row-level shells entirely, so cloned data rows become plain `w:tr` siblings of
  the header row and Word keeps the table whole, repeating the header exactly as BC does. Row-level
  only: inline field/label shells still flatten exactly as before, and neither the user's layout nor
  the default logical merge is touched. A table left rowless by a repeater that matched zero data
  rows (possible with a real `dataOverridesPath` dataset) is removed from the render copy outright.

- **`preview_layout`'s `dataOverridesPath` now accepts what BC actually exports**
  ([#4](https://github.com/TKapitan/bc-word-layout-mcp/issues/4)). The parameter was documented as
  taking "a real exported BC dataset XML", but only the layout's own data-store part shape
  (`NavWordReportXmlPart`) loaded — the report UI's *Send to → XML* export (a namespace-less
  `ReportDataSet` document) was refused, so the documented workflow failed at the first step. Both
  shapes are now accepted, sniffed by root element; the export is converted internally into the
  layout's shape (`ReportDataSetConverter` — the in-product version of the bridge that previously
  lived only in `tools/e2e/bc_compare.py`, which now passes exports straight through). Beyond the
  bridge, the conversion also applies each column's `decimalformatter` (raw `100` → `100.00` as BC
  itself renders) using the culture named by the export root's `formatRegion` attribute (fallback
  `language`, then invariant), strictly per column — pre-formatted columns are copied verbatim.
  Feeding an export from a different report than the layout's dataset now fails with an error naming
  both report ids instead of producing a preview where every binding is silently unresolved. This
  closes the recurring `decimalformatter` difference recorded by both BC sandbox comparison rounds
  (2026-08-01 / 2026-08-02, `docs/FIDELITY-CHECKLIST.md`).

- **A blank `create_layout` build now pins its typography**
  ([#3](https://github.com/TKapitan/bc-word-layout-mcp/issues/3)). A from-scratch layout used to ship
  no `word/styles.xml` and no theme, so nothing in the file named a typeface — Word rendered its
  application default and Business Central rendered a different one (observed in a real BC sandbox,
  2026-08-01). Blank builds now scaffold a default styles part: Calibri 11 pt named explicitly in
  `docDefaults` (the typeface every stock corpus layout resolves to), plus Word's four stock default
  styles and a `TableGrid` definition — so `insert_repeater_table`'s documented `tableStyle='TableGrid'`
  example resolves instead of referencing a style that does not exist. Existing layouts and
  `templatePath` builds are untouched: a template keeps its own styles/theme (or deliberately neither),
  and the scaffold is never retrofitted onto a pre-existing document. The fix is BC-verified: all three
  from-scratch layouts in the 2026-08-02 sandbox round rendered in the same typeface and sizes as the
  mock.

### Known limitations

- **The preview is a MOCK.** It merges sample data offline and converts with Word or LibreOffice — never
  through the real Business Central report engine. It is good enough to catch structural and binding
  mistakes on every edit; it is not a substitute for a real sandbox render, which remains the sign-off
  step. See [`docs/FIDELITY-CHECKLIST.md`](docs/FIDELITY-CHECKLIST.md) for what is guarded automatically
  and what still needs a human eye.
- **BC-sandbox validation covers seven layouts on one BC version family.** An authoring pack built
  with these tools was validated against real BC sandboxes (BC 28.0 and 28.3, Cronus AU) across two
  rounds — five items in round 1, seven in round 2, all passed
  ([`docs/FIDELITY-CHECKLIST.md`](docs/FIDELITY-CHECKLIST.md) records both runs).
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
