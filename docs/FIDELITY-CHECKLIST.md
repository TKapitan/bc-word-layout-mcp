# Fidelity Regression Checklist

Companion to [SOLUTION-DESIGN.md](SOLUTION-DESIGN.md) §5 ("Mock preview pipeline") and §7 ("Testing
strategy"). This is the per-release
fidelity regression procedure for the mock preview pipeline (`preview_layout` /
`MergeEngine` / `IPdfConverter`): what is already guarded automatically by `dotnet test`, and what a
human still has to check by eye against a real Business Central sandbox render.

## Purpose

`preview_layout` merges deterministic sample data into a working copy of the layout and converts it to
PDF via Word COM or LibreOffice — entirely offline, never through the real BC report engine. That is
good enough to catch structural and binding mistakes on every edit, and that part **is** fully
automated (see the dimensions table below). It is not, and can never be, a pixel-perfect stand-in for
what Business Central itself renders.

**Core caveat — repeated because it is the one thing this whole document exists to protect against:
the preview is a MOCK. Final sign-off on any layout is always a real Business Central sandbox render,
never this tool's PDF alone.** `preview_layout`'s own tool response carries this same disclaimer
verbatim (`PreviewDisclaimer` in `src/BcWordLayout.McpHost/Tools/LifecycleTools.cs`), and the automated
harness below (Deliverable 2) repeats it in every artifact it writes.

## When to run it

- **Each release** of this MCP server — in particular whenever `MergeEngine`, `SampleDataGenerator`,
  `PlaceholderImage`, an `IPdfConverter` implementation, or the layout corpus itself changes.
- Ad hoc, whenever a fidelity gap is reported or suspected (e.g. a developer says the preview looked
  fine but the sandbox render did not).
- **Not** required per commit or per PR — `dotnet test` already covers the Automated rows below on
  every build; this document is the additional manual layer on top, run less often.

## Fidelity dimensions

| Dimension | Tag | Automated guard / what "manual" means |
|---|---|---|
| Binding fill / field values | **[Automated]** | `MergeSnapshotTests` (merged main-part XML vs. approved snapshot, all 3 corpus layouts) + `MergeEngineTests.Field_fill_matches_the_generators_exact_value` (exact value equality against the generator's own output). |
| Repeater expansion / nesting / re-anchoring | **[Automated]** | `MergeEngineTests` — `Repeater_expands_configured_row_count_with_reanchored_per_row_values`, `Nested_repeater_reanchors_across_two_levels_and_counts_multiply`, the `MaxRowsPerRepeater` row-cap tests — plus `MergeSnapshotTests` for the real corpus, and `W15BindingAndDeepNestingTests` for the deepest nesting the corpus contains (a 5-level chain in a single table). |
| Structural OOXML validity of edited/created/merged layouts | **[Automated]** | `OpenXmlValidator`, run in: `FidelityHarnessTests` (merged corpus docx), `LayoutValidatorTests`/`FullValidatorTests` (quick/full validation), `MergeEngineTests.Corpus_layouts_merge_and_validate_with_zero_errors`, and every mutating tool's own before/after structural gate (`ToolGuards.GuardMutate`). |
| Merge determinism | **[Automated]** | `MergeSnapshotTests.Merging_the_same_corpus_file_twice_produces_byte_identical_pretty_printed_xml`, `MergeEngineTests.Merging_the_same_corpus_file_twice_is_deterministic`. |
| Caption text | **[Manual, generated data only]** | With GENERATED sample data the preview fills labels from the dataset's own label columns, falling back to humanized element names (`Item No Line`), so captions read differently from a real render (`No.`). With a real exported dataset via `dataOverridesPath` they MATCH — the dataset carries the captions BC itself renders (verified against a BC sandbox, 2026-08-01). Reading captions from the AL project's XLF files is **not planned**: supply a real dataset when caption text matters. |
| Font | **[Manual]** | A layout that PINS its fonts renders the same either side: the stock corpus layouts resolve `minorHAnsi` to Calibri through their theme part, and their sandbox renders matched the mock (2026-08-01). A blank `create_layout` build now pins the same typography itself — its scaffolded `styles.xml` names Calibri 11 pt explicitly in `docDefaults` ([#3](https://github.com/TKapitan/bc-word-layout-mcp/issues/3); `DefaultStylesScaffold`), so the two-renderers-two-fonts disagreement recorded below is fixed at create time. BC-side confirmation of the scaffolded pin still rides the next sandbox-pack re-run ([#14](https://github.com/TKapitan/bc-word-layout-mcp/issues/14)). A layout that names NO font (externally authored, or a `templatePath` shell without a styles part) still renders each side's own default — check which case you are in before reading anything into a font difference. |
| Pagination / page breaks / "different first page" headers | **[Manual]** | Word/LibreOffice's pagination engine is not BC's report engine. Measured once (2026-08-01, a 2-page purchase order): same page count, same first-page→default header/footer switch at page 2, page break within ONE row. Still manual — one document is not a guarantee. Note that BC honours the Word first/default/even part model exactly, so a layout with `w:titlePg` renders its FIRST-page parts on page 1. |
| Picture rendering (placeholder vs. real image) | **[Automated for Word path]** / **[Manual for LibreOffice + real image]** | `MergeEngineTests.Corpus_picture_controls_are_filled_with_a_valid_placeholder_png` asserts the placeholder PNG fill itself; `FidelityHarnessTests`/`PdfConverterTests` render it end-to-end via Word COM. The LibreOffice path is unverified — LibreOffice was never installed on the reference dev machine, so whether it accepts the placeholder image part's package-root/absolute relationship target is still open (see [#7](https://github.com/TKapitan/bc-word-layout-mcp/issues/7)). What a genuine (non-placeholder) picture looks like is manual either way — the mock always substitutes a placeholder, by design. |
| Number/date/locale formatting | **[Automated determinism]** / **[Manual vs BC locale]** | `SampleDataGeneratorTests` locks down the generator's own deterministic, type-aware output run over run. Whether it matches BC's locale-aware formatting is manual, and it will not: the sandbox rendered `31. März 2025` / `1.800,00` from its own regional settings. With a real exported dataset the values come from BC itself — except that BC's export carries RAW numbers plus a `decimalformatter` attribute which nothing here applies yet ([#4](https://github.com/TKapitan/bc-word-layout-mcp/issues/4)), so amounts can still read `100` where BC prints `100.00`. |
| Hide-if-empty conditional controls | **[Deferred]** | Phase 0.2 (reverse-engineering the BC add-in's Hide Field if Zero / Hide Empty Table / Hide Empty Table Row / layout comment OOXML) was never completed, so this is not implemented in `MergeEngine` and there is no `set_hide_if_empty` editing tool ([#8](https://github.com/TKapitan/bc-word-layout-mcp/issues/8)). Nothing to check yet on the current corpus either — no reviewed layout uses these controls. |
| Report limits / large datasets | **[Manual]** | `MergeOptions.MaxRowsPerRepeater` (default 100) bounds the MOCK preview's own row count for robustness — a safeguard against the mock blowing up, not a statement about BC's real limits. Actual BC report performance/limits at production data volumes are sandbox-only. |

## The sandbox pack (the prepared version of the manual procedure)

`python tools/e2e/sandbox_pack.py` builds five layouts into `preview-output/sandbox-pack/` (gitignored)
for uploading to a real BC sandbox: one authored entirely by the tools, and four exercising a distinct
editing capability against stock base-app layouts. Each item folder carries the `.docx` to upload, this
tool's own mock render of it, a README naming the questions that item's BC render answers, and the full
tool-call/validation record; the pack root carries `INSTRUCTIONS.md` (BC-side procedure) and
`COMPARISON.md` (a per-item sheet keyed to the dimensions table above).

`python tools/e2e/bc_compare.py` then closes the loop: it converts each BC dataset export into the shape
`dataOverridesPath` accepts and re-runs the mock on **the data BC itself used**, rendering PNGs beside
the BC ones. Without that step a side-by-side conflates two questions — "does the layout render the
same" and "is the data the same" — and only the first one is ours.

Prefer this over the ad-hoc procedure below: it fixes *what* gets rendered and *what to look for*, so two
people running it produce comparable results, and items are graded against their own base layout so a
corpus layout's pre-existing defects are not read as regressions.

### Recorded result — 2026-08-01, BC sandbox (Cronus AU)

The first time anything this tool produced was rendered by the real BC report engine. **All five items
passed.** Business Central accepted and correctly rendered: a layout authored entirely by these tools
(logo, address block, label/value grid, line repeater, totals rule); a stock layout with a bound column
added, widths re-stated, a header swapped to static text and a rule applied; a stock layout with a column
removed and two header cells merged; a nested detail repeater rebuilt with `insert_repeater_row`; and
bound fields in header and footer parts, carrying data on every page of a 2-page document.

Compared on identical data, the mock and BC renders were **structurally identical** — same fields, values,
row counts, column geometry, rules and page count. What differed: the picture placeholder (by design),
the font on the one layout that names no font ([#3](https://github.com/TKapitan/bc-word-layout-mcp/issues/3)),
and amounts where BC's export carries a `decimalformatter` nothing applies yet
([#4](https://github.com/TKapitan/bc-word-layout-mcp/issues/4)).

Two results worth keeping in mind when reading a finding from this tool:

- **`validate_layout`'s Error findings correspond to real BC upload rejections.** A stock corpus layout
  carrying a stray control bound to a foreign store item was reported as 3 errors here and refused by BC
  with `UnknownDataBinding`; removing the control (the repair the finding text prescribes) made BC accept
  it.
- **BC follows Word's header/footer part model exactly** — first-page parts on page 1, default parts from
  page 2 — so a layout with `w:titlePg` will not show anything inserted into its default header on a
  single-page document ([#5](https://github.com/TKapitan/bc-word-layout-mcp/issues/5)).

## Manual procedure (per release, 3 corpus reports)

## Manual procedure (per release, 3 corpus reports)

For each of the 3 corpus layouts — `SalesInvoiceForSubscriptionBilling.docx`, `InventoryOrderDetails.docx`,
`StandardStatement.docx` (`tests/corpus/`):

1. **Generate the mock preview.** Either run the harness with the repo-root output opted in —

   ```pwsh
   $env:BCWL_FIDELITY_OUTPUT_DIR = "$(Get-Location)\fidelity-output"
   dotnet test --filter FidelityHarnessTests
   ```

   (Deliverable 2 below — regenerates fresh artifacts under `fidelity-output/<LayoutName>/`) or call
   the `preview_layout` MCP tool directly against the same corpus file. Open `merged.docx` /
   `preview.pdf`.
2. **Render the same report from a BC sandbox.** Publish/run the corresponding AL report against a real
   Business Central sandbox environment with comparable sample data, and export/print to PDF.
3. **Compare side by side**, dimension by dimension, using the table above as the checklist — the
   Automated rows should already be green (a failing automated test means don't even start the manual
   pass; fix that first), so the manual pass only needs to look at the **[Manual]** rows (and the
   LibreOffice/real-image half of the Automated-for-Word-path row, if LibreOffice is the converter in
   use).
4. **Record pass/fail per dimension** (and per report) wherever the team tracks release sign-off — this
   document intentionally does not prescribe a specific tracker; attach the sandbox PDFs and the
   `fidelity-output/` artifacts as evidence either way.
5. Any newly discovered gap gets added to this document's dimensions table (or to SOLUTION-DESIGN.md §5's
   "Known fidelity gaps" list) rather than only living in a ticket, so the next release's pass starts from
   an up-to-date list.

## Where the automated artifacts land

By default (plain `dotnet test`, CI, clean clones) the harness writes to
`%TEMP%/bcwl-fidelity-output/` — the suite is hermetic and never touches the repo tree. For the
manual pass, set `BCWL_FIDELITY_OUTPUT_DIR` to `<repo>/fidelity-output` as shown above (gitignored —
regenerated fresh on every run, never committed; a human-review artifact, not a build output to keep
clean between runs). One subfolder per layout either way:

```
fidelity-output/
  SalesInvoice/
    merged.docx     — the merge-engine working copy (sample data filled, repeaters expanded)
    preview.pdf      — present only when a PDF converter (Word COM or LibreOffice) is available
    SUMMARY.md       — converter used/availability, merge stats, merge warnings, mock-render disclaimer
  CustomerStatement/
    ...
  StandardStatement/
    ...
```

## Automated regression guard (what `dotnet test` already covers today)

These are the existing suites the checklist's **[Automated]** tags above point at — run them (or just
`dotnet test`) before starting the manual pass; a red build here means the manual pass is premature:

- **`MergeSnapshotTests`** — merged main-document-part XML vs. an approved snapshot per corpus layout,
  plus merge determinism.
- **`MergeEngineTests`** (corpus checks) — field fill correctness, repeater/nested-repeater expansion and
  re-anchoring, row-cap robustness, unresolved-binding handling, picture-placeholder fill, corpus-wide
  zero-errors/zero-unresolved checks, determinism.
- **`LayoutValidatorTests`** / **`FullValidatorTests`** — quick (structural + binding) and full (dry-run
  merge) validation across the corpus.
- **The create/edit round-trip tests** — `LayoutBuilderTests` (`create_layout`), `LayoutEditorTests`
  (`insert_field`/`insert_label`/`remove_control`), `RepeaterTableTests` (`insert_repeater_table`,
  including its own merge round-trip proof), `LayoutRefresherTests` (`refresh_xml_part`),
  `SdtFactoryTests`, `LocationResolverTests` — every one of these round-trips a real save/reopen through
  OpenXmlValidator and `LayoutValidator.Quick`, several against the real corpus.
- **`PdfConverterTests`** (`PdfConverterFactoryTests`, `PdfConverterContractTests`, `LibreOfficeCliTests`,
  `PdfFileValidationTests`) — dependency-agnostic converter contract tests; green whether or not Word/
  LibreOffice is installed on the machine running them.
- **`FidelityHarnessTests`** (this task, Deliverable 2) — batch-produces the fresh artifacts above for all
  3 corpus layouts and asserts the automatable invariants: the merge produced output, zero bindings are
  unresolved against the corpus, and the merged docx stays `OpenXmlValidator`-clean.
