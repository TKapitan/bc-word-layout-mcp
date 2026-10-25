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
| Binding fill / field values | **[Automated]** | `MergeSnapshotTests` (merged main-part XML vs. approved snapshot, three corpus layouts) + `MergeEngineTests.Field_fill_matches_the_generators_exact_value` (exact value equality against the generator's own output). |
| Repeater expansion / nesting / re-anchoring | **[Automated]** | `MergeEngineTests` — `Repeater_expands_configured_row_count_with_reanchored_per_row_values`, `Nested_repeater_reanchors_across_two_levels_and_counts_multiply`, the `MaxRowsPerRepeater` row-cap tests — plus `MergeSnapshotTests` for the real corpus, and `W15BindingAndDeepNestingTests` for the deepest nesting the corpus contains (a 5-level chain in a single table). |
| Structural OOXML validity of edited/created/merged layouts | **[Automated]** | `OpenXmlValidator`, run in: `FidelityHarnessTests` (merged corpus docx), `LayoutValidatorTests`/`FullValidatorTests` (quick/full validation), `MergeEngineTests.Corpus_layouts_merge_and_validate_with_zero_errors`, and every mutating tool's own before/after structural gate (`ToolGuards.GuardMutate`). |
| Merge determinism | **[Automated]** | `MergeSnapshotTests.Merging_the_same_corpus_file_twice_produces_byte_identical_pretty_printed_xml`, `MergeEngineTests.Merging_the_same_corpus_file_twice_is_deterministic`. |
| Caption text | **[Manual, generated data only]** | With GENERATED sample data the preview fills labels from the dataset's own label columns, falling back to humanized element names (`Item No Line`), so captions read differently from a real render (`No.`). With a real exported dataset via `dataOverridesPath` they MATCH — the dataset carries the captions BC itself renders (verified against a BC sandbox, 2026-08-01). Reading captions from the AL project's XLF files is **not planned**: supply a real dataset when caption text matters. Generated data treats `Capt` (`JobTaskNoCapt`) as a caption word, and samples a percent formatted as text (`LineDiscountPercentText_Line`) as a short percent rather than a caption (both round 3). |
| Font | **[Manual]** | A layout that PINS its fonts renders the same either side: the stock corpus layouts resolve `minorHAnsi` to Calibri through their theme part, and their sandbox renders matched the mock (2026-08-01). A blank `create_layout` build now pins the same typography itself — its scaffolded `styles.xml` names Calibri 11 pt explicitly in `docDefaults` ([#3](https://github.com/TKapitan/bc-word-layout-mcp/issues/3); `DefaultStylesScaffold`), so the two-renderers-two-fonts disagreement is fixed at create time — **BC-verified 2026-08-02** across three from-scratch layouts (round 2, items P01/P06/P07: same typeface and sizes both sides, including a `tableStyle='TableGrid'` reference BC honoured from the scaffolded styles part and explicit `fontSizePoints`/`bold` runs over the pinned base). A layout that names NO font (externally authored, or a `templatePath` shell without a styles part) still renders each side's own default — check which case you are in before reading anything into a font difference. The rest of what the `format` object writes (italic, underline, colour, another font, a character style) and `set_cell_format`'s fills are **BC-verified 2026-10-24** (round 3, P13, [#15](https://github.com/TKapitan/bc-word-layout-mcp/issues/15)): they print as the mock shows them. A font Business Central online does not have prints in a substitute font, as in Word's mock; `validate_layout`'s `font-installed` warning names such fonts. |
| Pagination / page breaks / "different first page" headers | **[Manual]** | Word/LibreOffice's pagination engine is not BC's report engine. Measured once (2026-08-01, a 2-page purchase order): same page count, same first-page→default header/footer switch at page 2, page break within ONE row. Re-measured 2026-08-02 (a 6-page report): header/footer parts repeat identically both sides; repeated table header rows (`w:tblHeader`) were a mock-only gap — Word fragmented the table at the merged copy's flattened row-SDT wrappers, so the mock lost the every-page header repetition BC renders — **fixed** by unwrapping row-level sdt shells in the render flatten ([#19](https://github.com/TKapitan/bc-word-layout-mcp/issues/19), `MergeEngine.UnwrapRowLevelSdts`; the layout itself was always correct). Still manual — one document is not a guarantee. Note that BC honours the Word first/default/even part model exactly, so a layout with `w:titlePg` renders its FIRST-page parts on page 1. |
| Picture rendering (placeholder vs. real image) | **[Automated for Word path]** / **[Manual for LibreOffice + real image]** | `MergeEngineTests.Corpus_picture_controls_are_filled_with_a_valid_placeholder_png` asserts the placeholder PNG fill itself; `FidelityHarnessTests`/`PdfConverterTests` render it end-to-end via Word COM. The LibreOffice path is unverified — LibreOffice was never installed on the reference dev machine, so whether it accepts the placeholder image part's package-root/absolute relationship target is still open (see [#7](https://github.com/TKapitan/bc-word-layout-mcp/issues/7)). What a genuine (non-placeholder) picture looks like is manual either way — the mock always substitutes a placeholder, by design. |
| Number/date/locale formatting | **[Automated determinism]** / **[Manual vs BC locale]** | `SampleDataGeneratorTests` locks down the generator's own deterministic, type-aware output run over run. Whether it matches BC's locale-aware formatting is manual, and it will not: the sandbox rendered `31. März 2025` / `1.800,00` from its own regional settings. With a real exported dataset the values come from BC itself — including raw-number columns: each `decimalformatter` attribute the export carries is applied with the export's own `formatRegion` culture ([#4](https://github.com/TKapitan/bc-word-layout-mcp/issues/4), fixed; `ReportDataSetConverter`, guarded by `SampleDataGeneratorTests`' export-shape tests), so `100` renders `100.00` as BC prints it. One caveat stays manual: the separator GLYPHS come from `formatRegion`, while a BC user whose personal regional settings differ from the report's format region may see different ones. |
| Hide-if-empty conditional controls | **[Deferred]** | Not implemented in `MergeEngine`, and there is no `set_hide_if_empty` editing tool. Two corpus layouts carry the captured controls: SalespersonCommission (*Hide Empty Table Row*) and SubcontractorDispatchList (*Hide Empty Table*). Their mock previews ignore the wrapper and show the row or table regardless, so compare those two against a BC render with care. What BC treats as empty is the open question ([#123](https://github.com/TKapitan/bc-word-layout-mcp/issues/123)): round 3 could not answer it, because report 115 prints only salespeople with entries, so the pack's P09/P10 now wrap the stock quote's caption row, which is empty for a quote with no lines; *Hide Field if Zero* and layout comments are uncaptured ([#8](https://github.com/TKapitan/bc-word-layout-mcp/issues/8)). |
| Report limits / large datasets | **[Manual]** | `MergeOptions.MaxRowsPerRepeater` (default 100) bounds the MOCK preview's own row count for robustness — a safeguard against the mock blowing up, not a statement about BC's real limits. Actual BC report performance/limits at production data volumes are sandbox-only. |

## The sandbox pack (the prepared version of the manual procedure)

`python tools/e2e/sandbox_pack.py` builds seventeen layouts into `preview-output/sandbox-pack/` (gitignored)
for uploading to a real BC sandbox: three authored entirely by the tools (P01, P06, P07 — every one a
`create_layout` blank build, so all three carry the scaffolded default typography whose BC-side
verification rides this pack), four exercising a distinct editing capability against stock
base-app layouts (P02–P05), four research probes (P08–P11) that answer open questions blocked on BC
behaviour rather than check a capability — storeItemIDs naming no part
([#102](https://github.com/TKapitan/bc-word-layout-mcp/issues/102)), the add-in's hide-if-empty semantics
([#123](https://github.com/TKapitan/bc-word-layout-mcp/issues/123)) and last-page-only footer content
([#11](https://github.com/TKapitan/bc-word-layout-mcp/issues/11)) — and six capability checks of what no
round has printed from the tools:

- **P12**, a statement repeated per customer by `wrap_in_repeater`
  ([#31](https://github.com/TKapitan/bc-word-layout-mcp/issues/31)) — no tool-written block repeater,
  `w:group` included, has been rendered by BC before;
- **P13**, the formatting tools ([#15](https://github.com/TKapitan/bc-word-layout-mcp/issues/15)):
  direct italic (a shape no stock layout carries), colour, underline, a character style, a font that
  exists nowhere, a cell fill, and restyled stock content;
- **P14**, the reworked table tools
  ([#131](https://github.com/TKapitan/bc-word-layout-mcp/issues/131)–[#133](https://github.com/TKapitan/bc-word-layout-mcp/issues/133)):
  tables in a header and a footer, `borders="headerRule"` on a plain table and `"grid"` on a totals
  block, a rule under every repeated line, a column that takes what the others leave, and the work
  description repeated per line (the nested-item block repeater);
- **P15**, the totals rows every stock document keeps inside its lines table (`insert_table_row`) and the
  stock "Page X / Y" field codes (`insert_page_number`), shapes copied from Microsoft's layouts that
  no round has printed from tool output;
- **P16**, the stock group-subtotal rows (`insert_subtotal_row`): a spacer and a bold `Subtotals` row
  after every salesperson's entries, then a grand-total row;
- **P17**, `create_layout` over a branded `.dotx` template
  ([#77](https://github.com/TKapitan/bc-word-layout-mcp/issues/77)), made from the stock quote: its
  headers, footers and styles kept, the document number added to both of its headers.

Each item folder carries the `.docx` to upload, this
tool's own mock render of it, a README naming the questions that item's BC render answers, and the full
tool-call/validation record; the pack root carries `INSTRUCTIONS.md` (BC-side procedure) and
`COMPARISON.md` (a per-item sheet keyed to the dimensions table above).

`python tools/e2e/bc_compare.py` then closes the loop: it re-runs the mock on **the data BC itself
used** (each BC dataset export, passed straight to `dataOverridesPath` — accepted as-is since
[#4](https://github.com/TKapitan/bc-word-layout-mcp/issues/4)), rendering PNGs beside the BC ones. Without that step a side-by-side conflates two questions — "does the layout render the
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
and amounts where BC's export carries a `decimalformatter` nothing applied at the time
([#4](https://github.com/TKapitan/bc-word-layout-mcp/issues/4) — since fixed: the formatter is now
applied with the export's `formatRegion` culture).

Two results worth keeping in mind when reading a finding from this tool:

- **`validate_layout`'s Error findings correspond to real BC upload rejections.** A stock corpus layout
  carrying a stray control bound to a foreign store item was reported as 3 errors here and refused by BC
  with `UnknownDataBinding`; removing the control (the repair the finding text prescribes) made BC accept
  it.
- **BC follows Word's header/footer part model exactly** — first-page parts on page 1, default parts from
  page 2 — so a layout with `w:titlePg` will not show anything inserted into its default header on a
  single-page document ([#5](https://github.com/TKapitan/bc-word-layout-mcp/issues/5)).

### Recorded result — 2026-08-02, BC sandbox (round 2: seven items, typography focus)

Re-run after the issue-#3 fix (blank builds scaffold a default styles part), with the pack extended
from five to seven items — three of them from-scratch `create_layout` builds (P01, plus new P06/P07).
Every item's BC dataset export was captured, so `bc_compare.py` re-ran the mock on **the data BC
itself used** — every comparison below is layout-vs-layout, not data-vs-data. **All seven items
passed**; the filled comparison sheet lives with the round's pack output.

- **The issue-#3 typography fix is BC-verified.** All three from-scratch layouts rendered in the same
  typeface and sizes as the mock (Calibri 11 pt from the scaffolded `docDefaults`) — round 1's
  two-fonts disagreement is gone. That includes explicit `fontSizePoints`/`bold` static runs over the
  pinned base (P06) and text inside scaffolded header/footer parts and nested detail rows (P07).
- **BC honours a `w:tblStyle` reference against the scaffolded styles part**: P06's line table, styled
  only by `tableStyle='TableGrid'` (no direct border formatting), rendered the full ½ pt grid in BC —
  the style-reference mechanism itself is now verified, not just the border shapes the corpus uses.
- **The blank build's scaffolded header/footer parts render in BC** with bound data on every page of a
  6-page document (P07) — round 1 had only proved a stock layout's own parts.
- **An old-style report works from scratch**: P07 targets report 115 (dedicated `Labels` data item, no
  `*_Lbl` columns); labels bound from the `Labels` item carried their caption text, and a nested
  detail repeater built from nothing repeated per entry without drifting off the parent grid.
- The only recurring difference remained the **`decimalformatter` gap
  ([#4](https://github.com/TKapitan/bc-word-layout-mcp/issues/4) — since fixed)**: columns whose BC
  export carries raw numbers rendered unformatted in the mock (`1002060.00` vs BC's `1,002,060.00`);
  pre-formatted columns matched exactly. The formatter is now applied at load with the export's
  `formatRegion` culture, so the next round's comparison should show no such difference.
- **One new mock-fidelity gap, found in post-round review of P07's page 2** (maintainer-reported):
  BC repeats the repeater table's `w:tblHeader` header row on every page, the mock does not. The
  LAYOUT is correct — `insert_repeater_table` marks the header row and BC honours it; the loss is in
  the mock pipeline: the merged render copy keeps the (binding-stripped) row-level SDT wrappers, and
  Word fragments the table at them (its object model reports the header row as a one-row table,
  which can never span a page, so repetition never fires). Root cause proven by re-converting the
  same merged copy with the row SDTs unwrapped — the header then repeats exactly as in BC. Tracked
  as [#19](https://github.com/TKapitan/bc-word-layout-mcp/issues/19) and since fixed there:
  `MergeEngine`'s flatten-for-render step now unwraps row-level sdt shells entirely
  (`UnwrapRowLevelSdts`).

### Recorded result — 2026-10-24, BC sandbox (round 3: seventeen items on a BC 28 sandbox)

All seventeen items printed once two of them were repaired, and three research probes answered their
questions. Every capability check printed what the mock showed. The round found one tool gap, two mistakes
in the pack's own items, and three mock-data gaps, all fixed for 2.0.0-rc.2, whose re-run covers the items
they touched.

- **BC now refuses a layout that links an external attached template.** P02's upload failed with
  `UnsupportedExternalContentFields: The Word layout contains unsupported fields or external references:
  ExternalRelationship:attachedTemplate. Remove them and try again.` P02 starts from Microsoft's stock
  SalesInvoiceForSubscriptionBilling layout, whose settings link the Word template its developer built
  it from — as do three more stock captures — and rounds 1 and 2 uploaded the same bases with the same
  link, so the check is newer than they are (the environment is still on BC 28, not 29). Every tool that
  writes a layout now removes the link, `validate_layout` reports one as an error, and the pack unlinks
  the one item no tool saves (P11, edited by hand).
- **BC cannot print a layout whose main part is not `word/document.xml`.** P17 (`create_layout` over a
  `.dotx`) uploaded, then failed to print with BC's "Something went wrong" internal server error. The
  template conversion ([#77](https://github.com/TKapitan/bc-word-layout-mcp/issues/77)) used the OpenXML
  SDK's `ChangeDocumentType`, which re-creates the main part as `word/document2.xml`; Word, and so the mock,
  follows the package relationships to it. The conversion now retypes the main part in place, and the
  rebuilt P17 printed: the template's headers and footers on every page, its first-page header on page 1.
- **A binding whose storeItemID names no part prints** (P08,
  [#102](https://github.com/TKapitan/bc-word-layout-mcp/issues/102)). BC accepted Microsoft's own
  StandardSalesInvoiceVatSpec bindings of that kind and printed both with their data, so `validate_layout`
  reports the case as a warning now (an ID naming another part stays an error), and `refresh_xml_part`
  re-points such bindings at the BC part.
- **BC evaluates `IF PAGE = NUMPAGES` per page** (P11,
  [#11](https://github.com/TKapitan/bc-word-layout-mcp/issues/11)). On a two-page order the footer text
  printed on the last page only, and the header's page numbers counted correctly. A one-page order was not
  printed.
- **Hide-if-empty stays open** (P09, P10,
  [#123](https://github.com/TKapitan/bc-word-layout-mcp/issues/123)). Report 115 prints only salespeople
  with entries in the filter, so neither wrapper had an empty case to act on, and the two renders were
  identical. The re-run moves both wrappers to the stock quote's caption row, which is empty for a quote
  with no lines.
- **Capability checks printed as the mock showed:** a block repeater written by `wrap_in_repeater` (P12,
  the first BC render of one: each customer on its own page with its own data, no blank first page);
  direct italic, underline, colour, a character style, another installed font and a cell fill (P13);
  tables in a header and a footer, the three `borders` looks and a rule under every line (P14); totals rows
  inside a lines table and page numbers (P15); per-group subtotal rows that sum to the grand total (P16).
  P01–P03, P05 and P07 printed as in round 2. Not exercised: a quote with a work description (P14) and a
  quote long enough for two pages (P15).
- **A font Business Central online does not have prints in a substitute** (P13): a serif face, as in
  Word's mock. `validate_layout` now warns about such a font (`font-installed`), checked against
  Microsoft's list of the fonts BC online has.
- **Tool gap: a row the row tools build did not match the rows around it** (P04). The detail row
  `insert_repeater_row` rebuilt in Job Quote printed at 11 pt between Segoe UI 8 pt rows, and its values
  wrapped. The row tools now give each new cell the formatting of the row beside it, column by column, and
  `insert_repeater_row` takes `format`, which P04's emptied task item needs.
- **Two pack mistakes:** P04 laid its detail columns out in another order than the stock row and left one
  out, and P06 bound the vendor address an order leaves empty (`VendAddr1`–`6`) instead of the buy-from
  address the stock layout prints (`BuyFromAddr1`–`8`). Both are corrected; the mock could not show the
  second, because generated data fills every field.
- **Three mock-data gaps made tables far taller than BC's:** a percent formatted as text sampled as a
  caption that wrapped to five lines per row, `Capt`-suffixed captions sampled as codes, and name-echo
  text that repeated the data item's name. All three are fixed (see the *Caption text* row).
- **Report 115's group row is not a group total:** its own commission amounts are the FIRST entry's
  values, so the stock layout shows only code, name and commission % there, and group totals come from the
  report's `Subtotals` item (as P16 binds them).

## Manual procedure (per release, 3 corpus reports)

For each of these three corpus layouts — `SalesInvoiceForSubscriptionBilling.docx`, `InventoryOrderDetails.docx`,
`StandardStatement.docx` (`tests/corpus/`):

1. **Generate the mock preview.** Either run the harness with the repo-root output opted in —

   ```pwsh
   $env:BCWL_FIDELITY_OUTPUT_DIR = "$(Get-Location)\fidelity-output"
   dotnet test --filter FidelityHarnessTests
   ```

   (Deliverable 2 below — regenerates fresh artifacts under `fidelity-output/<LayoutName>/`) or call
   the `preview_layout` MCP tool directly against the same corpus file. Open `merged.docx` /
   `preview.pdf` (the harness's names; `preview_layout` returns its own files as `mergedDocxPath` /
   `pdfPath`).
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
5. Any newly discovered gap gets added to this document's dimensions table rather than only living in a
   ticket, so the next release's pass starts from an up-to-date list.

## Where the automated artifacts land

By default (plain `dotnet test`, CI, clean clones) the harness writes to
`%TEMP%/bcwl-fidelity-output/` — the suite is hermetic and never touches the repo tree. For the
manual pass, set `BCWL_FIDELITY_OUTPUT_DIR` to `<repo>/fidelity-output` as shown above (gitignored —
regenerated fresh on every run, never committed; a human-review artifact, not a build output to keep
clean between runs). One subfolder per layout either way:

```
fidelity-output/
  SalesInvoiceForSubscriptionBilling/
    merged.docx     — the merge-engine working copy (sample data filled, repeaters expanded)
    preview.pdf      — present only when a PDF converter (Word COM or LibreOffice) is available
    SUMMARY.md       — converter used/availability, merge stats, merge warnings, mock-render disclaimer
  InventoryOrderDetails/
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
- **`FidelityHarnessTests`** — batch-produces the fresh artifacts above for the same three corpus
  layouts and asserts the automatable invariants: the merge produced output, zero bindings are
  unresolved against the corpus, and the merged docx stays `OpenXmlValidator`-clean.
