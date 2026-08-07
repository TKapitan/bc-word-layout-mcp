---
name: al-word-layout
description: Use whenever inspecting, creating, editing, validating, or previewing a Business Central Word report layout (.docx) — BC Word layouts, report layouts, and their bound content controls (fields, labels, repeater tables) — through the bc-word-layout MCP tools, instead of hand-editing the underlying OOXML directly. Covers the inspect → edit → validate → preview workflow, the exact tool surface and parameters, what v1 does and does not support, and the error envelope every tool returns.
---

# al-word-layout

The `bc-word-layout` MCP server gives typed, deterministic tools for inspecting, creating, editing,
validating, and previewing Business Central Word report layouts, so you stop hand-authoring OOXML
(content controls, `w:dataBinding`, XPaths) for anything these tools already cover.

This skill is the **mechanics**: the tool surface, the workflow, and the error envelope. Its
sibling skill `al-word-layout-design` is the **design**: what a new layout should look like — the
archetype skeletons, the observed BC document conventions, and the BC-verified recipe for each
block. When creating a layout from scratch (or reshaping one), read that skill first and come back
here for how to make the calls.

## 1. Intended workflow

The server provides a fast **inner loop** that runs in seconds, entirely offline. It does **not**
replace the **outer loop** — that is unchanged and remains the only real sign-off.

```
inspect → edit → validate → preview → (repeat as needed)
                                            │
                                            ▼
        sandbox verify (OUTER, unchanged, outside this server):
        AL build → publish to dev sandbox → run the report
```

| Stage | Tools | Purpose |
|---|---|---|
| **Inspect** | `get_layout_info`, `list_dataset_fields` | Orient before touching anything: what controls exist, what dataset paths are available/bound, and the real `controlId`s/part names to address edits against. |
| **Edit** | `insert_field`, `insert_label`, `insert_text`, `insert_page_number`, `insert_picture`, `insert_table`, `insert_repeater_table`, `insert_repeater_row`, `remove_control`, `set_cell_text`, `clear_cell_text`, and the table-structure tools (`insert_column`, `remove_column`, `set_column_widths`, `merge_cells`, `split_cells`, `set_cell_borders`) — plus `create_layout` / `refresh_xml_part` for lifecycle events | Make one deterministic, typed change at a time. |
| **Validate** | `validate_layout` — `level=quick` while iterating, `level=full` before calling it done | Confirm the edit didn't break structure or bindings; `full` dry-run-merges sample data so unresolved bindings surface as findings. |
| **Preview** | `preview_layout` | Visual/structural sanity check: merges sample (or real override) data and renders an offline PDF. |
| **Sandbox verify (OUTER, unchanged)** | *(outside this server)* AL build → publish to dev sandbox (`al-mcp-app` tooling) → run the report | The **only** real sign-off. `preview_layout`'s PDF is a mock — see §3 and §4. |

Every mutating tool — all sixteen of them, from `insert_field` through the table-structure tools to
`refresh_xml_part` — already runs its own post-edit `quick` validation and folds it into the response —
you don't need a separate `validate_layout` call just to see whether an edit broke something, but a
`full` pass is still worth running before moving on to preview/sandbox.

### Where the schema comes from

Every tool binds and validates against the BC dataset custom XML part embedded in the layout itself,
and this server **never generates that part** — `create_layout`'s `schemaSource` and
`refresh_xml_part`'s `newSchemaSource` must point at a BC-produced artifact, whose part content is
transplanted byte-for-byte. So a task like "create a new Word layout for report X" always starts by
obtaining a schema source, in one of three ways:

- **Your own AL report** (the usual case): declare the layout on the report object (`WordLayout`
  property or a `rendering` entry) and **build the AL project once** — the compiler creates the
  referenced `.docx` if it does not exist, already carrying the dataset part. Author into that file
  directly, or pass it as `create_layout`'s `schemaSource` for variants and templated builds.
- **A standard BC report**: export its built-in Word layout from Business Central (the Report
  Layouts page) and use the exported `.docx` as `schemaSource`.
- **A standalone schema `.xml`** exported from Business Central works anywhere a `.docx` source
  does (`schemaSource`, `newSchemaSource`, `list_dataset_fields`' `source`).

The same rule holds after every later dataset change: rebuild or re-export to get the new schema
artifact, then `refresh_xml_part` with it (§2, Lifecycle). If no BC-produced artifact exists yet —
the report is uncompiled AL source and nothing has been exported — the missing step is an AL build,
not a hand-authored XML part (see §4).

## 2. Tool reference

All tools take **absolute file paths**. Every tool returns the uniform envelope described in §5 and
never throws.

### Inspect (read-only)

| Tool | Key params (defaults) | Returns | Reach for it when… |
|---|---|---|---|
| `get_layout_info` | `layoutPath` | Report name/id/namespace/`storeItemID`, full control inventory (`kind`, `alias`, `tag`, `xpath`, `storeItemID`, `part`, `sdtId`, `usesW15Binding`, parent repeater), a control-kind summary count, the list of OOXML parts plus `partDetails` (each part's `kind`, header/footer `role` — `default`/`first`/`even`, `null` when unreferenced — and `isDefaultTarget`, the part a `partName`-less header/footer edit lands in) and `hasTitlePage` (whether page 1 renders the `first`-role parts INSTEAD of the default ones), and a `quick` validation summary. | First call on any layout. Also how you look up a real `controlId`/`part`/`partName` before addressing an edit — never guess these. Check `partDetails`/`hasTitlePage` before any header/footer edit: the package-order first part is frequently the even-page or first-page one. |
| `list_dataset_fields` | `source` (`.docx` layout **or** standalone schema `.xml`) | Nested data-item hierarchy; each column flagged `isLabel`; when `source` is a layout, each column is also flagged `bound`/`unbound`. | Finding the exact dataset path to pass to `insert_field` / `insert_label` / `insert_repeater_table`, or checking what's still unbound. |

### Validate

| Tool | Key params (defaults) | Returns | Reach for it when… |
|---|---|---|---|
| `validate_layout` | `layoutPath`, `level` = `"quick"` \| `"full"` (default `"quick"`) | `level`, `passed`, `errorCount`, `warningCount`, `findings[]` (`check`, `severity`, `message`, `location`). `quick`: OpenXML validity, single BC XML part, `storeItemID` match, XPath resolution, repeater shape/location, `attachedTemplate` warning, dangling `w:tblStyle` reference warning (a table naming a style the layout's styles part does not define — the reference silently does nothing). `full`: everything `quick` does, plus a real dry-run merge (sample data + repeater expansion) against a throwaway copy, surfacing every merge warning as a finding. | `quick` liberally (cheap, run it constantly); `full` before treating the loop as done — it's the only level that proves every binding actually resolves. |

### Preview

| Tool | Key params (defaults) | Returns | Reach for it when… |
|---|---|---|---|
| `preview_layout` | `layoutPath`, `rows`=`3`, `seed`=`12345`, `converter`=`"auto"` (prefers Word, falls back to LibreOffice; or force `"word"`/`"libreoffice"`), `dataOverridesPath`=`null`, `outputDir`=`null` | `mergedDocxPath`, `pdfPath` (`null` if conversion failed), `converterUsed`, `converterAvailable`, `conversionOk`, `conversionError`, `stats` (`fieldsFilled`, `repeatersExpanded`, `rowsGenerated`, `unresolved`, `picturesFilled`), `warnings[]`, `quickValidation`, and a fixed mock-render `disclaimer` string. | Visual/structural sanity check after an edit. Pass `dataOverridesPath` (a real exported BC dataset — the report UI's *Send to → XML* export works as-is, `decimalformatter` columns included) instead of generated sample data for a realistic look. `rows` is capped per repeater by an internal safeguard (default cap 100) — a larger value is reported via a row-cap warning, never silently dropped or blown out. |
| `render_preview_pages` | `pdfPath` (the `pdfPath` a `preview_layout` call returned), `firstPage`=`1` (1-based), `maxPages`=`3` (hard cap 10 per call), `dpi`=`120` (clamped 36–300) | The uniform JSON envelope as the result's FIRST text content block (including a `truncated` flag when the PDF has more pages than were returned), followed by one MCP **image content block** (PNG) per rendered page. | Looking at your own work: render the preview PDF's pages inline so you can visually inspect layout/structure without any external viewer. Page onward through a long document via `firstPage`. |

### Edit (mutating — body by default; write is validated before save, atomic, all-or-nothing)

**Shared location model** (`insert_field` / `insert_label` / `insert_repeater_table`): `locationType`
is one of `documentEnd` (end of body), `afterControl` (needs `controlId`), `tableCell` (needs
`tableIndex` + `row` + `col`, all 0-based), or `atText` (needs `searchText`, an ordinal substring
match). `layoutPart` picks which OOXML part `locationType` resolves within — `body` (default),
`header`, or `footer` — with an optional `partName` to pick a specific part when more than one exists.
Omitting `partName` targets the first section's **default** part (the everyday header/footer — often
NOT the first part in the package; `get_layout_info`'s `partDetails` gives each part's role). In a
layout with `hasTitlePage` (a distinct first page), the default header/footer does **not** render on
page 1 — content that must appear there needs the `first`-role part too, and the insert tools warn in
their summary when an edit lands in that trap.

| Tool | Key params (defaults) | Returns | Reach for it when… |
|---|---|---|---|
| `insert_field` | `layoutPath`, `field` (non-label leaf dataset path), `locationType` + its required params, `layoutPart`=`"body"` (also `"header"`/`"footer"`), `partName`=`null` | `operation`, `controlId`, `alias`, `xpath`, `kind`, `part`, `summary`, post-edit `quickValidation` | Binding a plain-text field anywhere in the body, a header, or a footer. |
| `insert_label` | Same shape as `insert_field`, but `label` must be a label column path (`Lbl`/`_Lbl`-suffixed, or any direct column of a data item named `Labels`) | Same shape as `insert_field` | Binding a label column — same location/part options as `insert_field`. |
| `insert_text` | `layoutPath`, `text` (literal, whitespace preserved verbatim — `" "` is a valid separator; empty is rejected), `locationType` + its required params, `layoutPart`=`"body"` (also `"header"`/`"footer"`), `partName`=`null`, `bold`=`null`, `fontSizePoints`=`null` (4–96, halves allowed) | Same shape as `insert_field`, but `controlId` is `0` — no content control is created | Plain STATIC text — the glue between controls: the separator space, colon, or `" / "` two inline controls need between them (without it they render as `Document NoDOCU-0150`), or a fixed caption the dataset does not provide (anything the dataset DOES provide should be `insert_label`, so it stays translatable). The run cannot later be targeted by `remove_control` or `afterControl`. |
| `insert_page_number` | `layoutPath`, `locationType` + its required params, `includeTotal`=`true` (the full stock `X / Y` shape: `PAGE` field, literal `" / "`, `NUMPAGES` field; `false` = the `PAGE` field alone), `layoutPart`=`"body"` (page numbers belong in `"header"`/`"footer"`), `partName`=`null`, `bold`=`null`, `fontSizePoints`=`null` | Same shape as `insert_field`, but `controlId` is `0` — field codes are plain runs, not a content control | The "Page X / Y" chrome every stock BC document header carries — Word/BC recalculate the numbers per page on render. Compose the stock idiom yourself so the caption stays translatable: `insert_label` the dataset's `Page_Lbl`, `insert_text` two spaces (`afterControl` on the label), then this tool (`afterControl` again). Like `insert_text`, the runs cannot later be targeted by `remove_control` or `afterControl`. |
| `insert_picture` | `layoutPath`, `picture` (picture dataset path, e.g. `/Header/CompanyPicture` — see `list_dataset_fields`), `locationType` + its required params, `widthMm`=`30`, `heightMm`=`30` (each 1–500), `layoutPart`=`"body"` (also `"header"`/`"footer"`), `partName`=`null` | Same shape as `insert_field` | The logo placeholder a from-scratch layout needs: a PICTURE content control BC fills at render time; `preview_layout` substitutes a placeholder image (counted in `stats.picturesFilled`). |
| `insert_table` | `layoutPath`, `rows` (1–100), `columns` (1–30), `locationType` + its required params, `columnWidths`=`null` (comma-separated twips, one per column; omit for an even split of the full content width), `columnAlignments`=`null` (comma-separated `left`/`center`/`right`), `withBorders`=`false` (default borderless — the shape of BC's own address/info/totals blocks), `layoutPart`=`"body"` **(only supported value — `header`/`footer` rejected)**, `partName`=`null` (always omit) | Same shape as `insert_field` plus the new table's index | A plain (unbound) table — the building block for non-repeating sections: address columns, label/value header grids, right-anchored totals blocks. Fill it afterwards with `set_cell_text` / `insert_field` (`locationType='tableCell'`). |
| `insert_repeater_row` | `layoutPath`, `parentControlId` (the OUTER repeater's `controlId`), `dataItem` (repeating data item DIRECTLY under the parent's, e.g. `/Header/Line/AssemblyLine`), `cells` (comma-separated specs laying the detail row on the parent grid: `-` = empty spacer, `Name` = one bound column, `a+b` = several columns inline in one cell, optional `N:` span prefix; spans must sum to the parent table's `columnCount`), `alignments`=`null` (per-cell `left`/`center`/`right`, or `-` to keep the default) | Same shape as `insert_field` | The standard BC nested-detail shape: a DETAIL ROW under each parent line (assembly components, serial/lot nos), aligned to the parent grid via spans, repeating once per parent row. This is how you CREATE nested repeaters — `insert_repeater_table` itself stays one flat repeater per call. |
| `insert_repeater_table` | `layoutPath`, `dataItem` (repeating, non-system data item), `columns` (comma-separated leaf column names, in order), `locationType` + its required params, `headerFromLabels`=`true` (bind header cells to `*Lbl` columns when found, else humanized static text), `tableStyle`=`null` (Word table style name, e.g. `"TableGrid"`), `columnWidths`=`null` (comma-separated twips, one per column), `columnAlignments`=`null` (comma-separated `left`/`center`/`right`, one per column — applied to the header cell AND the data cell, the way real BC line tables right-align a numeric column as a whole), `layoutPart`=`"body"` **(only supported value in v1)**, `partName`=`null` (unused — always omit) | Same shape as `insert_field` plus `columnCount` | The flagship tool for adding a new line-items-style table: builds a header row (label controls or humanized static text) + one data row wrapped in a repeating-section control, one bound field per column. **Body only** — passing `header`/`footer` is rejected with `invalid_argument`. |
| `remove_control` | `layoutPath`, `controlId`, `keepText`=`false` | Same shape as `insert_field` | Deleting a control (searches body + every header/footer). `keepText=true` drops only the control wrapper and keeps the visible content/row/cell in place; it cannot target a repeating-section control (that would orphan its row template — use the default `keepText=false` on the repeater instead). Removing a control that sits in a table cell always PRESERVES the cell/column (it is emptied, not deleted). |
| `set_cell_text` | `layoutPath`, `tableIndex`, `row`, `col` (all 0-based, as in `get_layout_info`'s `tables[]`), `text`, `layoutPart`=`"body"` (also `"header"`/`"footer"`), `partName`=`null` | `operation`, `part`, `tableIndex`, `row`, `col`, `previousText`, `newText`, `summary`, post-edit `quickValidation` | Setting/replacing the **plain text** of a table cell that is NOT a content control — e.g. re-labelling a line-items column header (`GST Amount` → `Tax`). Collapses the cell to one styled run; `w:tcPr` (width/span/borders) and paragraph/run style are preserved. A cell holding a bound field/label is rejected with `invalid_argument` (use `remove_control`/`insert_field`/`insert_label`). |
| `clear_cell_text` | `layoutPath`, `tableIndex`, `row`, `col`, `layoutPart`=`"body"`, `partName`=`null` | Same shape as `set_cell_text` (`newText` empty) | Removing all plain text from a cell, leaving a valid empty cell — e.g. blanking an orphaned column-header label after `remove_control` emptied the matching data-field cells. The cell/column is preserved. Does **not** delete the column (that's a table-structure change — see §3). Rejects control cells like `set_cell_text`. |

### Table structure (mutating — same write safety, plus a grid-consistency guard)

All six are addressed by the SAME `tableIndex`/`row`/cell indices `get_layout_info` reports, take
`layoutPart`=`"body"` (also `"header"`/`"footer"`) + `partName`=`null`, and pass the shared
grid-consistency backstop that rejects (`edit_would_break_table`, file untouched) any edit desyncing
a row from its `w:tblGrid`. Rows that skip grid columns (`w:gridBefore`/`w:gridAfter` — real BC
line tables have them) are handled; tables using **vertical merges (`w:vMerge`) are rejected** by
every tool here except `set_cell_borders`.

| Tool | Key params (defaults) | Returns | Reach for it when… |
|---|---|---|---|
| `insert_column` | `layoutPath`, `tableIndex`, `mode` = `"field"` \| `"label"` \| `"plainText"`, `dataPath` (for `field`/`label`: full dataset path, exactly as `insert_field`/`insert_label` take), `headerText`=`null` (omit to humanize the `dataPath` leaf), `headerLabelPath`=`null` (bind the header cell to a label column instead of static text), `atColumn`=`null` (0-based GRID position; omit to append at the far right), `width`=`null` (twips; omit for the mean of existing columns) | Same shape as `insert_field` | Adding a column to a line-items or plain table. Bound control lands in the repeater's DATA row; header rows get a header cell; right-anchored totals-block rows are widened, not celled, so the block keeps its look. An interior `atColumn` that falls INSIDE a spanned content cell is refused, naming the cell (`split_cells` there first). |
| `remove_column` | `layoutPath`, `tableIndex`, `column` (0-based GRID index) | Same shape as `insert_field` | "Remove the GST Amount column from the lines." Drops each row's covering cell — including a bound one, whose binding goes with it (UNLIKE `remove_control`, which preserves the cell) — or decrements its `gridSpan`; removes the `w:gridCol`; redistributes the width proportionally (follow with `set_column_widths` for a different distribution). The last remaining column cannot be removed. |
| `set_column_widths` | `layoutPath`, `tableIndex`, `widths` (comma-separated twips, exactly one per GRID column) | Same shape as `set_cell_text` | Resizing columns; `gridSpan`-aware — each cell becomes the sum of the columns it spans. |
| `merge_cells` | `layoutPath`, `tableIndex`, `row`, `fromColumn`, `toColumn` (0-based PHYSICAL cell indices, inclusive) | Same shape as `set_cell_text` | Horizontally merging a run of adjacent cells in ONE row into a single `gridSpan` cell — the first cell (content and binding) is kept and widened, the rest are deleted. Refused if an absorbed cell holds a bound control (`remove_control` it first). HORIZONTAL only. |
| `split_cells` | `layoutPath`, `tableIndex`, `row`, `cellIndex` (0-based PHYSICAL index of the spanned cell) | Same shape as `set_cell_text` | Splitting one spanned cell back into single-column cells (content stays in the first). Refused if the cell is not spanned. |
| `set_cell_borders` | `layoutPath`, `tableIndex`, `row`, `edges` (comma-separated `top`/`bottom`/`left`/`right`, or `all`), `col`=`null` (omit = every cell in the row — the usual case), `style`=`"single"` (draws) \| `"none"` (explicitly clears), `size`=`4` (eighths of a point, 2–96; 4 = ½ pt, the only thickness the BC corpus uses) | Same shape as `set_cell_text` | The rules a BC document's look is actually made of: "a line above the totals row", "underline the grand-total cell". Edges you do not name are left as they were. Cosmetic only — works even on `w:vMerge` tables. |

### Lifecycle

| Tool | Key params (defaults) | Returns | Reach for it when… |
|---|---|---|---|
| `create_layout` | `schemaSource` (existing `.docx` layout **or** schema `.xml`), `outputPath`, `templatePath`=`null` (an **unbound** branded/styled shell — headers/footers/logo/fonts/styles, NOT a full BC layout with its own bound controls), `headingText`=`null` (the heading paragraph a BLANK build's body starts with — omit for the report's own name, pass a human document title when authoring from scratch, or `""` for no heading; ignored when a template's body already has content) | `outputPath`, report identity, `storeItemId`, `usedTemplate`, `replacedExistingBcPart`, `quickValidation` | Starting a brand-new layout, optionally over a branded template — §1 ("Where the schema comes from") covers how to obtain the `schemaSource` artifact when none exists yet. Always ships exactly one BC custom XML part (fresh `storeItemID`) plus the glossary part the `insert_*` placeholders depend on. A BLANK (non-template) layout also ships empty header/footer parts wired into its page setup **and a default styles part pinning its typography** (Calibri 11 pt `docDefaults` plus the standard `Normal`/`TableGrid` definitions), so it renders with the same font in Word and in Business Central — a template keeps its own headers/footers and styles/theme untouched. `templatePath` MUST be unbound: if it already carries a BC part **and** bound controls that would go stale against the fresh `storeItemID`, the call fails outright with `template_not_unbound` (nothing is written) instead of silently shipping a broken layout — see §5. A template whose BC part has zero bound controls of its own still succeeds. |
| `refresh_xml_part` | `layoutPath`, `newSchemaSource` (existing `.docx` **or** schema `.xml`) | Old/new report identity, `namespaceChanged`, `remappedCount`, `orphanedBindings[]` (`alias`, `xpath`, `part`), `newUnboundFields[]`, `quickValidation` | After the AL report dataset changes. Replaces the dataset part's content **in place** (keeps the same `storeItemID`, so every existing binding still links to the same part) and reclassifies every existing binding by element name — matches are remapped/kept, non-matches are reported as `orphanedBindings` and **left in place** (this tool never deletes or rebinds anything itself — that's your call via `remove_control`/`insert_field`/`insert_label`). A non-zero `errorCount` here from `xpath-resolves` findings is the *expected* corroboration of the orphan report, not a failed refresh. |

## 3. Supported matrix (v1)

**Works:**
- Field and label content controls (plain-text SDTs), in the document **body, headers, and footers
  alike** (`insert_field`/`insert_label` both take `layoutPart`).
- Repeater **TABLES in the document body** (`insert_repeater_table`: header row + one data row wrapped
  in a repeating-section control, per-column bound fields).
- Repeaters nested **one or more levels deep**: reading, validating, merging, and previewing a
  pre-existing multi-level repeater structure works end to end (proven against the real corpus, up to
  5 levels deep, via per-level XPath re-anchoring). Creation: `insert_repeater_table` builds one flat
  repeater per call, and `insert_repeater_row` adds a nested DETAIL-ROW repeater inside an existing
  repeater's item — the standard BC per-line-detail shape.
- Static text (`insert_text`, with optional `bold`/`fontSizePoints`), plain unbound tables
  (`insert_table`), and picture controls (`insert_picture`, mm-sized); existing and inserted picture
  controls get a placeholder image blip during merge (counted in `stats.picturesFilled`).
- Page-number field codes (`insert_page_number`): the stock `PAGE`/`NUMPAGES` "Page X / Y" header
  chrome, emitted exactly as the corpus layouts carry it. (Page-position-CONDITIONAL content —
  `IF PAGE = NUMPAGES` constructs — remains deferred, GitHub issue #11.)
- **Table-structure changes**: `insert_column` / `remove_column` / `set_column_widths` /
  `merge_cells` / `split_cells` (all `gridSpan`-aware and grid-column-addressed, guarded against
  desyncing a row from its `w:tblGrid`; rows skipping grid columns via `w:gridBefore`/`w:gridAfter`
  are handled) and `set_cell_borders` for the per-cell rules BC documents get their look from.
- Inline page images of the preview PDF (`render_preview_pages`), so you can visually inspect your
  own work without an external viewer.
- `create_layout` (blank or from an unbound branded template) and `refresh_xml_part` (in-place dataset
  swap + binding reclassification) for layout lifecycle.
- Both validation levels: `quick` (structural/binding) and `full` (dry-run merge).
- Offline PDF preview with a choice of converter: Word COM (primary) or LibreOffice (fallback), via
  `converter=auto|word|libreoffice`.

**NOT supported / deferred:**
- **Hide-if-empty / BC add-in conditional controls** (Hide Field if Zero, Hide Empty Table, Hide Empty
  Table Row, layout comment) — there is no `set_hide_if_empty` tool. Deferred pending research: the BC
  Word add-in's OOXML encoding for these has never been reverse-engineered (and no corpus layout
  happens to use them either, so there's nothing to imitate yet).
- **Repeater tables in headers/footers** — deferred. `insert_repeater_table` and `insert_table`
  reject any `layoutPart` other than `body` with `invalid_argument`; `validate_layout` (`quick`)
  still flags a pre-existing/hand-authored header or footer repeater with a
  `repeater-in-header-footer` **warning** (not an error) rather than silently accepting it.
- **Vertical merges (`w:vMerge`)** — every table-structure tool except the cosmetic
  `set_cell_borders` rejects a table using them, with a clear message, pending a real-add-in OOXML
  capture. (Ragged rows, `gridSpan` spans, and `w:gridBefore`/`w:gridAfter` skips are all supported —
  the refusal is vMerge specifically.)
- **Real XLF caption translation** — `preview_layout` fills labels from the dataset's own label
  columns / humanized element names, not real captions translated from the AL project's XLF files.
  Pass a real exported dataset via `dataOverridesPath` when caption text matters.
- **Broad cosmetic formatting** (fonts, colors, margins, styles/branding) — beyond the deliberate
  knobs that exist (`bold`/`fontSizePoints` on the insert tools, cell/column alignments,
  `set_cell_borders`' per-cell rules, and the deterministic default typography a blank
  `create_layout` ships — Calibri 11 pt via its scaffolded styles part), there is no styling tool;
  restyle via `create_layout`'s `templatePath` or hand-edit the OOXML directly, then
  run `validate_layout` to confirm the hand-edit didn't break structure or bindings.
- **RDL layouts, Excel layouts** — out of scope entirely; this server is Word-only.
- **Any BC-connected upload** of a layout to a tenant — no public API for this exists; it is UI-only
  in Business Central, a deliberate scope decision rather than a gap to script around.

## 4. Anti-patterns

- **Don't free-hand OOXML when a typed tool already covers the change.** The insert tools, the
  table-structure tools, `remove_control`, and `refresh_xml_part` exist precisely to avoid
  hand-authored bindings, ids, XPaths, and `w:tblGrid` edits going subtly wrong. Reserve hand-editing
  for what genuinely has no tool (broad cosmetic styling) — and always follow it with
  `validate_layout`. In particular: dropping a line-items column is `remove_column`, one call — not a
  remove-and-blank workaround, and not a hand edit.
- **Never hand-author or synthesize the BC dataset XML part itself** — not from the report's `.al`
  source, not from memory of the compiler's naming rules. The schema must come from a BC-produced
  artifact (a compiler-generated layout, an exported stock layout, or an exported schema `.xml` —
  see §1, "Where the schema comes from"). A synthesized part that drifts from the compiler's
  element-name/label derivation produces bindings BC orphans the moment it regenerates the part.
- **Never treat `preview_layout`'s PDF as final sign-off.** It's an explicit MOCK: deterministic
  sample data (or your override), converted outside the real BC report engine. The outer loop — AL
  build → publish to dev sandbox → run the report — is unchanged and is the only real verification.
- **Don't try to create a repeater table in a header or footer.** `insert_repeater_table` only accepts
  `layoutPart='body'`; this is a known deferred gap (GitHub issue #10), not something to work around by
  hand-authoring one.
- **Don't ignore a failure's `error.hint`/`error.code`.** Every failure is agent-actionable — it names
  the exact argument to fix or the inspection tool to call first (typically `get_layout_info` /
  `list_dataset_fields`), not a generic "something went wrong."
- **Don't assume `ok:true` means "fully clean."** Check the post-edit `quickValidation` on every
  mutating response, and `orphanedBindings` / `newUnboundFields` after `refresh_xml_part`.
- **Don't pass a full BC layout as `create_layout`'s `templatePath`.** It must be an unbound
  branded/styled shell — a template that already carries a BC part *and* bound controls that would go
  stale is refused outright (`template_not_unbound`). To reuse an existing layout's design against a
  new schema, copy it and call `refresh_xml_part` on the copy instead; to reuse it as a branded shell,
  strip its bound controls with `remove_control` first (or supply a genuinely unbound template).
- **After any AL dataset change, call `refresh_xml_part`, then re-validate** (ideally `level=full`) —
  don't hand-replace the XML part or leave stale bindings. Then act on `orphanedBindings` explicitly
  via `remove_control` / `insert_field` / `insert_label` — the tool deliberately never does this for
  you.
- **Supply a real dataset via `dataOverridesPath`** for a `preview_layout` call that needs to look
  realistic — generated sample data is type-aware but still synthetic.
- **Don't over-read the mock preview's visual details.** Caption text, exact fonts, and pagination are
  known, documented fidelity gaps (see `docs/FIDELITY-CHECKLIST.md`) — they can differ from a real BC
  sandbox render even when the preview looks "right."
- **Don't guess a `controlId`, table index, or header/footer `partName`.** They are per-document and
  not sequential or guessable — always look them up via `get_layout_info` (`controls[].sdtId`,
  `controls[].part`, `partDetails[]`) first. In particular, never assume `header1.xml` is "the"
  header: in most stock BC layouts it is the even-page or first-page part — `partDetails[].role`
  says which is which.

## 5. Error handling

Every tool returns the same envelope and never throws across the MCP boundary:

```
{ "ok": bool, "data": <tool-specific payload or null>, "error": { "code", "message", "hint" } | null }
```

- `ok:false` always comes with a populated `error`. `code` is one of `file_not_found`,
  `invalid_layout`, `invalid_argument`, `not_found`, `edit_would_corrupt` (a mutating-tool
  structural-safety rejection — the file is left untouched), `edit_would_break_table` (the
  table-structure tools' grid-consistency rejection — a corruption class the OpenXML validator
  accepts silently; file untouched), `file_locked` (another process holds the file — e.g. the layout
  is open in Word), `template_not_unbound` (`create_layout`-only: `templatePath` was a full BC
  layout, not an unbound shell — see §"Lifecycle" above; nothing is written), or `internal_error`.
  `hint` is always non-empty and actionable: which argument to fix and its valid values, or which
  inspection tool to call. Act on it before retrying blindly.
- `ok:true` doesn't mean "nothing to look at" — check the tool's own payload for its validation
  signal:
  - Mutating tools (`insert_field` / `insert_label` / `insert_repeater_table` / `remove_control` /
    `refresh_xml_part`): `data.quickValidation.passed` / `errorCount` / `warningCount`.
  - `create_layout`: `data.quickValidation` — a template whose own pre-existing bound controls would
    go stale is refused outright (`template_not_unbound`, see above) rather than reported as data here.
  - `refresh_xml_part`: `data.orphanedBindings` and `data.newUnboundFields` in addition to
    `quickValidation` — a non-zero `errorCount` here from `xpath-resolves` findings is expected
    corroboration of the orphan report, not a failure.
  - `preview_layout`: `data.conversionOk` / `data.conversionError` (PDF conversion can fail or be
    unavailable even though the merge itself succeeded), plus `data.warnings[]` and
    `data.quickValidation`.
- Mutating tools apply a **structural gate only**: a write is rejected solely if it would introduce a
  *new* OpenXML structural error versus the file's own pre-edit baseline. Semantic issues (like an
  orphaned binding) are always allowed through and reported as validation findings, never blocked.
