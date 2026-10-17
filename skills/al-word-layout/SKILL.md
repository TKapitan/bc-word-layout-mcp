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
| **Edit** | `insert_field`, `insert_label`, `insert_text`, `insert_page_number`, `insert_picture`, `insert_table`, `insert_repeater_table`, `insert_repeater_row`, `insert_table_row`, `insert_subtotal_row`, `remove_control`, `set_cell_text`, `clear_cell_text`, and the table-structure tools (`insert_column`, `remove_column`, `set_column_widths`, `merge_cells`, `split_cells`, `set_cell_borders`) — plus `create_layout` / `refresh_xml_part` for lifecycle events | Make one deterministic, typed change at a time. |
| **Validate** | `validate_layout` — `level=quick` while iterating, `level=full` before calling it done | Confirm the edit didn't break structure or bindings; `full` dry-run-merges sample data so unresolved bindings surface as findings. |
| **Preview** | `preview_layout` | Visual/structural sanity check: merges sample (or real override) data and renders an offline PDF. |
| **Sandbox verify (OUTER, unchanged)** | *(outside this server)* AL build → publish to dev sandbox (`al-mcp-app` tooling) → run the report | The **only** real sign-off. `preview_layout`'s PDF is a mock — see §3 and §4. |

Every one of the 20 mutating tools — the 12 in §2's Edit table, the 7 table-structure tools, and
`refresh_xml_part` — already runs its own post-edit `quick` validation and folds it into the response
(`create_layout` does too, for the layout it builds) — you don't need a separate `validate_layout` call
just to see whether an edit broke something, but a `full` pass is still worth running before moving on
to preview/sandbox.

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

All tools take **absolute file paths**: a relative one is refused with `invalid_argument` (it would
otherwise resolve against the server's working directory, not your workspace). Every tool returns the
uniform envelope described in §5 and never throws.

### Inspect (read-only)

| Tool | Key params (defaults) | Returns | Reach for it when… |
|---|---|---|---|
| `get_layout_info` | `layoutPath` | `report` (name, id, namespace, `storeItemId`), full control inventory `controls[]` (`kind`, `alias`, `tag`, `xPath`, `storeItemId`, `part`, `sdtId`, `usesW15Binding`, `parentRepeaterXPath`, `level` — the `w:sdt` kind: `run`/`block`/`cell`/`row` — and, for a control in a table, `tableIndex`/`rowIndex`/`colIndex`), a `controlSummary` count per kind, the list of OOXML `parts` plus `partDetails` (each part's `kind`, header/footer `role` — `default`/`first`/`even`, absent when no section references the part — and `isDefaultTarget`, the part a `partName`-less header/footer edit lands in) and `hasTitlePage` (whether page 1 renders the `first`-role parts INSTEAD of the default ones), a `tables[]` section (per table `rowCount`/`columnCount`/`gridColumnWidths`; ONE row per `w:tr` — every row of a multi-row repeater item, nested detail and subtotal rows included, each with its own `rowIndex` and tagged `repeaterControlId`, the innermost repeater it repeats with — and per row `gridBefore`/`gridAfter`; per cell its owning control, text, physical `colIndex` and grid `gridColumn`/`gridSpan`), and `validation`, a `quick` validation summary. | First call on any layout. Also how you look up a real `controlId`/`part`/`partName` before addressing an edit — never guess these. Check `partDetails`/`hasTitlePage` before any header/footer edit: the package-order first part is frequently the even-page or first-page one. |
| `list_dataset_fields` | `source` (`.docx` layout **or** standalone schema `.xml`) | Nested data-item hierarchy; each column flagged `isLabel`; when `source` is a layout, each column is also flagged `bound`/`unbound`. | Finding the exact dataset path to pass to `insert_field` / `insert_label` / `insert_repeater_table`, or checking what's still unbound. |

### Validate

| Tool | Key params (defaults) | Returns | Reach for it when… |
|---|---|---|---|
| `validate_layout` | `layoutPath`, `level` = `"quick"` \| `"full"` (default `"quick"`) | `level`, `passed`, `errorCount`, `warningCount`, `findings[]` (`check`, `severity`, `message`, `location`). `quick`: OpenXML validity, `document-content-type` **error** (the main part is typed as a Word template or macro-enabled file — Word refuses to open it as a `.docx`, whatever the file name says), single BC XML part (with several, the finding lists them and the other checks measure against the one the bindings name by `storeItemID`; `refresh_xml_part` and every tool that writes a binding refuse such a layout with `invalid_layout` until the extra part is removed from the package), `storeItemID` match, `binding-namespace` **error** (a binding names another report's BC namespace), `binding-prefix-mapping` **error** (a dataset binding's `w:prefixMappings` does not declare every prefix its XPath uses as the layout's own namespace — a misspelled or `dynamics-365` namespace, or `ns1` declared for an `ns0` XPath; Word cannot map the control. A bare URI naming the right namespace is tolerated: stock layouts ship it and BC prints them), XPath resolution, repeater shape/location, `attachedTemplate` warning, dangling `w:tblStyle` reference warning (a table naming a style the layout's styles part does not define — the reference silently does nothing), `compatibility-mode` warning (the layout CONTAINS repeaters but declares a Word compatibility mode below 15 — a Word save would convert them to plain rich-text controls and drop their bindings; blank builds declare 15, so this fires for a `templatePath` shell or a layout authored before that scaffold). `repeater-downgraded` **error** (a control whose alias names a whole data item but which is not a repeating section — what a Word save leaves behind when it converts a repeater to a plain rich-text control; the row then renders once instead of once per record). `top-level-repeater-group` **error** (a block-level repeating section over the TOP-LEVEL data item whose `repeatingSectionItem` is not inside a `w:group` sdt — BC prints the whole layout as an empty document, although `preview_layout` and `full` fill it normally; bind the body straight to `<Item>[1]` instead, since BC already renders one document per top-level record, or use the StandardStatement shape repeatingSection › `w:group` › repeatingSectionItem). `duplicate-control-id` **warning** (two content controls share a `w:sdtPr/w:id` — the file still opens, but `remove_control`, an `afterControl` location and `parentRepeaterId` refuse that id with `invalid_layout` rather than act on the first match). One structural complaint is deliberately a **warning**, not an error: `w:rPr/w:w=0`, the character scaling Word writes into a content control's placeholder run — it renders correctly in Word and BC and no tool here can edit run properties. `full`: everything `quick` does, plus a real dry-run merge (sample data + repeater expansion) against a throwaway copy, surfacing every merge warning as a finding. | `quick` liberally (cheap, run it constantly); `full` before treating the loop as done — it's the only level that proves every binding actually resolves. |

### Preview

| Tool | Key params (defaults) | Returns | Reach for it when… |
|---|---|---|---|
| `preview_layout` | `layoutPath`, `rows`=`3` (1 or more), `seed`=`12345`, `converter`=`"auto"` (prefers Word, falls back to LibreOffice; or force `"word"`/`"libreoffice"`), `dataOverridesPath`=`null`, `outputDir`=`null` | `mergedDocxPath`, `pdfPath` (absent if conversion failed), `converterUsed`, `converterAvailable`, `conversionOk`, `conversionError`, `stats` (`fieldsFilled`, `repeatersExpanded`, `rowsGenerated`, `unresolved`, `picturesFilled`), `warnings[]`, `quickValidation`, and a fixed mock-render `disclaimer` string. | Visual/structural sanity check after an edit. Pass `dataOverridesPath` (a real exported BC dataset — the report UI's *Send to → XML* export works as-is, `decimalformatter` columns included) instead of generated sample data for a realistic look; a file that is not well-formed XML, not one of the two shapes, or exported from a different report is refused with `invalid_argument`. `rows` is capped per repeater by an internal safeguard (default cap 100) — a larger value is reported via a row-cap warning, never silently dropped or blown out. |
| `render_preview_pages` | `pdfPath` (the `pdfPath` a `preview_layout` call returned), `firstPage`=`1` (1-based, 1 or more), `maxPages`=`3` (1 or more; hard cap 10 per call), `dpi`=`120` (clamped 36–300; each page is also capped at 5000 px a side and 20 MP — a larger page comes out smaller, flagged `pages[].downscaled`), `outputDir`=`null` (also WRITE each page as `<pdf-basename>-pageNN.png` there; paths returned on `pages[].path`, directory created if missing, its lifetime is yours), `inlineImages`=`true` (`false` + `outputDir` writes the files without spending tokens on the images; `false` without `outputDir` is refused) | The uniform JSON envelope as the result's FIRST text content block (including a `truncated` flag when the PDF has more pages than were returned), followed by one MCP **image content block** (PNG) per rendered page. | Looking at your own work: render the preview PDF's pages inline so you can visually inspect layout/structure without any external viewer. Page onward through a long document via `firstPage`. **When a human asks to see the render**, pass `outputDir` — the inline blocks are visible to you only, so a file on disk is the only thing they can open. |

### Edit (mutating — body by default; write is validated before save, atomic, all-or-nothing)

**Shared location model** (`insert_field` / `insert_label` / `insert_repeater_table`): `locationType`
is one of `documentEnd` (end of body), `afterControl` (needs `controlId`), `tableCell` (needs
`tableIndex` + `row` + `col`, all 0-based), or `atText` (needs `searchText`, an ordinal substring
match). `afterControl` inserts **directly after the anchor control**, so an inline line is built
**right to left**: insert the anchor, then the element that ends up last, then the one before it,
each `afterControl` on the same anchor. For `Label: Value`: `insert_label`, then `insert_field`
(`afterControl` = the label), then `insert_text ": "` (`afterControl` = the label). Plain runs
(`insert_text`, `insert_page_number`) return no `controlId`, so a control is always the anchor —
and `documentEnd` starts a new paragraph on every call, so it cannot chain a line. `layoutPart` picks which OOXML part `locationType` resolves within — `body` (default),
`header`, or `footer` — with an optional `partName` to pick a specific part when more than one exists.
Omitting `partName` targets the first section's **default** part (the everyday header/footer — often
NOT the first part in the package; `get_layout_info`'s `partDetails` gives each part's role). In a
layout with `hasTitlePage` (a distinct first page), the default header/footer does **not** render on
page 1 — content that must appear there needs the `first`-role part too, and the insert tools warn in
their summary when an edit lands in that trap.

A location that resolves **inside a data-bound control's own content** is refused with
`invalid_argument`, for every insert tool: `afterControl` on a **cell-level** field/label (it wraps
its whole cell, so there is no "after it" in that cell), `tableCell` on a cell that *is* such a
control (`get_layout_info`'s `tables[]` cells report `isControlCell`), or `atText` matching a block-
or cell-level control's text. Word discards — or, for page numbers/tables in a plain-text control,
refuses to open — anything placed there. Put the content in a plain cell via `tableCell`, or
`afterControl` an **inline** control (an inline field's paragraph is outside the field, so the
label → `insert_page_number` → `insert_text` chain, each anchored on the label, still works). Cells of a repeater row are plain
cells — the repeater is a container, not a bound field.

| Tool | Key params (defaults) | Returns | Reach for it when… |
|---|---|---|---|
| `insert_field` | `layoutPath`, `field` (non-label leaf dataset path), `locationType` + its required params, `layoutPart`=`"body"` (also `"header"`/`"footer"`), `partName`=`null`, `bold`=`null` (`true`/`false` sets or strips bold; omitted, the control's runs stay unstyled), `fontSizePoints`=`null` (4–96, halves allowed) | `operation`, `controlId`, `alias`, `xPath`, `kind`, `part`, `summary`, post-edit `quickValidation` | Binding a plain-text field anywhere in the body, a header, or a footer. |
| `insert_label` | `layoutPath`, `label` (a label column path: `Lbl`/`_Lbl`-suffixed, or any direct column of a data item named `Labels`), `locationType` + its required params, `layoutPart`=`"body"`, `partName`=`null`, `bold`=`null`, `fontSizePoints`=`null` — exactly as `insert_field` | Same shape as `insert_field` | Binding a label column — same location/part options as `insert_field`. Stock captions are often smaller and bold: pass `bold=true` and a `fontSizePoints`. |
| `insert_text` | `layoutPath`, `text` (literal, whitespace preserved verbatim — `" "` is a valid separator; empty is rejected), `locationType` + its required params, `layoutPart`=`"body"` (also `"header"`/`"footer"`), `partName`=`null`, `bold`=`null`, `fontSizePoints`=`null` (4–96, halves allowed) | Same shape as `insert_field`, but `controlId` is `0` — no content control is created | Plain STATIC text — the glue between controls: the separator space, colon, or `" / "` two inline controls need between them (without it they render as `Document NoDOCU-0150`), or a fixed caption the dataset does not provide (anything the dataset DOES provide should be `insert_label`, so it stays translatable). Insert a separator **after** the controls it separates, anchored on the left one (the right-to-left rule above). The run cannot later be targeted by `remove_control` or `afterControl`. |
| `insert_page_number` | `layoutPath`, `locationType` + its required params, `includeTotal`=`true` (the full stock `X / Y` shape: `PAGE` field, literal `" / "`, `NUMPAGES` field; `false` = the `PAGE` field alone), `layoutPart`=`"body"` (page numbers belong in `"header"`/`"footer"`), `partName`=`null`, `bold`=`null`, `fontSizePoints`=`null` | Same shape as `insert_field`, but `controlId` is `0` — field codes are plain runs, not a content control | The "Page X / Y" chrome every stock BC document header carries — Word/BC recalculate the numbers per page on render. Compose the stock idiom yourself so the caption stays translatable: `insert_label` the dataset's `Page_Lbl`, then this tool (`afterControl` on the label), then `insert_text` two spaces (`afterControl` on the label again) — right to left, so the spaces land between the label and the fields. Like `insert_text`, the runs cannot later be targeted by `remove_control` or `afterControl`. |
| `insert_picture` | `layoutPath`, `picture` (picture dataset path, e.g. `/Header/CompanyPicture` — see `list_dataset_fields`), `locationType` + its required params, `widthMm`=`30`, `heightMm`=`30` (each 1–500), `layoutPart`=`"body"` (also `"header"`/`"footer"`), `partName`=`null` | Same shape as `insert_field` | The logo placeholder a from-scratch layout needs: a PICTURE content control BC fills at render time; `preview_layout` substitutes a placeholder image (counted in `stats.picturesFilled`). |
| `insert_table` | `layoutPath`, `rows` (1–100), `columns` (1–30), `locationType` + its required params, `columnWidths`=`null` (comma-separated twips, each 1–31680, one per column; omit for an even split of the layout's content width — its page width minus margins, or 10206 when it declares none — the widths are **binding**: the table declares `w:tblLayout="fixed"` + a `w:tblW` total, so Word does not autofit them away), `columnAlignments`=`null` (comma-separated `left`/`center`/`right`), `withBorders`=`false` (default borderless — the shape of BC's own address/info/totals blocks; `true` draws a single-line border on every edge and between every cell, the grid `insert_repeater_table` draws with `look="grid"`), `layoutPart`=`"body"` **(only supported value — `header`/`footer` rejected)**, `partName`=`null` (always omit) | The table-edit result (see *Table structure* below), whose `tableIndex` is the new table's index | A plain (unbound) table — the building block for non-repeating sections: address columns, label/value header grids, right-anchored totals blocks. Fill it afterwards with `set_cell_text` / `insert_field` (`locationType='tableCell'`). |
| `insert_repeater_row` | `layoutPath`, `parentControlId` (the OUTER repeater's `controlId`), `dataItem` (repeating data item DIRECTLY under the parent's, e.g. `/Header/Line/AssemblyLine`), `cells` (comma-separated specs laying the detail row on the parent grid: `-` = empty spacer, `Name` = one bound column, `a+b` = several columns inline in one cell, optional `N:` span prefix (N 1–63); spans must sum to the parent table's `columnCount`), `alignments`=`null` (per-cell `left`/`center`/`right`, or `-` to keep the default) | Same shape as `insert_field` | The standard BC nested-detail shape: a DETAIL ROW under each parent line (assembly components, serial/lot nos), aligned to the parent grid via spans, repeating once per parent row. This is how you CREATE nested repeaters — `insert_repeater_table` itself stays one flat repeater per call. |
| `insert_subtotal_row` | `layoutPath`, `parentControlId` (the repeater's `controlId`), `cells` (same DSL as `insert_table_row`: `-` spacer, **FULL dataset paths** — the corpus subtotal binds a sibling non-repeating `Subtotals` item and `Labels` captions; `a+b` chaining, `N:` spans; spans must sum to the parent table's `columnCount`), `alignments`=`null`, `bold`=`null`, `fontSizePoints`=`null` | Same shape as `insert_field`, but `controlId` is `0` (the row is a plain `w:tr`; its bound cells' ids are in `summary`) | A STATIC per-group row appended at the END of a repeater's item — the stock group-subtotal shape (SalespersonCommission): renders once per PARENT row, after the nested detail rows, never per detail row (contrast `insert_table_row`, which renders once for the whole table). Repeated calls stack: spacer row (`cells="9:-"`) first, then the bold subtotal row. The group HEADER needs no extra call — the repeater's own line row is it. |
| `insert_repeater_table` | `layoutPath`, `dataItem` (repeating, non-system data item), `columns` (comma-separated leaf column names, in order), `locationType` + its required params, `headerFromLabels`=`true` (bind header cells to `*Lbl` columns when found, else humanized static text), `tableStyle`=`null` (Word table style name, e.g. `"TableGrid"`), `columnWidths`=`null` (comma-separated twips, each 1–31680, one per column; omit for an even split of the document's content width — the host cell's width for a table placed in a cell — so it always fits the page — **binding**: emitted with `w:tblLayout="fixed"` + a `w:tblW` total, so Word renders the grid as specified rather than autofitting it to the merged content), `columnAlignments`=`null` (comma-separated `left`/`center`/`right`, one per column — applied to the header cell AND the data cell, the way real BC line tables right-align a numeric column as a whole), `look`=`"bc"` (the BC-native look: no grid, a ½ pt rule under the header row, as every stock lines table; `"grid"` draws a single-line border on every edge and between every cell), `layoutPart`=`"body"` **(the only supported value)**, `partName`=`null` (unused — always omit) | Same shape as `insert_field` plus `columnCount`, `tableIndex` (the new table's index) and `dataRowIndex` (the index of its repeating data row), so follow-up edits need no re-read | The flagship tool for adding a new line-items-style table: builds a header row (label controls or humanized static text) + one data row wrapped in a repeating-section control, one bound field per column. **Body only** — passing `header`/`footer` is rejected with `invalid_argument`. |
| `remove_control` | `layoutPath`, `controlId`, `keepText`=`false` | Same shape as `insert_field` | Deleting a control (searches body + every header/footer). `keepText=true` drops only the control wrapper and keeps the visible content/row/cell in place; it cannot target a repeating-section control (that would orphan its row template — use the default `keepText=false` on the repeater instead), nor the unbound `w:group` control between a block-level top-level repeater and its item (e.g. StandardStatement's body-wide `/Customer` repeater — without the group BC prints an empty document). Removing a control that sits in a table cell always PRESERVES the cell/column (it is emptied, not deleted). |
| `set_cell_text` | `layoutPath`, `tableIndex`, `row`, `col` (all 0-based, as in `get_layout_info`'s `tables[]`), `text`, `layoutPart`=`"body"` (also `"header"`/`"footer"`), `partName`=`null`, `bold`=`null`, `alignment`=`null` (`left`/`center`/`right`), `fontSizePoints`=`null` (4–96) — each formatting knob applies only when passed; omitted, the cell keeps its existing look. They are for cells of a fresh `insert_table`, which have nothing to inherit: bold captions, right-aligned amounts, a bigger title | `operation`, `part`, `tableIndex`, `row`, `col`, `previousText`, `newText`, `summary`, post-edit `quickValidation` | Setting/replacing the **plain text** of a table cell that is NOT a content control — e.g. re-labelling a line-items column header (`GST Amount` → `Tax`). Replaces the cell's text runs with one styled run; `w:tcPr` (width/span/borders) and paragraph/run style are preserved. Only text is replaced — a nested table, picture or field in the cell stays in place (the `summary` names it), and `previousText` is the replaced text only. A cell holding a bound field/label is rejected with `invalid_argument` (use `remove_control`/`insert_field`/`insert_label`). |
| `clear_cell_text` | `layoutPath`, `tableIndex`, `row`, `col`, `layoutPart`=`"body"`, `partName`=`null` | Same shape as `set_cell_text` (`newText` empty) | Removing all plain text from a cell, leaving a valid empty cell — e.g. blanking an orphaned column-header label after `remove_control` emptied the matching data-field cells. The cell/column is preserved, and so is any non-text content in it (nested table, picture, field). Does **not** delete the column (that's a table-structure change — see §3). Rejects control cells like `set_cell_text`. |

### Table structure (mutating — the same write safety and grid-consistency guard as every mutating tool)

All seven are addressed by the SAME `tableIndex`/`row`/cell indices `get_layout_info` reports — a
cell's physical `colIndex` for the cell tools (`merge_cells`, `split_cells`, `set_cell_borders`' `col`),
its grid `gridColumn` for the column tools (`insert_column`'s `atColumn`, `remove_column`'s `column`,
the position in `set_column_widths`' list). The two differ on any row with spanned cells, which is
nearly every real BC table: on the stock VAT-spec invoice, Line Amount is `colIndex` 7 but
`gridColumn` 9 — take
`layoutPart`=`"body"` (also `"header"`/`"footer"`) + `partName`=`null`, and pass the shared
grid-consistency backstop that rejects (`edit_would_break_table`, file untouched) any edit desyncing
a row from its `w:tblGrid`. Rows that skip grid columns (`w:gridBefore`/`w:gridAfter` — real BC
line tables have them) are handled; tables using **vertical merges (`w:vMerge`) are rejected** by
every tool here except `set_cell_borders`. Every one of them — and `insert_table` — returns the
**table-edit result**: `operation`, `part`, `tableIndex`, `columnIndex` (the grid column added or
removed, or the grid column an edited cell starts at; absent for a whole-table or whole-row edit),
`rowsAffected`, `columnCountBefore`, `columnCountAfter`, `summary`, and post-edit `quickValidation`.

| Tool | Key params (defaults) | Returns | Reach for it when… |
|---|---|---|---|
| `insert_table_row` | `layoutPath`, `tableIndex`, `cells` (comma-separated specs laid on the table's grid: `-` = empty spacer, a **FULL dataset path** = one bound cell — unlike `insert_repeater_row`'s leaf names, because a totals row binds whatever the dataset provides, typically a non-repeating `Totals` item; label-shaped paths become labels; `a+b` chains paths in one cell; optional `N:` span prefix (N 1–63) — spans must sum to `columnCount`), `atRow`=`null` (0-based row position in `tables[]`' one-row-per-`w:tr` numbering; a position BETWEEN two rows of the same repeater is refused — a static row cannot sit inside a repeating item — so use the repeater's first row index or the index after its last; omit to append after the last row, where a totals row goes), `alignments`=`null` (per-cell, `-` to skip), `bold`=`null`, `fontSizePoints`=`null` | The table-edit result (`rowsAffected` 1; the new bound controls' ids are in `summary`) | The stock totals block INSIDE the line-items table — every stock document layout ends its lines table with static right-anchored rows (e.g. `cells="-,-,-,-,3:/Header/Totals/TotalIncludingVATText,/Header/Totals/TotalAmountIncludingVAT"`, `alignments="-,-,-,-,-,right"`, `bold=true`), or a spacer row (`cells="8:-"`). The row renders exactly ONCE, never per data row. Follow with `set_cell_borders` for the rule and `set_cell_text` for a literal caption. |
| `insert_column` | `layoutPath`, `tableIndex`, `mode` = `"field"` \| `"label"` \| `"plainText"`, `dataPath` (for `field`/`label`: full dataset path, exactly as `insert_field`/`insert_label` take), `headerText`=`null` (omit to humanize the `dataPath` leaf), `headerLabelPath`=`null` (bind the header cell to a label column instead of static text), `atColumn`=`null` (0-based GRID position — a cell's `gridColumn` puts the column before it, `gridColumn`+`gridSpan` after it; omit to append at the far right), `width`=`null` (the NEW column's twips, taken proportionally from the existing columns — the table keeps its total width; omit for an equal share) | The table-edit result, `columnIndex` being the new column's grid index. It returns no `controlId` for a control it creates: call `get_layout_info` to address one | Adding a column to a line-items or plain table. Bound control lands only in the rows of the repeater whose data item contains `dataPath` (a `/Header/Line` field never goes into the same table's `ReportTotalsLine` row; a parent/Header field only in a single-repeater table, nested detail repeaters aside; no match is refused), and in ONE row per repeater item — the first of the item's rows already holding a field of `dataPath`'s data item, else its first row; the item's other rows get filler; header rows get a header cell — `w:tblHeader` rows, else a lines table's first row, else a plain table's caption (all-label) rows — and `headerText`/`headerLabelPath` on a table with none are refused, not dropped; right-anchored totals-block rows are widened, not celled, so the block keeps its look. An interior `atColumn` that falls INSIDE a spanned content cell is refused, naming the cell (`split_cells` there first). |
| `remove_column` | `layoutPath`, `tableIndex`, `column` (0-based GRID index — the `gridColumn` of the column's cells, never their `colIndex`) | The table-edit result | "Remove the GST Amount column from the lines." The summary says per row whether a cell was deleted, a spanned cell narrowed or a skipped run shrunk, and warns when no row had a cell of its own there (the sign of a wrong index). Drops each row's covering cell — including a bound one, whose binding goes with it (UNLIKE `remove_control`, which preserves the cell) — or decrements its `gridSpan`; removes the `w:gridCol`; redistributes the width proportionally (follow with `set_column_widths` for a different distribution). The last remaining column cannot be removed, nor a column that is some row's ONLY cell (a `w:gridBefore`/`w:gridAfter` row — Word cannot open a row with no cells); a table none of whose rows matches its `w:tblGrid` is refused too. |
| `set_column_widths` | `layoutPath`, `tableIndex`, `widths` (comma-separated twips, each 1–31680, exactly one per GRID column) | The table-edit result | Resizing columns; `gridSpan`-aware — each cell becomes the sum of the columns it spans. |
| `merge_cells` | `layoutPath`, `tableIndex`, `row`, `fromColumn`, `toColumn` (0-based PHYSICAL cell indices — `colIndex` — inclusive) | The table-edit result | Horizontally merging a run of adjacent cells in ONE row into a single `gridSpan` cell — the first cell (content and binding) is kept and widened, the rest are deleted. Refused if an absorbed cell holds a bound control (`remove_control` it first). HORIZONTAL only. |
| `split_cells` | `layoutPath`, `tableIndex`, `row`, `cellIndex` (0-based PHYSICAL index — `colIndex` — of the spanned cell) | The table-edit result | Splitting one spanned cell back into single-column cells (content stays in the first). Refused if the cell is not spanned. |
| `set_cell_borders` | `layoutPath`, `tableIndex`, `row`, `edges` (comma-separated `top`/`bottom`/`left`/`right`, or `all`), `col`=`null` (omit = every cell in the row — the usual case), `style`=`"single"` (draws) \| `"none"` (explicitly clears), `size`=`4` (eighths of a point, 2–96; 4 = ½ pt, the only thickness the BC corpus uses) | The table-edit result | The rules a BC document's look is actually made of: "a line above the totals row", "underline the grand-total cell". Edges you do not name are left as they were. Cosmetic only — works even on `w:vMerge` tables. |

### Lifecycle

| Tool | Key params (defaults) | Returns | Reach for it when… |
|---|---|---|---|
| `create_layout` | `schemaSource` (existing `.docx` layout **or** schema `.xml`), `outputPath`, `templatePath`=`null` (an **unbound** branded/styled shell — headers/footers/logo/fonts/styles, NOT a full BC layout with its own bound controls; a `.docx`, or a `.dotx` template, which is converted to a document — a macro-enabled `.dotm`/`.docm` is refused with `invalid_argument`), `headingText`=`null` (the heading paragraph a BLANK build's body starts with — omit for the report's own name, pass a human document title when authoring from scratch, or `""` for no heading; ignored when a template's body already has content) | `outputPath`, report identity, `storeItemId`, `usedTemplate`, `replacedExistingBcPart`, `quickValidation` | Starting a brand-new layout, optionally over a branded template — §1 ("Where the schema comes from") covers how to obtain the `schemaSource` artifact when none exists yet. Always ships exactly one BC custom XML part (fresh `storeItemID`) plus the glossary part the `insert_*` placeholders depend on. A BLANK (non-template) layout also ships empty header/footer parts wired into its page setup, **a default styles part pinning its typography** (Calibri 11 pt `docDefaults` plus the standard `Normal`/`TableGrid` definitions), so it renders with the same font in Word and in Business Central, **and a settings part declaring `compatibilityMode` 15** — the mode every stock layout declares and the one in which repeating-section controls exist, so a Word round trip cannot silently convert its repeaters to plain rich-text controls — a template keeps its own headers/footers, styles/theme and settings untouched (`validate_layout`'s `compatibility-mode` warning covers the shell that leaves the layout below 15). `templatePath` MUST be unbound: if it already carries a BC part **and** bound controls that would go stale against the fresh `storeItemID`, the call fails outright with `template_not_unbound` (nothing is written) instead of silently shipping a broken layout — see §5. A template whose BC part has zero bound controls of its own still succeeds. `outputPath` may name `schemaSource` itself only while that file has no bound controls (scaffolding a just-compiled layout in place); over an authored layout it is refused with `invalid_argument` and nothing is written, because the result would keep only the dataset part — use `refresh_xml_part` to re-bind an existing layout to a new schema. |
| `refresh_xml_part` | `layoutPath`, `newSchemaSource` (existing `.docx` **or** schema `.xml`) | Old/new report identity, `namespaceChanged`, `remappedCount`, `orphanedBindings[]` (`alias`, `xPath`, `part`, `sdtId` — the control's id, ready for `remove_control`), `repointedForeignBindings[]` (`alias`, `xPath`, `part`, `sdtId`, `previousNamespace`), `newUnboundFields[]`, `quickValidation`, `repairedUnreadablePart` | After the AL report dataset changes — and to repair a dataset part that no longer parses (`validate_layout` reports it as `schema-parse`, naming the part and the parser error): this is the one tool that can replace such a part, keeping its `storeItemID`; `repairedUnreadablePart` then describes what was replaced and `newUnboundFields` is empty (there was no readable old schema to diff). Replaces the dataset part's content **in place** (keeps the same `storeItemID`, so every existing binding still links to the same part) and reclassifies every existing binding by element name — matches are remapped/kept, non-matches are reported as `orphanedBindings` and **left in place** (this tool never deletes or rebinds anything itself — that's your call via `remove_control`/`insert_field`/`insert_label`). It DOES re-point namespaces: every BC-namespaced binding's `prefixMappings` is rewritten to the layout's namespace, always — including a binding left naming a **foreign** report's namespace, which BC rejects at upload (`InvalidPrefixMapping`) and which nothing else here can repair; those are reported as `repointedForeignBindings`. A re-pointed mapping written as a bare URI comes back as an `xmlns:ns0='…'` declaration, the form Word maps. A non-zero `errorCount` here from `xpath-resolves` findings is the *expected* corroboration of the orphan report, not a failed refresh. |

## 3. Supported matrix (v1)

**Works:**
- Field and label content controls (plain-text SDTs), in the document **body, headers, and footers
  alike** (`insert_field`/`insert_label` both take `layoutPart`).
- Repeater **TABLES in the document body** (`insert_repeater_table`: header row + one data row wrapped
  in a repeating-section control, per-column bound fields).
- Repeaters nested **one or more levels deep**: reading, validating, merging, and previewing a
  pre-existing multi-level repeater structure works end to end (proven against the real corpus, up to
  5 levels deep, via per-level XPath re-anchoring). Creation: `insert_repeater_table` builds one flat
  repeater per call, `insert_repeater_row` adds a nested DETAIL-ROW repeater inside an existing
  repeater's item — the standard BC per-line-detail shape — and `insert_subtotal_row` appends the
  static per-group spacer/subtotal rows that close each group in Microsoft's grouped list reports.
- Static text (`insert_text`, with optional `bold`/`fontSizePoints`), plain unbound tables
  (`insert_table`), and picture controls (`insert_picture`, mm-sized); existing and inserted picture
  controls get a placeholder image blip during merge (counted in `stats.picturesFilled`).
- Page-number field codes (`insert_page_number`): the stock `PAGE`/`NUMPAGES` "Page X / Y" header
  chrome, emitted exactly as the corpus layouts carry it. (Page-position-CONDITIONAL content —
  `IF PAGE = NUMPAGES` constructs — remains deferred, GitHub issue #11.)
- **Table-structure changes**: `insert_table_row` (a static totals/spacer row inside an existing
  table — the stock BC totals-block shape), `insert_column` / `remove_column` / `set_column_widths` /
  `merge_cells` / `split_cells` (all `gridSpan`-aware and grid-column-addressed, guarded against
  desyncing a row from its `w:tblGrid`; rows skipping grid columns via `w:gridBefore`/`w:gridAfter`
  are handled) and `set_cell_borders` for the per-cell rules BC documents get their look from.
- Inline page images of the preview PDF (`render_preview_pages`), so you can visually inspect your
  own work without an external viewer — and, via its `outputDir`, the same pages written to disk as
  PNGs for a human to open.
- `create_layout` (blank or from an unbound branded template) and `refresh_xml_part` (in-place dataset
  swap + binding reclassification) for layout lifecycle.
- Both validation levels: `quick` (structural/binding) and `full` (dry-run merge).
- Offline PDF preview with a choice of converter: Word COM (primary) or LibreOffice (fallback), via
  `converter=auto|word|libreoffice`.

**NOT supported:**
- **Hide-if-empty / BC add-in conditional controls** (Hide Field if Zero, Hide Empty Table, Hide Empty
  Table Row, layout comment) — there is no `set_hide_if_empty` tool. *Hide Empty Table Row* and *Hide
  Empty Table* are captured in two corpus layouts, and the table tools address cells inside their
  wrapper, but authoring them waits on what BC treats as empty (GitHub issue #123). *Hide Field if
  Zero* and layout comments have never been captured (#8).
- **Repeater tables in headers/footers** — not supported: nothing shows BC renders a repeating section there. `insert_repeater_table` and `insert_table`
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
  `layoutPart='body'`; this is a deliberate limit (nothing shows BC renders a repeating section in a header/footer), not something to work around by
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
  you. It DOES repair one class on its own: a binding naming a foreign report's namespace, reported
  back as `repointedForeignBindings`, because BC refuses to upload a layout carrying one.
- **Don't rebuild controls by hand to fix a `binding-namespace` error.** One `refresh_xml_part`
  against the report's current schema re-points every such binding — including on a layout whose
  namespace is otherwise unchanged. Rebuilding is only needed when the same control is ALSO in
  `orphanedBindings` (its XPath names a column this report does not have).
- **Supply a real dataset via `dataOverridesPath`** for a `preview_layout` call that needs to look
  realistic — generated sample data is type-aware but still synthetic. An export printed for several
  documents previews the first only (a `multiple-top-level-records` warning names the count, since BC
  would render the body once per record): to look at another document, pass a dataset holding just that
  one.
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
{ "ok": true,  "data": { <tool-specific payload> } }
{ "ok": false, "error": { "code", "message", "hint" } }
```

Members that would be `null` are **omitted, never sent as `null`**: a success carries no `error` key,
a failure no `data` key, and an optional field inside `data` (an unreferenced part's
`partDetails[].role`, a failed conversion's `pdfPath`, a control's `tableIndex` outside any table) is
simply absent. Field names are camelCase as shown throughout §2 (`xPath`, `storeItemId`, `sdtId`).

- `ok:false` always comes with a populated `error`. `code` is one of `file_not_found`,
  `invalid_layout` (a layout, template or schema file you passed is not usable as one: not a `.docx`
  package at all — e.g. a schema `.xml` saved under a `.docx` name, or an empty or truncated file —
  not well-formed XML, a damaged package part, or no BC dataset part; the hint names the file and,
  for XML, the line and position), `invalid_argument` (including a `dataOverridesPath` file that is
  not one of the two dataset shapes BC produces for this report), `not_found`,
  `edit_would_corrupt` (a mutating-tool structural-safety rejection — the file is left untouched),
  `edit_would_break_table` (the
  table-structure tools' grid-consistency rejection — a corruption class the OpenXML validator
  accepts silently; file untouched), `file_locked` (the file the call reads or must replace is busy
  or not writable: open in Word or another program, read-only — e.g. a TFVC/Perforce file that is
  not checked out — or not writable by your account; the hint names the cause to check, and the file
  is untouched. A read-only layout can still be inspected, validated and previewed; only writing it
  is refused), `template_not_unbound` (`create_layout`-only: `templatePath` was a full BC
  layout, not an unbound shell — see §"Lifecycle" above; nothing is written), `rasterize_failed`
  (`render_preview_pages`-only: the file is not a PDF or is too large, PDFium cannot read it, or
  `firstPage` is past the last page — the message gives the page count when known, so page back;
  for anything else run `preview_layout` again and use its fresh `pdfPath`), or `internal_error`
  (a bug in the server, not in your call — including `create_layout`'s own check of the layout it
  just built failing; nothing was written, and the message is worth reporting).
  `hint` is always non-empty and actionable: which argument to fix and its valid values, or which
  inspection tool to call. Act on it before retrying blindly.
- A malformed call gets the same envelope (with `isError: true`), never bare SDK text: an unknown
  tool name, a missing required argument, a wrong JSON type (a quoted number, `"true"` for a
  boolean, `1.5` for an integer), an argument name the tool does not have (names are
  case-sensitive — `layoutpart` is refused, with `layoutPart` suggested), or a relative path is
  `invalid_argument` naming the tool or parameter. To use an
  optional argument's default, omit it: `null` is accepted only where the schema type allows it, and
  for an optional string, which then takes its default.
- `ok:true` doesn't mean "nothing to look at" — check the tool's own payload for its validation
  signal:
  - All 20 mutating tools (the Edit and Table structure tables in §2, plus `refresh_xml_part`):
    `data.quickValidation.passed` / `errorCount` / `warningCount`.
  - `create_layout`: `data.quickValidation` — a template whose own pre-existing bound controls would
    go stale is refused outright (`template_not_unbound`, see above) rather than reported as data here.
  - `refresh_xml_part`: `data.orphanedBindings` and `data.newUnboundFields` in addition to
    `quickValidation` — a non-zero `errorCount` here from `xpath-resolves` findings is expected
    corroboration of the orphan report, not a failure.
  - `preview_layout`: `data.conversionOk` / `data.conversionError` (PDF conversion can fail or be
    unavailable even though the merge itself succeeded; `pdfPath` is then absent and no PDF is left on
    disk, not even an earlier preview's, so never fall back to a `pdfPath` from a previous call), plus
    `data.warnings[]` and `data.quickValidation`.
- Mutating tools apply a **structural gate only**: a write is rejected solely if it would introduce a
  *new* OpenXML structural error versus the file's own pre-edit baseline (`edit_would_corrupt`), or a
  new table row out of step with its grid (`edit_would_break_table` — every one of the 20 runs this
  check, not only the table tools). Semantic issues (like an orphaned binding) are always allowed
  through and reported as validation findings, never blocked.
