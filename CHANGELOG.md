# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- **New `validate_layout` check `top-level-repeater-group` (error).** A block-level repeating
  section bound directly to the top-level data item (e.g. the whole invoice body wrapped in a
  repeater over `/Header`) makes Business Central print an EMPTY document unless its
  `repeatingSectionItem` sits inside a `w:group` content control — verified on report 1306 with
  three otherwise identical layouts (no outer repeater: prints; repeater without group: empty page;
  repeater with group: prints). Quick and full validation and `preview_layout` all passed the broken
  shape before, so the first symptom was a blank printout. The check is scoped so no captured layout
  trips it: StandardStatement's body-wide `/Customer` repeater carries the group, list reports'
  top-level repeaters are row-level, and the invoice/quote block-level repeaters are over nested
  items. No tool emits a block-level repeater, so this guards layouts written or edited outside the
  tools.

- **New `validate_layout` check `duplicate-control-id` (warning)** (GitHub issue #101). Two content
  controls sharing a `w:sdtPr/w:id`, which only a hand edit at the package level produces, passed
  validation silently while every id-addressed tool acted on the first of them. The check names
  both controls, and the id-addressed tools now refuse such an id.

- **New merge warning `multiple-top-level-records`** (GitHub issue #105). A `dataOverridesPath`
  dataset printed for several documents, such as two purchase orders, previewed only the first
  and said nothing. `preview_layout` now raises one warning per top-level data item the dataset
  holds more than once and the layout does not repeat itself, naming the count ("The dataset has 2
  top-level 'Purchase_Header' records, but the preview renders the first only…"). Items a repeater
  repeats, such as a list report's rows, are exempt. The body is still rendered once: there is no
  observed OOXML for how BC joins per-record documents.

### Fixed

- **Inserts that would land inside a data-bound field, label or picture control are refused**
  (GitHub issue #76). Three locations could resolve inside a bound control's content:
  - `afterControl` on a cell-level field;
  - `tableCell` on a cell that is itself a control;
  - `atText` matching such a control's text.

  The insert was written in there and reported as success. Word then either refused to open the
  file (page numbers or a table inside a plain-text field) or silently dropped the inserted content
  on its next save (notably in `w15:dataBinding` fields). Every insert tool now refuses these
  locations with `invalid_argument`, and the hint points at `tableCell` on a plain cell or
  `afterControl` on an inline control. Cells of repeater rows and inline controls are unaffected.
  The write gate also rejects field codes, tables and drawings inside a plain-text control as
  `edit_would_corrupt`. One existing refusal changes code: `afterControl` on a plain-text
  cell-level label now fails as `invalid_argument`, before any edit, instead of
  `edit_would_corrupt`.

- **A header or footer an insert creates is now referenced from the first section** (GitHub
  issue #91). Only the last section referenced it, so in a multi-section layout such as JobQuote the
  earlier sections, page 1 included, rendered without it. The first insert into such a new header
  or footer also no longer leaves an empty line above its content.

- **`create_layout` with a Word template (`.dotx`) as `templatePath` now produces a `.docx` Word
  can open** (GitHub issue #77). The template was copied as-is, so the output kept the template's
  main-part content type under a `.docx` name. Word refused it ("Word was unable to read this
  document"), while `create_layout`'s own quick validation and `validate_layout` reported it clean.
  - A template is now converted to a document before the BC part is attached.
  - A macro-enabled `templatePath` (`.dotm`/`.docm`) is refused with `invalid_argument` instead
    of having its macros silently dropped.
  - A new `validate_layout` check, `document-content-type` (error), flags any layout whose main
    part is not typed as a plain Word document, such as a template renamed to `.docx`.

- **`validate_layout` now catches bindings whose `w:prefixMappings` cannot resolve their XPath**
  (GitHub issue #79). These mappings validated clean although Word cannot map the control:
  - a namespace in the `dynamics-365/report` form;
  - the singular `nav/report` form;
  - `ns1` declared while the XPath uses `ns0`.

  This was possible because `binding-namespace` only looked at namespaces that resemble BC ones,
  and `xpath-resolves` ignores prefixes. The new `binding-prefix-mapping` check (error) requires
  every prefix a dataset binding's XPath uses to be declared as the layout's own namespace. A
  missing attribute, or a bare URI naming the right namespace (a shape stock layouts ship), draws
  no finding. `refresh_xml_part` now writes a bare mapping it re-points as an `xmlns:ns0='…'`
  declaration; before, it stayed bare and Word still reported the control as unmapped.

- **`insert_repeater_table` without `columnWidths` no longer runs off the page** (GitHub issue
  #84). Every column got a fixed 2000 twips, and now that the grid is pinned, tables of six or more
  columns ran past the margin. The columns now split the document's content width evenly, as
  `insert_table` does; a table placed inside a cell splits that cell's width instead.

- **`set_cell_text` and `clear_cell_text` no longer delete a nested table, picture or field in the
  target cell** (GitHub issue #86). Deleting a nested table also shifted every later `tableIndex`
  in that part. Only the cell's text runs are now replaced. Other content stays in place and the
  summary names it, and `previousText` reports only the text actually replaced.

- **`remove_control` with `keepText=true` is refused on the `w:group` control between a
  block-level top-level repeater and its item** (GitHub issue #87). On StandardStatement's
  `/Customer` group it returned ok and left a layout Business Central prints as an empty document.
  The hint points to removing the repeater itself with `keepText=false`.

- **`remove_column` no longer leaves a table row with no cells** (GitHub issue #78). Removing the
  only cell of a row whose other columns are skipped (`w:gridBefore`/`w:gridAfter`) returned
  `ok:true` and left a file Word cannot open. It is now refused, naming the row, and the table-grid
  safety check rejects any edit that leaves a row with no cells. `remove_column` also refuses a
  table none of whose rows matches its `w:tblGrid`; it used to drop a grid column there and report
  "updating 0 row(s)".

- **`insert_column` binds the new field only into the rows it belongs to** (GitHub issue #82).
  - **Repeater tables:** on the stock quote/invoice lines tables, a `/Header/Line/...` field was
    also bound into the same table's `/Header/ReportTotalsLine` rows. Line 1's value showed beside
    every totals line, and the right-anchored totals shifted. The field now goes only into the rows
    of the repeater whose data item contains `dataPath`, and the other repeaters' rows keep their
    totals layout. A `dataPath` that belongs to none of the table's repeaters is refused.
  - **Plain tables:** the header cell now goes into the caption rows (rows of label controls, the
    BC caption-over-value shape), and the field only into the value rows. `headerText` or
    `headerLabelPath` on a table with no header row is refused instead of silently dropped.
  - **Validation:** `headerLabelPath` is now checked against the dataset in every mode, including
    `plainText`.

- **`insert_column` keeps the table's width** (GitHub issue #83). The new column was added on top
  of the existing ones, so a full-width table such as the stock invoice's lines table grew past
  the page margin and pushed its grand total off the page. In tables whose cells are sized in
  percent (the quote's info block, footers), the new column was squeezed to a sliver.
  - The table now keeps its total width: the new column's width is taken proportionally from the
    existing columns.
  - Every row's cell widths are recomputed from the new grid, as `remove_column` already did.
  - `width` is now the new column's width, taken from the others; it must be smaller than the
    table. Omit it for an equal share.

- **Edits of a layout that is open in Word or read-only fail as `file_locked`, and a read-only
  layout can be previewed, validated and used as a template** (GitHub issue #94).
  - **Wrong error code:** with the layout open in Word, or read-only (a TFVC/Perforce file that is
    not checked out), every edit failed as `internal_error` with a "likely a bug in the tool" hint.
    Word still lets the file be read and copied, so the failure only appeared when the edit
    replaced the file. The same happened for `create_layout` over such an `outputPath`, and for
    `preview_layout` while the previous `merged.docx` was open in Word. All of these now fail as
    `file_locked`. The hint names the three likely causes: the file is open in Word or another
    program, it is read-only or not checked out, or you have no write permission. The file is left
    byte-identical.
  - **Read-only sources:** private copies inherited the read-only attribute. So `preview_layout`
    and `validate_layout level=full` failed on a read-only layout, `create_layout` failed on a
    read-only `templatePath`, and each failure left behind a read-only copy that nothing could
    delete. Private copies are now always writable, and a `create_layout` output is never
    read-only.
  - **Leftover temp files:** `.bcwl-build-*` files next to a `create_layout` output and
    `%TEMP%\bcwl-edit-*` working copies were never cleaned up after a crash. Both are now removed
    once they are a day old, like the other temporary files. This includes read-only leftovers
    from earlier versions.

- **A file that is not a usable layout, schema or dataset is reported as the caller's input, by
  name** (GitHub issue #95). These failed as `internal_error` ("likely a bug in the tool"), or with
  a hint about the wrong file:
  - **Not a `.docx` at all:** an empty file, a text file or schema `.xml` saved as `.docx`, or a
    truncated copy, passed to any tool that opens a layout or template. This is now
    `invalid_layout`, and the hint names the file and explains `.docx` versus schema `.xml`. The
    edit and preview tools used to name their private temp copy; they now name your file.
  - **XML that is not well-formed:** a truncated schema `.xml`, one with a DTD, or one whose
    encoding declaration does not match its bytes, passed as `schemaSource`, `newSchemaSource` or
    the `source` of `list_dataset_fields`. A `.docx` with a damaged XML part is covered too. This
    is now `invalid_layout`, and the hint names the file plus the line and position.
  - **`dataOverridesPath`:** any file that cannot be used now fails as `invalid_argument` for that
    parameter. That covers malformed XML, a wrong root element or namespace, and an export from
    another report. The hint lists the two accepted shapes. A wrong-shaped file used to be
    `invalid_layout`, with a hint about the layout.

- **`create_layout` refuses to overwrite its own `schemaSource` when that file is an authored
  layout** (GitHub issue #89). The new layout keeps only the dataset part of `schemaSource`. So an
  `outputPath` equal to `schemaSource` replaced the authored layout with a blank one, and the call
  returned `ok:true`: on the stock sales quote, 68 controls became 0. This is now refused with
  `invalid_argument` and nothing is written, however the path is spelled. The hint points at
  `refresh_xml_part` to re-bind a layout in place, or at a different `outputPath`. Running
  `create_layout` in place on a just-compiled layout, which has no controls yet, still works.

- **Malformed calls get an accurate error code, and an explicit `null` no longer crashes a tool**
  (GitHub issue #97).
  - **Wrong JSON type:** `"abc"` or `1.5` for an integer, `"true"` for a boolean, or `null` for a
    number or boolean came back as `internal_error` ("The JSON value could not be converted to
    System.Int32…") with no parameter named. It is now `invalid_argument` naming the parameter, the
    type it needs, and what the call sent.
  - **Unknown tool name:** was an `internal_error` envelope; now `invalid_argument` naming the
    tool, with a hint pointing at `tools/list`.
  - **Missing required argument:** still `invalid_argument`; the message now names just the
    missing parameters, and the hint lists every required one.
  - **Explicit `null` for an enum-like string:** crashed the tool body (`internal_error`, "Object
    reference not set to an instance of an object."). An optional one (`layoutPart`, `look`,
    `style`, `converter`, `level`) now takes its documented default, as if omitted; a required one
    (`locationType`, `mode`) is `invalid_argument`.
  - **No more `invalid_request`:** the one path that emitted this undocumented code (a binding
    failure the SDK swallowed into its own text) now reports `invalid_argument`.
  - **Docs:** README, ADR-0004 and SOLUTION-DESIGN §4 said a binding failure returns the SDK's bare,
    non-enveloped text. They now describe what the server actually returns, and a host-level test
    that runs the real server pipeline pins every case above.

- **Unknown argument names and relative paths are refused instead of quietly changing the
  outcome** (GitHub issue #90).
  - **Unknown argument names:** the SDK ignored any key it had no parameter for, so a misspelled
    optional argument took its default. `insert_text` with `"layoutpart":"footer"` wrote into the
    body, and `set_cell_borders` with `"column":1` drew every cell. A key outside the tool's input
    schema is now `invalid_argument`, naming the closest real parameter (`layoutpart` →
    `layoutPart`, `column` → `col`) and listing them all. An unknown tool name gets the same
    "did you mean" suggestion.
  - **Relative paths:** a relative path, or a drive- or root-relative one (`C:Report.docx`,
    `\Report.docx`), resolved against the server's working directory, and `create_layout` wrote
    its output there. Every path parameter (`layoutPath`, `source`, `schemaSource`,
    `newSchemaSource`, `templatePath`, `outputPath`, `dataOverridesPath`, `outputDir`, `pdfPath`)
    must now be fully qualified, or the call is `invalid_argument` naming it.

- **Numeric arguments are range-checked where they are parsed** (GitHub issue #98).
  - **Column widths** (`set_column_widths`' `widths`, `insert_table`'s and
    `insert_repeater_table`'s `columnWidths`, `insert_column`'s `width`) must be 1–31680 twips
    (22 in, Word's maximum page width). A negative width failed as `edit_would_corrupt`, whose
    hint blames the tool; a zero width was written as `w:gridCol w:w="0"`.
  - **Cell spans** in the `cells` syntax (`insert_table_row`, `insert_repeater_row`,
    `insert_subtotal_row`) must be 1–63, Word's maximum table column count. `"3:-,-1:-"` crashed
    with "Index was out of range"; the refusal now has a hint explaining the `cells` syntax.
  - **`preview_layout`'s `rows`** below 1 is `invalid_argument`. It used to return `ok:true` with
    an empty body, plus warnings caused only by the row count.
  - **`render_preview_pages`' `firstPage` and `maxPages`** below 1 are `invalid_argument`. Before,
    page 1 rendered while the response echoed `firstPage: 0`.
  - **Hints for parameter names that mean different things on different tools** now depend on the
    tool as well as the name. `insert_table`'s `columns` (a count) no longer gets the hint for
    `insert_repeater_table`'s list of column names, and its `rows` no longer gets the generic hint.

- **Cells wrapped in two cell-level content controls are visible to the table tools** (GitHub
  issue #85). The Business Central Word add-in's "Hide Table Row If Empty" and "Hide Table If
  Empty" controls wrap a bound field's own cell-level control (`w:tr > w:sdt > w:sdt > w:tc`), as in
  the stock SalespersonCommission and SubcontractorDispatchList layouts. The tools read one level
  only, so they found no cell there:
  - **`get_layout_info`** reported an empty "Unbound" cell, and the bound field had no table
    coordinates. The cell now reports the bound field as its control, with its text, and every
    wrapper level gets the cell's coordinate.
  - **Rules, widths and spans:** `set_cell_borders` left a gap under the cell, `set_column_widths`
    left its width stale, and a wrapped `gridSpan`-4 cell counted as one grid column.
  - **Text and inserts:** `set_cell_text`, `tableCell` inserts and `afterControl` on either level
    answered `not_found`. They now refuse the cell as a bound control (`invalid_argument`), like
    any single-level bound cell.

- **`get_layout_info` reports the grid column of every cell** (GitHub issue #81). Cells reported
  only `colIndex`, their physical position in the row, while `remove_column`, `insert_column`'s
  `atColumn` and `set_column_widths` take a grid column. On spanned tables the two differ: on the
  stock VAT-spec invoice's lines table, Line Amount is `colIndex` 7 but grid column 9, and
  `remove_column` 7 narrowed the VAT % cells instead and reported success.
  - **New fields:** each cell now carries `gridColumn` (where it starts, `gridBefore` included)
    and `gridSpan`, and each row `gridBefore`/`gridAfter`. The column tools' descriptions and hints
    point at `gridColumn`.
  - **`remove_column`'s summary** says per row whether a cell was deleted, a spanned cell narrowed
    or a skipped run shrunk, and warns when no row had a cell of its own at that column.
  - **`set_cell_borders`** returned the physical `col` as `columnIndex`. It now returns the grid
    column the cell starts at, as the result type documents and `merge_cells`/`split_cells` do.

- **Column and width edits reach every row of a repeater item** (GitHub issue #80). The table tools
  saw only the first `w:tr` of each repeater: an item's other rows, nested detail repeaters, and the
  subtotal and spacer rows `insert_subtotal_row` adds were skipped by `insert_column`,
  `remove_column`, `set_column_widths`, the `w:vMerge` refusal, `get_layout_info`, `tableCell`
  addressing and the grid-consistency check. A column edit left those rows on the old grid (on
  InventoryOrderDetails the removed amount still printed past the table's right edge), and
  validation passed.
  - **Row numbering changes:** `get_layout_info`'s `rows[]` now lists every `w:tr`, each with a new
    `repeaterControlId` (the innermost repeater it repeats with). On tables with multi-row repeater
    items `rowCount` grows and the rows after such an item get higher indexes, so re-read
    `get_layout_info` before reusing a stored row index. Controls in nested rows now have table
    coordinates.
  - **`insert_column`** binds a field into one row per repeater item: the first of the item's rows
    that already holds a field of the same data item, else its first row.
  - **`insert_table_row`** refuses an `atRow` between two rows of the same repeater
    (`invalid_argument`) and names the positions just before and after it.

- **A BC dataset part that no longer parses is reported as such, and `refresh_xml_part` repairs it**
  (GitHub issue #96).
  - **Diagnosis:** a part recognisable by its itemProps schemaRef, or by the BC namespace in its raw
    bytes, used to be reported as "no BC part" by every tool, with the parser's error discarded.
    `validate_layout` now reports one `schema-parse` error naming the part, the parser's message and
    any encoding declaration that contradicts the bytes or DOCTYPE it carries; `get_layout_info`,
    the insert tools and `create_layout`'s `schemaSource` fail `invalid_layout` with a repair hint.
  - **Repair:** `refresh_xml_part` replaces such a part's content, keeping its `storeItemID`, and
    reports it as `repairedUnreadablePart` (`newUnboundFields` is then empty).

- **Layouts with two BC dataset parts, or two controls sharing an id, are no longer resolved by
  "first match"** (GitHub issue #101).
  - **Two BC parts:** reading uses the part most bindings name by `storeItemID`, then by namespace,
    then the lowest package path, never relationship order; `single-bc-part` lists the parts and
    says which one was measured. `refresh_xml_part`, `create_layout`'s `schemaSource` and every
    tool that writes a binding refuse such a layout with `invalid_layout`.
  - **Shared control ids:** `remove_control`, `afterControl` locations and `parentRepeaterId` refuse
    an id more than one control carries instead of acting on the first.

- **Bound inserts work on layouts whose dataset part has no itemProps** (GitHub issue #93).
  - **Stock layouts:** on `SalespersonCommission`, `StandardStatement` and every other such layout,
    a new binding reuses the `storeItemID` all existing bindings agree on; no itemProps is minted.
  - **No consensus:** with no bindings, or disagreeing ones, the call fails `invalid_layout` with a
    hint saying what to do, instead of `invalid_argument` naming internal methods.

- **The mock merge re-anchors each binding through every enclosing repeater row** (GitHub issue
  #104). It tried only the innermost row and read anything else from the document root with every
  step indexed `[1]`, warning once per generated row:
  - **Outer-row fields:** a field in a nested row bound to an outer row's column showed the first
    outer row's value in every group. On InventoryOrderDetails, an `/Item/No_Item` field in the
    Sales_Line row printed item 1's number under every item. Fields and nested repeaters now
    resolve against the deepest enclosing row whose path they extend.
  - **Absolute bindings:** a binding that no enclosing row holds is read from the document root,
    as BC reads it, with no warning. SalespersonCommission's `/Labels` caption in the Salesperson
    row no longer raises three of them. A binding outside every repeater row is always silent.
  - **When it still warns:** `xpath-fallback` is raised once per control, with its evaluation
    count, in two cases:
    - a `//` binding;
    - a control inside a repeater row that no enclosing row re-anchors, which reads the first of
      several records of a repeating item, such as a `/Header/Line` field in a ReportTotalsLine
      row. It repeats with rows it does not belong to, and the warning names the item and the
      record count. Caption (label) columns are exempt, since they read the same in every record.

- **`preview_layout` never leaves a stale PDF behind, and default previews get distinct PNG names**
  (GitHub issue #106).
  - **Stale PDF:** a failed re-conversion (a Word timeout, LibreOffice missing) returned
    `conversionOk: false` but left the previous render on disk, so an earlier `pdfPath`, or the
    folder itself, showed the pre-edit layout. The previous PDF is now deleted before converting,
    and anything a failed attempt leaves there is removed. When the previous PDF cannot be deleted
    (a viewer has it open), the call skips conversion and says so in `conversionError`.
  - **PNG collisions:** every default preview was `preview.pdf`, so `render_preview_pages` wrote
    two layouts' pages into one `outputDir` as the same `preview-page01.png`. The default PDF is
    now `preview-<hash>.pdf`, as with an explicit `outputDir` (`merged.docx` keeps its name). A
    `preview.pdf` an earlier version left in a default folder is removed.

- **Edits made through a junction, a symlink, or a subst or mapped drive now wait for edits made
  through the plain path** (GitHub issue #92). Both edit locks, the in-process one and the named
  mutex shared between server processes, were keyed on the full path as written, which does not
  follow those spellings. Two hosts editing one layout through two spellings therefore never waited
  for each other, and one host's commit replaced the other's. Of 12 concurrent `insert_text` calls
  from two server processes, 6 through a folder and 6 through a junction or a subst drive to it, all
  returned `ok: true` and the file kept 6; it now keeps all 12. Every path is resolved through the
  file system before it keys a lock, and the same resolution:
  - commits an edit made through a symlink to the file into the file it points at, instead of
    replacing the link with a plain copy that holds the edit;
  - keys `preview_layout`'s default output folder, so every spelling of a layout shares one;
  - decides "the same file" for `create_layout`'s refusal to overwrite its own `schemaSource` and
    for the preview merge's output check, so a junction spelling of `schemaSource` is refused too.

- **`preview_layout` never kills a Word it did not start** (GitHub issue #103). The Word converter
  took the first WINWORD.EXE that appeared while it was creating `Word.Application` as its own, and
  force-killed it after every conversion, or at once on a timeout. When another server process
  converted at the same moment (a second IDE window, or the test suite next to an agent's preview),
  or the user opened Word in that window, it killed that Word instead, unsaved documents included,
  and could leave its own instance running.
  - **Positive identification:** the instance now gets a unique window caption, and only the
    process owning that hidden window, started after the conversion began, counts as its own. A
    Word that can't be identified that way is only asked to quit, never killed.
  - **Word hanging at startup** (an invisible first-run, licensing or safe-mode prompt): the
    timeout used to kill nothing and left that Word running. It now kills the one new WINWORD.EXE
    that COM started (its command line carries `-Embedding`, which a Word the user starts never
    does), and only when there is exactly one.
  - **No late export:** a conversion that has timed out stops before opening or exporting the
    document if Word recovers later, so it can no longer write a PDF the next preview may own.

- **`tools/scrub_corpus_metadata.py` no longer passes a `customXML/` capture as clean without
  scrubbing it** (GitHub issue #111). Its bibliography-drop and dataset sample-value rules matched
  the `customXml/` folder case-sensitively, so a capture whose export spells it `customXML/` (as
  `StandardStatement.docx` does) skipped both rules and was reported clean. The script now reads
  package plumbing the way OPC defines it:
  - **Part names** compare case-insensitively everywhere, and every part keeps the casing the
    package gives it.
  - **Relationship targets** resolve against the source part's folder, or from the package root
    when absolute, so a part referenced by an absolute target is no longer refused.
  - **Every bibliography part** is dropped in one run (only the first was), so a second run is a
    true no-op.
  - **`docProps/app.xml`:** `Company` and `Manager` are now blanked. A missing `core.xml` or
    `app.xml` no longer crashes the script.
  - **Exit codes:** `--check` exits 1 only when something would be scrubbed. Errors exit 2: an
    unknown option (previously ignored), or a package it cannot scrub cleanly, such as one with
    duplicate Overrides or a reference it cannot resolve. That file is left untouched, and the
    other files are still processed.
  - **Atomic rewrite:** the scrubbed package replaces the original in one rename. On Windows the
    move used to run while the original was still open and fell back to copying over it.

  The corpus was already clean and is unchanged. A standard-library `unittest` suite covers the
  rules and exit codes (`python -m unittest discover -s tools/tests`), and CI runs it.

- **The documented "Page X / Y" recipe now renders "Page  1 / 3"** (GitHub issue #88).
  `insert_page_number`'s description and both skills said to insert the two spaces after the
  `Page_Lbl` label first and the page-number fields second, both anchored on the label. An
  `afterControl` insert lands directly after its anchor, so that order put the spaces after the
  fields: "Page1 / 3  ". The recipe now inserts the fields first and the spaces second. The skill's
  location model states the general rule (an inline line is built right to left, every insert
  anchored on its first control), and `insert_text`'s description shows it for `Label: Value`. A
  test pins the run order.

### Security

- **A crafted layout can no longer kill the server through the OpenXml SDK's own parse** (GitHub
  issue #116). The SDK's part loads are recursive, so a ~1.4 KB `.docx` nested about 2,250
  elements deep overflowed the stack inside the SDK. A chain of about 1,500 parts did the same
  inside `OpenXmlValidator`. A stack overflow can't be caught, so the whole server process died.
  Every `.docx` is now opened through one gate that scans each part before the SDK parses it, and
  a package over a cap fails the call with `invalid_layout` instead. The caps are 512 levels of
  XML nesting, 16 MB per XML part, 64 MB of XML in total, and 256 parts. The deepest real layout
  nests 31 levels, the largest real part is 153 KB, and a real layout has about 30 parts.
  `MaxCustomXmlParts` drops from 1024 to 64 so it stays the more specific check under the new
  package-wide part cap.
- **`preview_layout` now removes every way its render copy could reach outside itself, not only
  external relationships** (GitHub issue #115). Before conversion it removes or unlinks:
  - Every field that isn't computed from the document itself, such as `INCLUDEPICTURE`,
    `INCLUDETEXT`, `LINK` or `DDE`. Each is unlinked to its cached result, so the preview looks
    the same. Fields whose type is hidden behind a nested field are unlinked too.
  - Mail-merge, update-fields-on-open and schema-library settings.
  - `altChunk` imports.
  - VML location attributes.

  Each removal is reported as a `field-unlinked` or `external-reference-stripped` warning. Plain
  hyperlinks and the self-contained fields layouts use (`PAGE`, `NUMPAGES`, `DATE`, `IF`, `REF`,
  …) stay live. No corpus layout loses anything. Observed before the fix: a layout with
  update-fields-on-open made Word show a "fields that may refer to other files" prompt on the
  desktop during a hidden preview, and confirming it made Word fetch each field's URL.
- **`render_preview_pages` caps each page image at 5000 px per side and 20 megapixels** (GitHub
  issue #117). Page size comes from the PDF, so the DPI cap alone let a page with a huge MediaBox
  request a multi-gigapixel bitmap. A page over the cap is now rendered smaller with its aspect
  ratio kept, and flagged `downscaled: true` on `pages[]`. Pages within the cap render exactly as
  before.

## [1.1.0] - 2026-08-23

### Added

- **New tool `insert_subtotal_row`** (GitHub issue #30): append a static per-group row at the END
  of an existing repeater's item — the stock BC shape for group subtotals, exactly as
  SalespersonCommission (115) closes each salesperson's group with a spacer row and a bold subtotal
  row INSIDE the `repeatingSectionItem`, after the nested detail rows. The row renders once per
  PARENT row (per group), never per detail row — verified against a real merge with a nested
  detail repeater in play. Cells use the same DSL as `insert_table_row` (full dataset paths — the
  corpus subtotal binds a sibling non-repeating `Subtotals` item and `Labels` captions); repeated
  calls stack rows in the corpus order (spacer, then subtotal). Together with
  `insert_repeater_table` (whose line row is the per-group header) and `insert_repeater_row`
  (nested details), the whole Microsoft grouped-list-report shape is now buildable from scratch.
- **New tool `insert_table_row`** (GitHub issue #28): insert ONE static (non-repeating) row into an
  existing table — the stock BC shape for the totals block INSIDE the line-items table, which every
  stock document layout ends with (right-anchored trailing rows: leading spacer cells, a
  `gridSpan`-merged bound caption cell, a right-aligned bound amount cell — corpus:
  StandardSalesQuote/StandardPurchaseOrder/StandardSalesInvoiceVatSpec) and which previously could
  only be approximated with a separate totals table. Cells use the `insert_repeater_row` DSL but
  bind FULL dataset paths (a totals row binds whatever the dataset provides, typically a
  non-repeating `Totals` item; label-shaped paths become label controls); `atRow` positions the row
  anywhere (append by default — where a totals row goes); optional per-cell alignments and
  `bold`/`fontSizePoints` for the stock grand-total look. The row is a direct child of the table,
  never part of a repeating section, so it renders exactly once — verified against a real merge.
  Works on body and header/footer tables; vMerge tables are rejected like every structure op.
- `get_layout_info` now reports each OOXML part's rendering facts alongside the flat `parts` list
  (GitHub issue #5): a `partDetails` section giving every part's `kind` (`document`/`header`/`footer`),
  its header/footer `role` (`default`/`first`/`even`, from the section's
  `w:headerReference`/`w:footerReference`; `null` when no section references it), and
  `isDefaultTarget` — the part a `layoutPart='header'/'footer'` location with no explicit `partName`
  actually resolves to, computed through the same selection logic the edit tools use. A new top-level
  `hasTitlePage` flag says whether the layout renders a distinct first page (`w:titlePg`), in which
  case page 1 shows the `first`-role parts instead of the default ones.
- The insert tools (`insert_field`, `insert_label`, `insert_text`, `insert_page_number`,
  `insert_picture`) now append a
  warning to their summary when content lands in the DEFAULT header/footer of a `w:titlePg` layout
  without an explicit `partName` — the "correctly inserted but invisible on a one-page render" trap
  that previously cost a sandbox round to diagnose. (A page number in the default header of a
  `w:titlePg` layout is the canonical instance of that trap.)
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
- **New `validate_layout` check `compatibility-mode`** (GitHub issue #51): warns when a layout that
  CONTAINS repeating sections declares a Word compatibility mode below 15 — the state in which a Word
  save silently converts them to plain rich-text controls. Warning, not error: Business Central
  merges the repeater regardless, so the file is correct until someone edits it in Word. Deliberately
  conditional on a repeater being present, so it stays silent on the field-only layouts (and stock
  captures) where the risk does not exist.
- **`render_preview_pages` can write the rendered pages to disk** (GitHub issue #55): the new optional
  `outputDir` writes each page as `<pdf-basename>-pageNN.png` and returns the paths on
  `pages[].path`, and `inlineImages=false` (only valid together with `outputDir`) writes them without
  spending response tokens on the inline copies. The images were previously returned as inline MCP
  content blocks only — visible to the calling agent but to nobody else, so "put the render somewhere
  I can look at it" could not be served by the server at all and ended in ad-hoc PDF rasterization
  outside it. Named by page NUMBER, so paging through a document with `firstPage` produces a set that
  sorts correctly; the directory's lifetime belongs to the caller, exactly like `preview_layout`'s own
  `outputDir`. Omitting both parameters leaves the response byte-for-byte as before.

- **New `validate_layout` check `repeater-downgraded`** (GitHub issue #54): an ERROR when a control's
  alias names a whole data item but the control is not a repeating section — what a Word save leaves
  behind when it converts a repeater to a plain rich-text control (it strips
  `w15:repeatingSection` and `w15:dataBinding` while `w:alias`/`w:tag`/`w:id` survive). Such a layout
  passed every other check: the field controls inside the row are intact, nothing is orphaned, and the
  table simply renders one row instead of one per record. The alias is the signal precisely because it
  is the part that survives — and because only a repeater ever binds a whole data item, while fields
  and labels bind leaf columns. Every corpus layout is asserted not to trip it.

### Changed

- **`refresh_xml_part` now re-points EVERY BC-namespaced binding at the layout's namespace** (GitHub
  issue #2), instead of swapping one known old namespace for the new one only when the report's own
  identity changed. A binding naming a THIRD namespace — a superseded namespace of the same report, or
  another report's after a control was copied between layouts — survived every refresh untouched, and
  20 of `PaymentPracticeByPeriod.docx`'s 25 bindings are exactly that. The 2026-08-02 sandbox rounds
  (issue #1) settled why it matters: BC validates every binding's `prefixMappings` against the target
  report's CURRENT namespace and refuses the upload with one `InvalidPrefixMapping` per offender,
  accepting the same layout once they are all re-pointed. So re-pointing is not a guess about intent —
  it is the only state BC accepts, and the previous behaviour left the layout provably un-uploadable
  with hand-rebuilding each control as the only repair. Repaired bindings come back as the new
  `repointedForeignBindings` (with the namespace each one previously named); a binding whose
  `prefixMappings` names no BC namespace at all is still left untouched, and the XPath is never
  rewritten, so a foreign binding that ALSO names a missing column is re-pointed *and* reported as an
  orphan. `validate_layout`'s `binding-namespace` message now points at this repair instead of stating
  that refresh will not perform it.
- `validate_layout` now reports `w:rPr/w:w = 0` as a **warning** instead of an error (GitHub issue
  #54). That is the character scaling Word itself writes into a content control's placeholder run
  (alongside `w:sz=0`), so it appears in any layout that has been opened and saved in Word;
  `ST_TextScale` requires at least 1, so `OpenXmlValidator` flags it. It renders correctly in both
  Word and BC and no tool here can edit run properties, so it was an unactionable failure on a file
  Word considers valid. Matched on the offending node, not on the validator's message text, and every
  other structural error keeps Error severity.
- Tool descriptions and error hints no longer claim a missing `partName` targets "the FIRST
  header/footer part": resolution has always preferred the first section's DEFAULT part, and the
  wording now says so and points at `partDetails`.
- **`ModelContextProtocol` SDK 1.4.1 → 2.2.0** (#48, #67; added after release, the note was
  missing). No observable change: tool lookup, argument binding and the call-tool filter behave the
  same on both versions, measured on the 1.0.0 and 1.1.0 packages. A host-level test now pins that
  boundary (GitHub issue #97), so a future SDK upgrade that changes it fails the suite.

### Fixed

- **A blank `create_layout` build now declares Word `compatibilityMode` 15** (GitHub issue #51). It
  shipped no `word/settings.xml` at all, and a document that declares no mode is mode 12 — Word 2007
  — to Word (measured: `Document.CompatibilityMode` reported `12`). Repeating-section content
  controls do not exist in that mode, so the first time anyone opened a tool-authored layout in Word
  and saved it, the Compatibility Checker converted every repeater to a plain rich-text control and
  dropped its `w15:dataBinding`: the lines table silently stopped repeating and binding while still
  looking like a layout. Blank builds now scaffold a minimal settings part declaring the mode every
  stock corpus layout declares (`DocumentSettingsScaffold`), so the Word round trip the docs
  recommend for restyling is safe. A `templatePath` build keeps its shell's own settings untouched —
  compatibility mode also selects Word's layout metrics, so retrofitting one could move a template's
  pagination; that case is reported instead, by the new check below.
- **`insert_table` and `insert_repeater_table` now pin the column widths they are given** (GitHub
  issue #52). Both emitted a `w:tblGrid` and per-cell `w:tcW` but no `w:tblW` and no
  `w:tblLayout` — and the OOXML default is AUTOFIT, so Word recomputed every column from cell
  content and a caller's `columnWidths` were advisory at best: a 5-column line table authored at
  `1500,4700,1200,1300,1300` read as near-uniform in Word, and one plain Word save rewrote an
  untouched 2-column grid from `2800,3000` to `3177,3066`. Tables the tools create now declare
  `w:tblLayout="fixed"` plus a `w:tblW` total equal to the grid's sum — both corpus-observed (all
  four tables in `StandardPurchaseOrder.docx` carry them), and the total is kept in step with the
  grid by the existing `set_column_widths`/`insert_column`/`remove_column` sync. The mock preview hid
  this: Word-COM's docx→PDF conversion happened to render close to the declared proportions, so
  `preview_layout` looked right while Word's own UI did not. Existing tables are untouched — a
  captured layout keeps its own layout algorithm.
- The NuGet MCP registry manifest (`.mcp/server.json`) shipped three `description` values over the
  registry schema's 100-character cap (the 2025-10-17 schema enforces it); all are shortened and a
  packaging test now walks the whole manifest so any future description stays within the limit.

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
