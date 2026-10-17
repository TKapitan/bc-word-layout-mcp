---
name: al-word-layout-from-design
description: Use when someone already has a finished Word DESIGN of a Business Central document (a styled .docx with sample values typed in — logo box, colours, an invoice/quote/order mock-up) plus the report's empty BC layout, and wants a working, data-bound BC Word layout that keeps that design. Written for a functional consultant who does not code - the agent does all file work. Covers what to ask for, mapping sample text to dataset fields (and confirming it), building the bound layout, checking it, and the handover. Complements al-word-layout (tool mechanics) and al-word-layout-design (from-scratch conventions).
---

# al-word-layout-from-design

**Scenario.** A consultant (or a designer working for one) has made the document in Word the way the
customer wants it — colours, fonts, a logo box, a line table with sample rows, totals — and the BC
report already exists (a standard report such as 1306 *Standard Sales - Invoice*, or a custom one).
They should not need to know OOXML, XPath or AL. The agent turns the design into a layout BC can
print; the consultant answers questions about *meaning* ("is PO-5521 the customer's order number?")
and uploads the result.

`al-word-layout-design` answers "what should a new layout look like". This skill answers "the look is
decided — make it print real data". Read `al-word-layout` for the tool mechanics; everything below
uses those tools by name.

Evidence tags as in the sibling skills:
- **[verified]** — rendered correctly by a real Business Central. Two cases, both report 1306
  *Standard Sales - Invoice*:
  - the single-file route (designed sales invoice + the report's empty layout → bound layout) was shown
    live at **BC North America Summit 2026** (session demo by @miljance): the generated layout printed
    correctly on its first import into BC;
  - a body built with this recipe printed correctly as part of a composable layout on 2026-10-17 (see
    `al-word-layout-composite`).
- **[mock]** — checked only with `validate_layout level=full` + `preview_layout`.

## 1. What to ask the consultant for

Ask only for what is missing, in plain words:

| Input | Why | How the consultant gets it |
|---|---|---|
| **The report's empty layout** (`.docx`) | It carries the BC dataset part — the list of every field the report offers. Never synthesized (al-word-layout §1). | In Business Central: *Report Layouts* → pick the report → create a new **Word** layout → export it. An exported standard layout of the same report works as well. |
| **The design** (`.docx`) | The look to keep. | Their Word file. It must not already be a BC layout (no bound fields). |
| **Which report** | Picks the dataset. | Usually obvious from the empty layout (`get_layout_info` → `report`). |

Tips worth passing on to whoever makes designs, so the next one converts cleanly:
- Type one realistic sample value per field ("Oct 7, 2026", "PO-5521"), not "XXXX".
- Lay data out in tables and paragraphs. Keep it out of floating text boxes and shapes: every tool here
  addresses body paragraphs and table cells only.
- Draw the line-items table with a header row and two or three sample rows.
- Mark a logo spot with a placeholder (a box or the word LOGO).

## 2. Read both files before deciding anything

1. `list_dataset_fields` on the empty layout: the data items (`Header`, `Header/Line`,
   `Header/Totals`, …), their columns, and which columns are labels (`isLabel`).
2. Read the design's text in document order, grouped by block (masthead, info strip, addresses, lines,
   totals, notes, signatures, footer). A plain dump of `word/document.xml` text with table/row/cell
   breaks is enough. Also note what is design, not data: colours, shading, borders, fonts, sizes.

## 3. Map every piece of sample text — and confirm the map

This is the step no tool can do, and the one that decides whether the layout is right. Build a table
with one row per text item in the design:

| Design text | Kind | Binds to | Note |
|---|---|---|---|
| `INVOICE` | label | `/Header/DocumentTitle_Lbl` | caps come from formatting |
| `No.` / `2026-0142` | label + field | `/Header/DocumentNo_Lbl`, `/Header/DocumentNo` | |
| `PO-5521` | field | `/Header/ExternalDocumentNo` | customer's own reference — confirm |
| `Payment information` | static | — | no label column exists: not translated |
| `Payment is due within 15 days…` | drop | — | states terms as fixed text; would print wrong |

Kinds: **field** (data), **label** (a caption the dataset provides — always bind it rather than typing
it, so it translates), **static** (no dataset column; stays literal and is NOT translated — say so),
**drop** (sample text that would print false information), **picture** (logo → a picture control).

Show the table to the consultant in plain language and get a yes, or corrections, **before building**
— one round. Ask only about real ambiguities; decide the rest with the rules below and list them.

### Mapping rules that held up [verified: report 1306]

- **Captions:** bind the `…_Lbl` / `Labels` column whenever one exists; static text only when none
  does. Expect a few static captions (section headings such as "Details", "Notes").
- **Company block:** `/Header/CompanyAddress1…8` are BC's formatted address lines and line 1 is the
  **name**; VAT and registration numbers each have a value and a `_Lbl` column. Lines a company does not
  use print empty — there is no hide-if-empty (al-word-layout §3).
- **Counterparty addresses:** `CustomerAddress1…8` (bill-to), `ShipToAddress1…8`; line 1 is the name.
  Reserve as many lines as the design shows; say that a contact or second name line needs a 5th.
- **References:** a customer's purchase-order number is `ExternalDocumentNo`; BC's own order number is
  `OrderNo`. "Delivery date" on an invoice is `ShipmentDate`. When in doubt, ask.
- **Line items:** item/description/unit/quantity/price/amount → `ItemNo_Line`, `Description_Line`,
  `UnitOfMeasure`, `Quantity_Line`, `UnitPrice`, `LineAmount_Line`, captions from the line's `_Lbl`
  columns. A running line counter ("1, 2, 3") does not exist in the dataset (`LineNo_Line` is 10000,
  20000…) — offer the item number instead, or drop the column.
- **Totals:** take them from the `Totals` data item, not from line sums: subtotal, invoice discount, the
  VAT *text* field (it carries the rate, so "VAT 20%" is never hard-coded), grand total text + amount,
  `CurrencyCode`.
- **Notes:** `WorkDescriptionLines/WorkDescriptionLine` repeated per line. Boilerplate promises (due
  dates, "valid without stamp") are dropped or replaced by a neutral sentence — confirm with the consultant.
- **Signatures, stamp boxes:** static captions.
- **Logo:** a picture control on the report's company picture column (`/Header/CompanyPicture` in 1306).
- **Footer contact line:** the company phone/e-mail/home page columns of the report's top-level item.

## 4. Build the bound layout

The design's look (shading, colours, borders, fonts) has no tool (al-word-layout-design §5), and no
tool turns sample text into a bound field *in place*: `insert_field` with `locationType="atText"`
inserts beside the match and nothing deletes plain text outside a table cell. So the binding is a
deliberate, validated package edit, exactly the route al-word-layout §4 reserves for "what genuinely
has no tool". Keep it mechanical:

### 4.1 Start from the design's package

Copy the design `.docx`; keep its styles, fonts, footer and section setup; remove review leftovers
(comments, tracked changes). Give it the dataset part one of two ways:

- `create_layout` with `schemaSource` = the empty layout and `templatePath` = the design: it ships one
  dataset part with a fresh `storeItemID`, the glossary part and compatibility mode 15, and keeps the
  design body as it is. Read the `storeItemId` back with `get_layout_info`. [mock]
- Or copy the empty layout's `customXml/item*.xml` (the `urn:microsoft-dynamics-nav/reports/…` part)
  into the package with an `itemProps` part declaring a new `storeItemID`, and set
  `compatibilityMode` 15 in `settings.xml`. [verified]

### 4.2 Replace each sample run with a bound control

Per the confirmed map, replace the run holding the sample text with a run-level plain-text content
control, copying the run's `w:rPr` into both the `w:sdtPr` and the inner run so the look is unchanged:

```xml
<w:sdt><w:sdtPr><w:rPr>…same as the sample run…</w:rPr>
  <w:alias w:val="#Nav: /Header/DocumentNo"/><w:tag w:val="#Nav: Standard_Sales_Invoice/1306"/><w:id w:val="…unique…"/>
  <w:dataBinding w:prefixMappings="xmlns:ns0='urn:microsoft-dynamics-nav/reports/Standard_Sales_Invoice/1306/' "
                 w:xpath="/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:DocumentNo[1]" w:storeItemID="{…}"/>
  <w:text/></w:sdtPr>
  <w:sdtContent><w:r><w:rPr>…same…</w:rPr><w:t>DocumentNo</w:t></w:r></w:sdtContent></w:sdt>
```

- Fields inside a line keep their static neighbours as plain runs (`"No.  "`, `":  "`, `" · "`).
- A **label** is the same shape bound to the `_Lbl` column. Static captions stay plain runs.
- The newer Word add-in writes `w:alias` = the column name and `w:tag` =
  `#BC:InsertDataItem,DataItems/Header/DocumentNo`. Microsoft's own sample bodies mix both; the `#BC`
  form is the one printed here [verified]. The binding is what BC reads — pick one convention per file.
- The **logo** becomes a picture control (`w:picture` in `w:sdtPr`, a `w:drawing` with a placeholder image
  inside, bound to the picture column) — the same shape `insert_picture` writes.

### 4.3 The line-items table

Keep the header row (bind its captions) and the **first** sample row; delete the other sample rows;
bind the kept row's cells; wrap that row in a row-level repeater over the line item — the shape
`insert_repeater_table` produces:

```xml
<w:sdt><w:sdtPr><w:alias w:val="#Nav: /Header/Line"/><w:tag w:val="…"/><w:id w:val="…"/>
  <w15:dataBinding w:prefixMappings="…" w:xpath="/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line" w:storeItemID="{…}"/>
  <w15:repeatingSection/></w:sdtPr>
  <w:sdtContent><w:sdt><w:sdtPr><w:id w:val="…"/><w15:repeatingSectionItem/></w:sdtPr>
    <w:sdtContent><w:tr>…the bound row…</w:tr></w:sdtContent></w:sdt></w:sdtContent></w:sdt>
```

Mark the header row `w:tblHeader` so it repeats on every page. Alternating row shading in the design
cannot survive (one template row repeats): drop it, or move it into a table style's `band2Horz` and
set the table's `w:tblLook` (§5).

### 4.4 Never wrap the whole body in a repeater over the top-level item

Bind straight to `/Header[1]/…`: BC already prints the body once per document. A block repeater over
`/Header` without a `w:group` prints an EMPTY page (validate_layout's `top-level-repeater-group`);
when one is really wanted, `wrap_in_repeater` writes the correct shape. Block repeaters over nested
items (work-description lines) need no group.

## 5. Check it

1. `validate_layout level=full` — 0 errors; every binding resolves.
2. `preview_layout`, then `render_preview_pages` with `outputDir` so the consultant can open the pages.
   Use `rows` ≥ 20 once, to see the line table break across pages with its header repeated.
3. Compare page 1 side by side with the design. Expected differences only: values, dropped zebra rows,
   anything the map changed.

Mock-render traps met on this route (not layout bugs):
- A LibreOffice conversion fills bound controls from the dataset part, so a stripped-down mock may show
  element names where you expected sample text.
- LibreOffice reads only the legacy hex `w:tblLook w:val`; keep it consistent with the
  `firstRow`/`lastRow`/`noHBand` attributes (`0x20` first row, `0x40` last row, `0x200` no banding,
  `0x400` no vertical banding) or header/total-row formatting silently disappears in the preview.
- A paragraph style's colour beats a table style's conditional colour: a header-row paragraph styled
  grey stays grey on a navy band. Leave such paragraphs unstyled (`Normal`) and keep the base text
  colour in `docDefaults`.

## 6. Hand over

Send the consultant the layout plus the preview pages, and say in plain words:
- what is bound (in their terms, not XPaths), what stayed **static and will not translate**, what was
  dropped and why;
- what to watch in the first real print (empty address lines, long descriptions wrapping);
- next step: upload it on the *Report Layouts* page as a new Word layout for the report, set it as the
  default if wanted, print one real document. That print is the sign-off (al-word-layout §1, outer loop).

When the customer wants the same look across many reports — a shared header/footer and theme — split
the result into a composable layout: `al-word-layout-composite`.

## 7. Missing tool

No such tool exists yet — do not try to call one. Everything in §4.2 would be a single typed call with
a tool that replaces matched sample text with a bound field or label control, inheriting the run's
formatting (proposed working name: *replace text with field*: layout path, search text, occurrence,
dataset path), plus a row variant that turns a sample table row into a repeater row and deletes its
sibling sample rows. Until it exists, the package edit above is the route — always followed by
`validate_layout level=full`.
