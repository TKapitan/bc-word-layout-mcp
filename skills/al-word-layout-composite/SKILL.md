---
name: al-word-layout-composite
description: Use when building, splitting or debugging a Business Central composable (composite) Word layout - a report BODY, a reusable HEADER/FOOTER layout and a report THEME (.dotx) that BC combines at print time. Covers what each of the three files must and must not contain, the style contract between theme and body, splitting one designed document into the three parts, the checks to run, and why a theme can fail to restyle a body. Complements al-word-layout (tool mechanics) and al-word-layout-from-design (binding a finished design).
---

# al-word-layout-composite

Business Central (2026 release wave 2) can print a document from three separately managed Word files:
a **body** layout for the report, a **header/footer** layout reused across reports, and a **theme** that
gives every report the same colours, fonts and table looks. Defaults can be set globally, per company,
per report or per layout. Microsoft's announcements describe the feature but not the file rules; the
rules below come from Microsoft's own sample files and from one end-to-end case.

Evidence tags:
- **[MS sample]** — observed in Microsoft's composable samples: themes *Calm*, *Default*, *Playful*
  (`.dotx`), header/footer *ExternalDefaultDetailedHFDesign*, bodies *WordBody_1/2/3*. Not in
  `tests/corpus/`.
- **[verified]** — a body + header/footer + theme set built with these rules for report 1306
  (*Standard Sales - Invoice*, a teal/navy "Fresh" design) printed correctly in BC on 2026-10-17.
- **[observed failure]** — a variant that printed WITHOUT the theme's styling in BC.

## 1. The three parts at a glance

| | Body | Header/footer | Theme |
|---|---|---|---|
| File | `.docx` | `.docx` | **`.dotx`** (main part content type `…wordprocessingml.template.main+xml`) |
| Dataset part(s) | the report's own (`…/reports/<Report>/<id>/`) | **Tenant Report Defaults** (`…/reports/Tenant_Report_Defaults/2000000001/`) | Tenant Report Defaults (no bindings) |
| May bind | anything in the report's dataset | only `BCReportInformation/…` (`ReportMetadata`, `ReportRequest`, `CompanyMetadata`, `CompanyMetadata/CompanyAddressLines`) | nothing |
| Headers/footers | none | the content: first-page and default parts (`w:titlePg`) | empty parts only |
| Look comes from | **named styles only**, no direct formatting | its own direct formatting (fixed colours) on top of a few standard styles | the style definitions + `theme1.xml` colours/fonts |

Each `.docx`/`.dotx` carries both dataset parts BC exports — the classic
`urn:microsoft-dynamics-nav/reports/…` one and an `…/businesscentral/internal/extended/reports/…` twin
[MS sample]. Bind to the classic one; the tools here recognise only that namespace as the BC part, so
the twin does not trip the single-part check. Both parts are transplanted from a BC export, never
written by hand (al-word-layout §1): the report's empty layout for a body, an exported header/footer or
theme layout for Tenant Report Defaults.

## 2. Body

- **No headers or footers, no direct formatting** — no colours, sizes, fonts or shading on runs,
  paragraphs or cells [MS sample: WordBody_2/3 carry none]. Alignment (`w:jc`), widths, spacing and
  `w:tblLook` flags are layout, not look, and are fine.
- Paragraphs, runs and tables reference **style ids** only. Every id used must exist in the theme (§4).
  Microsoft's bodies use the BC style set (`BCLabel`, `BCNumber`, `BCLabelNumber`, `BCBlockTable`,
  `BCDataTable`, `BCTotalsTable`) plus `NoSpacing` [MS sample]; the verified Fresh body used built-in
  Word styles only (`Title`, `Subtitle`, `Heading2`, `Heading3`, `NoSpacing`, `Strong`, `Emphasis`,
  `SubtleEmphasis`, `SubtleReference`, `ListTable4-Accent2`, `ListTable1Light-Accent1`, `PlainTable2`,
  `ListTable1Light`) [verified].
- Ship **neutral** definitions of every style the body uses in its own `styles.xml`, theme-colour based
  (`w:themeColor`, `w:asciiTheme`), plus the stock Office `theme1.xml` — WordBody_2 has 31 such styles
  [MS sample]. The body then looks plain on its own and takes its look from the theme. Copy Word's
  default definitions for built-in table/character styles from a Microsoft theme's `styles.xml` rather
  than inventing them.
- Tables: the BC skeleton is three tables — a block table for layout (`firstRow`, no banding), a data
  table for lines (`firstRow`, banded), a totals table (`lastRow`) [MS sample]; layout tables carry
  `w:tblCaption` / `w:tblDescription` ("Layout table. Used for positioning, not for tabular data.").
- **Which tool calls keep a body style-only.**
  - Fine: `format` with only `characterStyle` (a style reference; it is refused until the body's own
    `styles.xml` defines the id, which the bullets above require anyway), `alignment` on
    `set_text_format` and `set_cell_text`, `tableStyle` on `insert_repeater_table` and `insert_table`
    (e.g. `BCDataTable`, `BCBlockTable` — refused until the body's `styles.xml` defines it), and
    `borders: "none"` on both (no border at all, `insert_table`'s default). Inserts take the formatting
    of the text before them, which adds nothing to a body that has none.
  - Not fine, because they write direct formatting: every other `format` member (bold, italic,
    underline, size, font, colour), on the insert tools and on `set_text_format`; `set_cell_format`
    (rules and fills — its `borderStyle: "none"` writes an explicit "no border", which also hides the
    theme table style's own rule); and the other two `borders` values on both table tools
    (`headerRule` draws a rule under each first-row cell, `grid` a table-level grid). So a body's line
    table is `insert_repeater_table` with `borders: "none"` and `tableStyle: "BCDataTable"`.
- Bindings and repeaters follow al-word-layout and al-word-layout-from-design; the body keeps the
  report title and document number, because the header/footer cannot see report data. Tags in the
  newest samples: field `w:alias` = column name, `w:tag` = `#BC:InsertDataItem,DataItems/<Item>/<Column>`
  (`Labels/<Column>` for a `Labels` item); repeater alias `Repeater: <Item>`, tag
  `#BC:InsertRepeater,DataItems/<Item>` [MS sample: WordBody_3; verified].
- Page size and margins sit in the body's `w:sectPr`; keep the header/footer layout's margins the same.

## 3. Header/footer

- Binds only Tenant Report Defaults: `BCReportInformation/CompanyMetadata/CompanyLogo`,
  `CompanyDisplayName`, `CompanyPhone`, `CompanyEmail`, `CompanyHomePage`, VAT/registration/bank fields
  and their `…Caption` columns, `CompanyAddressLines/CompanyAddressLine1…8` (line 1 is the name),
  `ReportMetadata/ReportCaption`, `ReportRequest/DateAndTime`, … Tags: `w:alias` = column name,
  `w:tag` = `#BC:InsertDataItem,BCReportInformation/CompanyMetadata/CompanyPhone` [MS sample].
- The look is carried by **direct formatting** in the header/footer parts — fixed colours, sizes,
  bold — on top of standard styles (`Header`, `Footer`, `NoSpacing`) [MS sample: ExternalDefaultDetailedHFDesign].
  So the header/footer keeps its look whichever theme is chosen; restyle it in its own file. The tools
  write exactly this: the insert tools' `format` (bold, italic, underline, size, font, colour) on
  fields, labels, static text and page numbers, and `set_text_format` to restyle what is already there
  (a control by id, or a `cell`/`row`/`text` target with `layoutPart` `header`/`footer`).
- First page vs following pages: set `w:titlePg`, put the full masthead in the `first` part and a slim
  running header in `default`; give both footers the page number (`PAGE` / `NUMPAGES` fields) [verified].
  Use `get_layout_info` `partDetails` to address the right part (al-word-layout §4).
- The document body of this file stays empty except its `w:sectPr`.
- `insert_table` / `insert_field` / `insert_label` / `insert_picture` / `insert_page_number` with
  `layoutPart` `header`/`footer` work on this file like on any layout — build the masthead and the
  company-details footer as plain tables; there is no repeater in a header/footer.

## 4. Theme

- **A Word template.** Name it `.dotx` and give the main part the *template* content type
  [MS sample: all three themes]. A theme exported from BC can arrive with a `.docx` extension; rename it
  to `.dotx` before opening or editing it.
- Contents, nothing more [MS sample]: `styles.xml`, `theme/theme1.xml` (colour scheme + major/minor
  fonts), `numbering.xml` (for the `BCNoteBullet` bullet), empty header/footer parts, the two Tenant
  Report Defaults dataset parts, and a body holding only a small preview table (rows *Header / Band /
  Band / Total*) that shows the main table style — it has no function.
- `settings.xml` declares `<w:linkStyles/>` (update the document's styles from the template) and a
  `w:clrSchemeMapping` (maps `text1`/`background1`/`accent1…` to the theme colours) [MS sample; verified].
- **Restyle, do not invent.** The theme redefines styles that bodies reference; a style id no body
  uses does nothing. Define BOTH:
  - the **BC style set** — `BCLabel`, `BCLabelNumber`, `BCNumber`, `BCNoteBullet`, `BCBlockTable`,
    `BCDataTable`, `BCTotalsTable`, `BCCalculationTable`, `BCSignatureTable`, `BCDataTableInternal` —
    because Microsoft's bodies use them [MS sample: present in all three themes];
  - the **built-in** paragraph, character and table styles bodies use (Microsoft's Default theme defines
    the full built-in set, 427 styles; Playful only 60 and lacks the List/Plain table styles).
  Start from a Microsoft theme's `styles.xml`, keep every id, and change definitions.
- Put the base text colour in `docDefaults`, not in `Normal`: a paragraph style's colour beats a table
  style's conditional colour, so a coloured `Normal` would block a white header row or total row
  [verified].
- Table looks live in table styles: header band via `firstRow` (`w:shd` + white `w:rPr`), zebra via
  `band1Horz`/`band2Horz`, grand-total band via `lastRow`, hairlines via `insideH`. Paragraphs in those
  rows of the body should be unstyled so the table style wins.

## 5. The style contract — check it every time

Before handing over, list every `w:pStyle` / `w:rStyle` / `w:tblStyle` the body and the header/footer
use and confirm each id exists in: the theme, the body's own `styles.xml`, and any Microsoft theme the
customer might pick instead (Calm, Default, Playful). In BC a missing id renders as plain text with no
error. Within one file, `validate_layout` catches two of the three kinds: `table-style-resolves` (a
`w:tblStyle`) and `character-style-resolves` (a `w:rStyle`, including one that names a paragraph style)
warn about ids that file's own `styles.xml` does not define. `w:pStyle` ids, and every check against the
theme file, stay manual. Watch for Word's duplicate ids (`BCLabel0`, `BCLabelNumber0`) that appear when a body was attached to
a template in Word — no theme defines them [MS sample: WordBody_3].

## 6. Splitting one designed document into the three parts

Start from a bound single-file layout (al-word-layout-from-design) or the design itself:

| Design element | Goes to |
|---|---|
| Logo, company name/address/VAT, contact line, page numbers | header/footer (bound to Tenant Report Defaults) |
| Document title, document number, dates, addresses, lines, totals, notes, signatures | body |
| Colours and fonts | theme `theme1.xml` (`accent1`, `text2`, `background2`, major/minor fonts) |
| Text looks (title, captions, section headings, emphasised values) | theme definitions of the styles the body references |
| Table looks (header band, zebra, total band, shaded panels) | theme table-style definitions; in the body only `w:tblStyle` + `w:tblLook` |

Choose the body's style ids first (§2), then define them in the theme (§4). Say which design details
moved or changed — e.g. the title moving below the header's masthead, or a panel's accent border
becoming a plain shaded panel.

## 7. Checks

1. Body and header/footer: `validate_layout level=full`, `preview_layout`, `render_preview_pages` —
   each part alone. The theme has no bindings; open it and look at the preview table.
2. The style contract (§5).
3. No tool composes the three parts; a faithful offline approximation is: body `document.xml` + the
   header/footer parts and `w:sectPr` references + the theme's `styles.xml` and `theme1.xml`, then
   convert. Treat it as a mock.
4. Sign-off is a BC print with the three layouts selected (outer loop, al-word-layout §1).

## 8. What failed, what fixed it

- [observed failure] A `.dotx` theme with BC + built-in styles but **no** `w:linkStyles` and **no**
  `w:clrSchemeMapping`, with a body referencing those styles, printed without the theme's styling.
- [verified] Adding `w:linkStyles` + `w:clrSchemeMapping`, restyling only existing ids (built-in + the
  BC set), and a body using built-in ids only, printed correctly.
- The single cause was not isolated (settings vs style-id choice). Follow all of §4 until a narrower
  BC test exists.
