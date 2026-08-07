---
name: al-word-layout-design
description: Use when designing a NEW Business Central Word report layout (or re-shaping an existing document) — deciding what the document should look like before any edit tool is called. Covers the two archetype skeletons (trading document, list/analysis report), the observed BC design conventions (address blocks, caption placement, line-table alignment, borders, totals, header/footer chrome, typography), the BC-verified tool recipe for each block, and the stock idioms the tools cannot build. Complements al-word-layout, which covers the tool mechanics, workflow, and error envelope — that skill is HOW to call the tools; this one is WHAT to build.
---

# al-word-layout-design

A prompt like "create a new layout for Purchase Order" leaves every design decision open, and
`validate_layout` cannot catch a bad answer: a structurally valid, fully bound, *ugly* document
passes `level=full` with zero findings. This skill closes that gap with the conventions a BC Word
layout is expected to follow. Every claim is tagged with its evidence:

- **[stock]** — observed in Microsoft's own layouts in `tests/corpus/` (document-shaped: Standard
  Sales Quote 1304, Standard Purchase Order 1322, Standard Sales Invoice w. VAT Spec 1306, Standard
  Statement 1316; list-shaped: Salesperson Commission 115, Job Quote 1016).
- **[verified]** — built from scratch with these tools and rendered correctly by a real Business
  Central sandbox (from-scratch pack items P01 quote / P06 purchase order / P07 commission report,
  round 2, BC-verified 2026-08-02 — the same verification round `docs/FIDELITY-CHECKLIST.md` cites).

Everything here is a **default, not a law**: explicit user requirements always win, and where the
stock layouts disagree with each other this skill says so rather than inventing a rule. Maintenance
rule (same discipline as `docs/FIDELITY-CHECKLIST.md`): a sandbox round or corpus addition that
refutes a convention updates this file.

## 1. Decide the shape first

Pick the archetype before calling any tool:

| Archetype | Fits | Skeleton |
|---|---|---|
| **Trading document** | Quote, order, invoice, credit memo, shipment — one document per record, addressed to a counterparty | Title/logo chrome → address block → document-info grid → line-items table → totals (§2) |
| **List / analysis report** | Commission, register, aging, activity lists — rows grouped under entities, running chrome | Header/footer parts with report chrome → grouped repeater table with nested detail rows → totals (§3) |
| **Per-entity statement** | One sub-document per customer/vendor with a page break between them | **Not buildable from scratch** — the whole body sits inside a per-entity repeating section [stock: 1316], and repeaters here are tables only. Export the stock layout and edit it instead (§5). |

Settle before building (ask only what the request leaves open): which parties' addresses (two is
the default; add ship-to when it differs [stock: 1322 buy-from + ship-to + company]), which
document-info fields, which line columns and their order, and whether a branded `templatePath`
shell exists or the build is blank.

Then obtain the schema artifact (al-word-layout §1 — transplanted, never synthesized), run
`create_layout`, and call `list_dataset_fields` **before designing anything**: design against what
the dataset actually offers. Bind label columns wherever they exist; when a report has none (old
`Labels`-data-item style), `headerFromLabels=true` falls back to humanized static headers —
BC-verified acceptable [verified: P07]. A blank build's body already carries a bold heading
paragraph with the report name; pass `create_layout`'s `headingText` to override it (the verified
builds pass the all-caps document name, e.g. `"STANDARD PURCHASE ORDER"`).

## 2. The trading-document skeleton [verified: P01, P06]

Build top-to-bottom; every block is appended with `locationType=documentEnd`, then filled through
the table index the insert response returns. The whole recipe below rendered correctly in BC.

| # | Block | Calls | Verified shape |
|---|---|---|---|
| 1 | Title | (from `create_layout headingText`) | All-caps report/document name, bold heading paragraph |
| 2 | Logo | `insert_picture` | `/…/CompanyPicture`, 35×18 mm |
| 3 | Address block | `insert_table` 6×2, `columnWidths "4600,4600"`, then 12× `insert_field` (`tableCell`) | Borderless; company addresses down col 0, counterparty (customer/vendor) down col 1, one bound field per cell, filled row-major |
| 4 | Info grid | `insert_table` 3×2, `columnWidths "2800,3000"`, then per row `insert_label` col 0 + `insert_field` col 1 | Caption **beside** value — see §4 for the caption-**above** stock alternative |
| 5 | Section heading *(optional)* | `insert_text` | `bold=true, fontSizePoints=16` heading (e.g. `"ORDER LINES"`), plus a default-size note line [verified: P06] |
| 6 | Line items | `insert_repeater_table`, `headerFromLabels=true`, `columnWidths`, `columnAlignments` | Numeric columns right, text left, description widest; `tableStyle="TableGrid"` for a full grid **or** no style + a header rule for the open look (§4) |
| 7 | Totals | `insert_table` 1×2, `columnWidths "6200,2400"`, `columnAlignments "right,right"`, 2 fills, `set_cell_borders` | Totals text/label left, **bold** amount right, ½ pt rule on top (`edges="top", size=4`) |

Checkpoint rhythm: the verified builds ran one `validate_layout level=full` + one `preview_layout`
at the very end and passed 0/0 — sufficient at this scale, though per-block `preview_layout` +
`render_preview_pages` is cheap whenever unsure.

## 3. The list-report skeleton [verified: P07]

1. **Header part**: `insert_label` (report-name label) with `layoutPart="header"`, then any static
   text via `insert_text locationType="afterControl"` anchored to that label's `controlId` —
   control first; plain runs return no id and can never be the anchor.
2. **Footer part**: same pattern (e.g. an "amounts are in LCY" label). The scaffolded blank-build
   header/footer parts render with data on every page in BC [verified: P07].
3. **Main table**: `insert_repeater_table` over the top-level data item; humanized static headers
   are fine when no label columns exist.
4. **Nested detail rows**: `insert_repeater_row` on the parent repeater's `controlId`; start the
   `cells` spec with `"-"` so the first column stays an empty indent spacer and the detail data
   sits under parent columns 2+; right-align amount cells via `alignments`. Holds the parent grid
   without drift across pages in BC [verified: P07].
5. **Totals**: the same 1×2 idiom as §2 — the caption may be an `insert_label` when it lives in a
   `Labels` data item [verified: P07].

Stock contrast [stock: 115, 1016]: Microsoft's list reports are ONE table with group-header rows
and per-group subtotal rows inside the repeat, and put the chrome (title, date, page, user,
printed filter expressions) in the running header part. Group-header/subtotal *rows* inside a
repeater are not tool-reachable (§5) — nested detail rows plus a separate totals table is the
buildable equivalent, and it is the shape BC verified.

## 4. Per-element conventions

### Address block
- Borderless plain table, one address line per row, one bound FIELD per cell [stock: all four
  document layouts; verified: P01, P06]. Blank address lines are simply empty controls — the stock
  layouts reserve the rows (no hide-if-empty markers anywhere in the corpus).
- Two columns at `4600,4600` is the verified default. Stock varies by requirement: 2-col
  customer|company [1304], 4-col buy-from|spacer|ship-to|company [1322], stacked single-column
  blocks [1316] — party count follows the document, not a rule.
- Captions: mostly none. The only stock caption is a label **above the ship-to column** when
  ship-to differs from buy-from [1322]. Don't caption sell-to/company blocks by default.
- Side-by-side company blocks are right-aligned in two stock files [1304, 1306], left in one
  [1322]; the verified builds left everything left-aligned — either renders fine.

### Info grid (document no, dates, terms, references)
- Captions are **always bound label controls**, never static text [stock: all four — zero static
  captions exist in the entire document corpus]. `insert_label`, not `insert_text`.
- Two observed caption placements — both fine, say which you chose:
  - **Beside** — label col 0, value col 1, one datum per row (3×2 at `2800,3000`): the verified
    recipe [P01, P06] and the Statement's stock shape [1316].
  - **Above** — caption row over value row (e.g. 2 rows × 5 equal columns [1304], stacked
    caption/value row-pairs [1322]): the *dominant stock shape* [1304, 1322, 1306]. Build as
    `insert_table` rows×cols, `insert_label` into the caption row's cells, `insert_field` below.
    Stock captions are smaller + bold (9 pt bold over 11 pt values).
- Info-grid values are left-aligned in every stock file — don't right-align them.
- Keep the grid narrower than the page (verified: 5800 twips, left-anchored) — it is not a
  full-width band.

### Line-items table
- Numeric columns AND their header captions right-aligned; text columns left; unit-of-measure
  stays left [verified: P06]. Description is always the widest column, ~30–40% of the grid
  [stock: all; verified: all].
- Header row: bind label columns when the dataset has them (`headerFromLabels=true`); humanized
  static text otherwise [verified: P07].
- The header row repeats on every page of a multi-page table without you doing anything —
  `insert_repeater_table` always marks it `w:tblHeader`, a corpus-observed shape [stock: 1304,
  115]. BC honours it [verified: P07], and since the #19 fix the mock preview repeats it too.
- Two buildable looks, both BC-proven:
  - **Open** [stock: 1304, 1322]: no table style, single ½ pt rule under the header row —
    `set_cell_borders` `row=0, edges="bottom", size=4`.
  - **Full grid** [verified: P06]: `tableStyle="TableGrid"` — the style reference alone draws the
    ½ pt grid (BC honours `w:tblStyle` against the scaffolded styles part); don't also paint
    per-cell borders.
  - Stock 1306 additionally rules under *every* data row — a bottom border on the repeater's data
    (template) row should repeat per line, but that exact shape is **not BC-verified**; its solid
    colour header band is template/hand-edit territory (no tool paints shading).
- ½ pt (`size=4`) is the only rule weight observed anywhere in the corpus — use it unless asked.

### Totals block
- Tool recipe [verified: P01, P06, P07]: a **separate** borderless 1×2 table directly after the
  line table — `columnWidths "6200,2400"`, `columnAlignments "right,right"`, totals TEXT field (or
  label) in col 0, amount field with `bold=true` in col 1, then `set_cell_borders`
  `edges="top", size=4` on row 0.
- A totals *ladder* (excl. VAT / VAT / incl. VAT [stock: 1322]) is the same table with more rows:
  amounts share the right column, the grand-total row gets the bold + rule treatment.
- Stock layouts put totals **inside** the line table as trailing right-anchored rows — not
  reachable for a tool-built repeater (§5). Preserve that shape when *editing* a stock layout;
  build the separate table when *creating*.

### Chrome: title, logo, date, page numbers
- Two BC-accepted placements:
  - **Body masthead** [verified: P01, P06; stock: 1306, 1316]: heading + logo + addresses at the
    top of the body. The simple default for from-scratch builds.
  - **Header-part chrome** [stock: 1304, 1322]: title/doc-no/date/page in the header parts,
    contact grid in the first-page footer, `w:titlePg` set. Header/footer *bindings* are verified
    [P07, plus a stock-layout header/footer probe in the same sandbox rounds] — but remember the
    first-vs-default part trap (issue #5): with `titlePg`, page 1 renders the FIRST-page part, so
    an edit landing in the default part is invisible on page 1.
- Stock page numbers are `Page_Lbl` + two literal spaces + Word `PAGE`/`NUMPAGES` field codes
  [stock: 1304, 1322, 1306, 115 — identical construct in all four]. Buildable: `insert_label` the
  `Page_Lbl`, `insert_text` `"  "` (`afterControl` on the label), then `insert_page_number`
  (`afterControl` again) — it emits the exact stock field construct (`X / Y` by default;
  `includeTotal=false` for the bare page number).
- Inline sequences anywhere: insert the control first, then `insert_text` with
  `locationType="afterControl"` anchored to it — `documentEnd` appends a NEW paragraph per call.
  Separators are literal runs (`" "`, `"  |  "`, `", "`) [stock: every multi-control line;
  verified: P07].

### Typography
- A blank build ships pinned Calibri 11 pt + `Normal`/`TableGrid` styles, BC-verified to render
  identically in Word and BC [verified: all three]. Don't fight it — layer `bold`/`fontSizePoints`
  where needed (a 16 pt bold section heading over the pinned base is verified [P06]).
- Stock caption text is smaller and bolder than its values (9 pt bold captions over 11 pt values
  [1304, 1322]); totals amounts are bold everywhere [stock + verified].
- Broad styling (colours, shading, other fonts) has no tool: use a branded `templatePath` shell or
  hand-edit the OOXML and re-validate (al-word-layout §3).

### Widths (twips)
- Full A4 content width is **10206 twips** — what `insert_table`'s default even split fills.
- Verified sums: address 9200, info grid 5800, line tables 9700–10206, totals 8600 — all
  BC-accepted, so table sums need not hit the margin exactly; pick round numbers that keep the
  proportions above.

## 5. Stock idioms the tools cannot build — route, don't improvise

| Stock idiom | Why no tool | Route |
|---|---|---|
| Totals as trailing rows inside the line table [stock: 1304, 1322, 1306] | No tool appends static rows inside a repeater table (#28) | Separate totals table (§4) when creating; preserve when editing stock |
| Page-position-CONDITIONAL content (`IF PAGE = NUMPAGES`) | Blocked on add-in-compatible OOXML evidence (#11); plain `PAGE`/`NUMPAGES` itself IS buildable — `insert_page_number` (§4 Chrome) | `templatePath` shell carrying the construct, or hand-edit + `validate_layout` |
| Group-header / subtotal rows inside a repeat [stock: 115] | Repeater internals aren't row-editable (#30) | Nested `insert_repeater_row` detail rows + separate totals [verified: P07] |
| Colour header bands / accent colours [stock: 1306] | No shading tool (#15) | Branded `templatePath` or hand-edit + validate |
| Hide-if-empty add-in controls | Deferred, OOXML never captured (#8) | Reserve the rows/cells; stock layouts render empty lines too |
| Whole-body per-entity repeat with page breaks [stock: 1316] | Repeaters are tables only (#31) | Export the stock layout and edit; don't build from scratch |

## 6. Design-time mock caveats

No known mock-preview gap currently looks like a design mistake. The one that did — a repeated
table header row never spanning pages in the mock even when BC repeats it — is fixed (#19): the
mock now repeats it exactly as BC does. Sandbox render stays the sign-off (al-word-layout §4);
a newly discovered gap gets listed here so nobody redesigns around a mock artefact.

## 7. Handover checklist

- [ ] Archetype, parties, and line columns match the *user's* requirements — the defaults here
      filled only what the request left open, and any deviation from a stock convention is deliberate.
- [ ] Every caption the dataset provides is a bound label (`list_dataset_fields` `isLabel`), not
      static text; static headers only where no label column exists.
- [ ] Numeric columns + their captions right-aligned; description widest; ½ pt rules only.
- [ ] Totals: bold amount, top rule, right-anchored.
- [ ] `validate_layout level=full` passes 0/0 and the preview pages were actually looked at
      (`render_preview_pages`), modulo any §6 caveats.
