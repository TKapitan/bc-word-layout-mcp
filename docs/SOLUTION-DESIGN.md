# Solution Design — BC Word Layout MCP Server

| | |
|---|---|
| **Name** | `bc-word-layout-mcp` |
| **Status** | **Living document.** Describes v1.0.0 as built. Update it in the same change that alters a decision here. |
| **Audience** | AL developers using Claude Code / MCP-capable agents, and anyone maintaining this server |
| **Owner** | Tom Kapitan |

This is the **source of truth for design decisions**: what the tool is, how it is built, and *why* it is
shaped this way. Two neighbours cover what this deliberately does not:
[`README.md`](../README.md) is the user-facing reference for each tool's parameters and behaviour, and
[`BACKLOG.md`](BACKLOG.md) is what is deferred and what each item is blocked on. Where this document
describes a rule ("refuse rather than guess"), the code is expected to match it; where they disagree, one
of the two is a bug.

---

## 1. Purpose and scope

Enable AI agents to **create, modify, validate and preview Microsoft Word report layouts** for Business
Central through deterministic, typed MCP tools — instead of free-hand OOXML editing by the model, which is
unreliable in a file format where a single misplaced element silently corrupts the document.

### In scope

- Layouts as **files in the AL workspace** (extension-owned layouts referenced from the `rendering`
  section) or exported from a BC client.
- Reading report dataset schemas from the layout's custom XML part or an exported report schema XML.
- Deterministic editing: fields, labels, static text, pictures, plain tables, repeater tables, nested
  detail-row repeaters, and table structure (columns, widths, merges, cell borders and text).
- Structural + binding **validation**.
- **Mock preview**: sample-data merge + PDF conversion, entirely offline, plus page rasterisation so the
  agent can visually inspect its own work.

### Out of scope

- **Tenant/user layout upload to BC** — no public API exists; it is UI-only. Revisit only if Microsoft
  ships one.
- **RDL→Word conversion, Excel layouts.**
- **BC add-in conditional controls** (*Hide Field if Zero*, *Hide Empty Table*, *Hide Empty Table Row*,
  layout comments). Their OOXML encoding is undocumented and has never been captured, and this project does
  not emit OOXML it has not seen (§6). Backlog **B24**.
- **Broad cosmetic formatting** (colours, fonts, margins). A deliberate exception exists: `bold`,
  `fontSizePoints` and cell/column `alignment` are supported, because a control inserted into a freshly
  authored table cell has no styling to inherit and would otherwise look obviously wrong. Anything beyond
  that stays a hand-edit, which `validate_layout` makes safe.

### Users and workflow context

AL developers. The outer verification loop is unchanged: build → publish to a dev sandbox → run the report.
This server provides the **fast inner loop** — edit → validate → preview in seconds, with no deploy.

## 2. Background: what a BC Word layout actually is

Verified against 61 real base-app layouts. Several claims here contradict what the format looks like from
a distance, and each one has cost a defect at least once, so they are stated explicitly.

**1. A custom XML part holding the dataset** — nested data items and columns.

- Namespace: `urn:microsoft-dynamics-nav/reports/<Report_Name>/<id>/` — **nav**, plural `reports`. (Not
  `dynamics-365/report/…`.) Root element `NavWordReportXmlPart`.
- The first child is a `BCReportInformation` metadata block, not a data item; it must be skipped when
  building the dataset model.
- **Labels are usually ordinary dataset columns** following a `*Lbl`/`*_Lbl` naming convention — there is
  normally no `<Labels>` element. But a dedicated `<Labels>` data item *does* occur, and its columns are
  then suffixed `Caption`/`Label` instead. Neither shape is rare. This is why label detection is a
  **configurable convention** (`LabelConvention`, overridable per host via `BCWL_LABEL_SUFFIXES` /
  `BCWL_LABELS_DATA_ITEM`) rather than a hardcoded suffix test.
- Encoding varies: usually UTF-16 LE with BOM, but UTF-8 also occurs. The part may carry **no
  `itemProps`/`DataStoreItem`** at all, and its directory casing varies (`customXml/` vs `customXML/`).
  Nothing may assume any of these.
- A layout may contain **several unrelated custom XML parts** (Office bibliography, MSIP/SharePoint
  metadata — routine in real exports; the published corpus copies are scrubbed of them, so synthetic
  fixtures cover the multi-part case). The BC part must therefore be found **by namespace, never by
  index or count**.

**2. Content controls (`w:sdt`) bound to that part.** Two binding shapes are both real and both common:

- `w:dataBinding` with a `<w:text/>` marker and `#Nav:`-prefixed `w:alias`/`w:tag`; and
- `w15:dataBinding` with **no** `<w:text/>` and frequently **no** alias or tag at all.

So the `#Nav:` alias convention cannot be relied on to identify a control — classification is driven by
`w:sdtPr` markers plus the binding, never by alias text. Control kinds: plain-text field/label,
`w15:repeatingSection`/`w15:repeatingSectionItem` (a row per record), and picture controls.

**3. The whole report body is an IMPLICIT top-level repeater.** BC iterates the top-level data item and
re-renders the body per instance, but this is *not* a `w15:repeatingSection` in the layout. Top-section
tables are ordinary `w:tbl`s with bound cells at document-root context. Consequence: adding a bound column
must not require an explicit repeater, and must not derive a data item from one.

**4. Real BC tables are `gridSpan`-pervasive and RAGGED.** Per-row physical cell counts differ across rows
of the same table, each reconciled to the grid by different span patterns; rows may also skip leading or
trailing grid columns (`w:gridBefore`/`w:gridAfter`). A "rectangular tables only" precondition would reject
essentially every real layout. Columns are therefore addressed by **grid column index**, never physical cell
index, and every row operation reasons through `gridBefore + Σ gridSpan + gridAfter = gridCount`.

**5. Real layouts ship with real defects.** Bindings pointing at a superseded or foreign report namespace,
`storeItemID`s naming a part absent from the package, `prefixMappings` written as a bare URI with no
`xmlns:` declaration, external `attachedTemplate` relationships pointing at a developer's desktop. Stock
Microsoft layouts do all of this and BC prints them anyway. Validation must therefore distinguish "this
layout is unusual" (warning) from "this edit broke something" (error) — see §6, principle 4.

At runtime **BC — not Word — merges** the dataset into the layout and converts to the output format. The
mock preview approximates that pipeline (§5), and can never be authoritative about it.

## 3. Architecture

```
┌──────────────────────────────────────────────────────────────┐
│  MCP Host (Claude Code / VS Code / other MCP client)         │
└───────────────┬──────────────────────────────────────────────┘
                │ stdio (JSON-RPC / MCP)
┌───────────────▼──────────────────────────────────────────────┐
│  bc-word-layout-mcp  (.NET 10 console app)                   │
│                                                              │
│  ┌ McpHost ─────────────────────────────────────────────┐    │
│  │ ReadTools / EditTools / TableTools / LifecycleTools   │    │
│  │ ToolGuards  — the one open→validate→save-or-reject    │    │
│  │               choreography every mutating tool shares │    │
│  │ CrossProcessLock, ToolContracts ({ok,data,error})     │    │
│  └──────────────┬────────────────────────────────────────┘    │
│  ┌ Domain ──────▼────────────────────────────────────────┐    │
│  │ SchemaProvider    — dataset model from the XML part    │    │
│  │ LayoutReader      — control inventory + table structure│    │
│  │ LayoutValidator   — structure, bindings, BC rules      │    │
│  │ LayoutEditor      — control-level edits                │    │
│  │ TableStructureEditor / TableGridNavigator — grid ops   │    │
│  │ LocationResolver  — where an edit lands                │    │
│  │ SdtFactory        — builds correct content controls    │    │
│  │ LayoutBuilder / LayoutRefresher — create / re-schema   │    │
│  │ guards: TableGridConsistencyGuard, PlainTextNesting-   │    │
│  │         Guard, ResourceLimits, LabelConvention         │    │
│  └──────────────┬────────────────────────────────────────┘    │
│  ┌ Merge ───────▼────────────────────────────────────────┐    │
│  │ SampleDataGenerator — seeded, type-aware fakes         │    │
│  │ MergeEngine    — binding fill, repeater expansion,     │    │
│  │                  XPathReanchor per nesting level       │    │
│  │ ExternalRelationshipStripper, FullValidator            │    │
│  └──────────────┬────────────────────────────────────────┘    │
│  ┌ Render ──────▼────────────────────────────────────────┐    │
│  │ IPdfConverter → WordComConverter (primary, Windows)    │    │
│  │                 LibreOfficeConverter (fallback)        │    │
│  │ PdfRasterizer — PDF pages → PNG for agent inspection   │    │
│  └───────────────────────────────────────────────────────┘    │
└──────────────────────────────────────────────────────────────┘
        reads/writes .docx + schema .xml in the AL workspace
```

Dependencies run one way: `McpHost → Merge → Domain`, `McpHost → Render`. Domain never references Merge —
which is why `validate_layout level=full` (a dry-run merge) lives in `Merge.FullValidator` rather than
alongside the quick checks in `Domain.LayoutValidator`.

### Technology choices

| Decision | Choice | Rationale |
|---|---|---|
| Language/runtime | C# / .NET 10 | The OpenXML SDK is the only first-class library for content controls and repeating sections; `python-docx` has no SDT support. .NET 10 for a support horizon past launch. |
| OOXML | `DocumentFormat.OpenXml` | Battle-tested, and ships `OpenXmlValidator` — the write gate in §6 depends on it. |
| MCP | Official `ModelContextProtocol` C# SDK, stdio | No network surface; same distribution model as sibling MCP servers. |
| PDF | Word COM (primary), LibreOffice headless (fallback) | Word gives the highest fidelity and every target developer machine has it; LibreOffice keeps non-Windows and CI viable. Neither is required — without a converter, preview still returns the merged `.docx`. |
| Concurrency | Synchronous handlers, per-path locks, no DI | Single-user, local, one process. Async and a container would add machinery with no user-visible benefit; the one seam that needed injecting (converter selection) is a parameter. |

## 4. Tool surface

**23 tools.** [`README.md`](../README.md) is the per-parameter reference; this is the shape and the grouping.

| Group | Tools |
|---|---|
| Inspect | `get_layout_info`, `list_dataset_fields` |
| Validate | `validate_layout` (`quick` \| `full`) |
| Preview | `preview_layout`, `render_preview_pages` |
| Author | `create_layout`, `insert_field`, `insert_label`, `insert_text`, `insert_picture`, `insert_table`, `insert_repeater_table`, `insert_repeater_row` |
| Table structure | `insert_column`, `remove_column`, `set_column_widths`, `merge_cells`, `split_cells`, `set_cell_text`, `clear_cell_text`, `set_cell_borders` |
| Lifecycle | `remove_control`, `refresh_xml_part` |

Conventions that hold across the surface:

- **Absolute paths only.** Every tool takes absolute file paths; mutating tools write in place.
- **One uniform envelope**: `{ok, data, error}`, where `error` carries `code`, `message` and a **`hint`
  that is never empty** — it names the argument to fix and its valid values, or the inspection tool to call
  first. Nothing throws across the MCP boundary. The one uncovered edge is the SDK's own
  argument-binding step, which runs before any tool body: a wrong-typed argument returns an SDK-shaped
  error rather than this envelope.
- **Shared location addressing**: `{type: documentEnd | afterControl | tableCell | atText, controlId?,
  tableIndex?, row?, col?, searchText?, layoutPart?, partName?}`. Edits stay deterministic without the
  agent knowing OOXML positions. `layoutPart: header|footer` with no `partName` resolves the **first
  section's default** header/footer — not the first part in the package, which is frequently the even-page
  or first-page one.
- **Coordinates round-trip.** The `tableIndex`/`row`/`col` an agent reads from `get_layout_info` address
  the same physical spot in every editing tool, by construction: reader and editor share
  `TableGridNavigator` rather than keeping parallel implementations.

## 5. The mock preview pipeline

```
schema ──► SampleDataGenerator ──► sample dataset XML
                                        │
layout.docx ──► MergeEngine ◄───────────┘
                    │  1. resolve each binding XPath → fill the control's text
                    │  2. clone repeatingSectionItem per data row, re-anchoring
                    │     inner bindings per nesting level (XPathReanchor)
                    │  3. regenerate cloned drawing/bookmark ids so clones stay unique
                    │  4. substitute placeholder images for picture controls
                    │  5. strip external relationships (see §6, principle 5)
                    │  6. unresolved binding → warning + visible placeholder
                    ▼
             merged.docx ──► IPdfConverter ──► preview.pdf ──► render_preview_pages ──► PNGs
```

- **`SampleDataGenerator`** produces seeded, type-aware fakes inferred from column names and types (amounts
  → decimals, `*Date*` → dates, codes → short uppercase strings). A fixed `seed` makes previews
  reproducible. It pre-scans the document's repeater bindings so it only multiplies rows for data items the
  *layout* actually repeats — otherwise a deep schema multiplies rows^depth and the preview explodes.
  `dataOverridesPath` accepts a real exported dataset for a realistic look.
- **The merge doubles as the deep validator.** `validate_layout level=full` runs it in dry-run mode:
  whatever the engine cannot resolve is what BC would choke on.
- **Preview is a MOCK, and this is the project's central bet.** It never goes through the BC report engine.
  It is good enough to catch structural and binding mistakes on every edit; it is not a substitute for a
  sandbox render, which remains the sign-off step. Every preview response repeats this. Known gaps —
  caption translation, BC SaaS font substitution, pagination, locale formatting — and the per-release manual
  procedure are in [`FIDELITY-CHECKLIST.md`](FIDELITY-CHECKLIST.md). The bet was validated against real BC
  sandboxes on 2026-08-01 (BC 28.0 and 28.3): a five-layout authoring pack passed end to end — BC accepts
  what the tools author and the mock's structure matched the real render on the dimensions checked.
  Per-release manual fidelity dimensions remain open in the checklist.

## 6. Design principles

These are the rules the code is expected to follow. They are what makes the tool safe to point at a real
layout, and each exists because the alternative failed or would fail silently.

**1. Never open the original writable.** Every mutating tool works on a staged copy in the target's own
directory and commits with an atomic same-volume rename. A crash mid-edit cannot leave a torn file; a
rejected edit leaves the original byte-identical.

**2. Validate before saving, not after.** Each edit runs `OpenXmlValidator` on the result and is **rejected
outright** (`edit_would_corrupt`, file untouched) if it would introduce a *new* structural error — diffed
against the pre-edit state, so a layout that was already imperfect can still be edited. Table edits pass an
additional `TableGridConsistencyGuard` that catches rows desyncing from `w:tblGrid`, a corruption class
`OpenXmlValidator` accepts silently. `PlainTextNestingGuard` catches a second such class: a control nested
inside a plain-text control passes every structural check and still makes Word declare the file corrupt.

**3. Refuse rather than guess.** When an edit has more than one defensible outcome, the tool refuses and
says why. `insert_column` at a position falling inside a spanned cell of a row that must carry the new
content is refused, because widening would silently drop the requested field and splitting would make a
layout decision nobody asked for. `create_layout` refuses a `templatePath` that already carries its own
bound controls. Nested repeaters in a header/footer are refused. The cost of a wrong guess is a corrupt
deliverable a developer discovers in a sandbox; the cost of a refusal is one clear message.

**4. Emit only OOXML that has been seen in a real layout.** Shapes absent from every reviewed layout
(`w:vMerge`) stay unsupported rather than being implemented from the spec, because "valid per the schema"
and "what the BC add-in produces" are not the same thing. The corollary matters as much: when a shape turns
out to be real, the rejection is a bug — `w:gridAfter` was refused for months on the belief it was exotic,
while sitting on the line-items table of the stock sales-invoice layout.

**5. Treat the layout as untrusted input.** It comes from a repo the agent did not write.
`ResourceLimits` caps part size, part count, and recursion depth so a zip-bomb or a deeply-nested file
fails one call instead of killing the server on an uncatchable `StackOverflowException`.
`ExternalRelationshipStripper` removes external relationships from the merged copy before any renderer
opens it — a poisoned `attachedTemplate` pointing at a UNC path otherwise makes Word reach out on open,
leaking an NTLM hash or acting as SSRF. (Word's `AutomationSecurity` blocks macros, not template loading.)

**6. Structure over discipline.** Where correctness rested on remembering a convention, it was made
mechanical: one `SdtInspector` classification ladder instead of four hand-synced copies; one
`TableGridNavigator`; a dedicated `NotFoundException` so `not_found` cannot be produced by an unrelated
`InvalidOperationException`; error hints keyed on exception *type* rather than message text. This matters
more, not less, with outside contributors.

**7. Serialise per path, in and across processes.** Concurrent tool calls on one layout take a per-path
in-process lock; a `CrossProcessLock` (named mutex, keyed by normalised path) is acquired *inside* it so two
IDE windows cannot interleave. Read tools take the same pair briefly, so a read cannot observe a
half-committed rename.

## 7. Testing strategy

- **Corpus**: 10 real BC layouts in `tests/corpus/` (provenance per file in
  [`PROVENANCE.md`](../tests/corpus/PROVENANCE.md)), chosen so each contributes a structural shape
  the others lack (`Corpus.cs` records what each is for). Coverage of a shape *nothing* in the corpus has
  is treated as absent, not as passing.
- **684 xUnit tests**: schema parsing, control classification, every editor operation asserted on the
  resulting OOXML, validator positive *and* negative paths (each defect paired with a valid sibling that
  must **not** trip the same check, so a test cannot pass by firing indiscriminately).
- **Snapshot tests**: merged main-part XML against approved snapshots per corpus layout, plus a
  byte-identical determinism check.
- **Tool-surface tests**: every tool driven through its real MCP entry point, including the guard/envelope
  behaviour and every `hint` branch.
- **e2e scenarios** (`tools/e2e/scenarios.py`): 16 multi-step edit journeys against corpus copies, each
  verifying that an edit introduces **no new validation finding** versus the pristine layout, then
  previewing and rendering it. Caches are keyed on the server build, so a code change that flips an outcome
  cannot be masked by a stale pass.
- **Fidelity harness**: regenerates preview artefacts for manual comparison — see
  [`FIDELITY-CHECKLIST.md`](FIDELITY-CHECKLIST.md).
- **Converter tests are dependency-agnostic**: they pass whether or not Word/LibreOffice is installed; an
  unavailable converter is itself a tested outcome, not a skip.

## 8. Distribution and operations

- Ships as a stdio MCP server (no secrets, no network): a NuGet package launched via `dnx` plus
  self-contained GitHub Release binaries, and a plugin for Claude Code and VS Code/Copilot — all
  documented in the README's Install section; the release process is [`RELEASING.md`](RELEASING.md).
- Semver, with release notes in [`CHANGELOG.md`](../CHANGELOG.md). Tool names, parameters, the
  `{ok,data,error}` envelope and the error codes are a **public API**; the stability policy, the
  BC-version commitment, and the MCP SDK upgrade cadence are stated in the README
  ("Public contract and versioning") and ADR-0004.
- A **companion skill** (`al-word-layout`) documents the intended agent workflow — inspect → edit → validate
  → preview → sandbox verify — and when *not* to hand-edit XML.
- Telemetry: none. Structured logs go to stderr; stdout is the MCP protocol stream and must never carry
  anything else.

## 9. Known-open design questions

Tracked in [`BACKLOG.md`](BACKLOG.md), which is the live list. The ones that would change this document:

| Question | Item |
|---|---|
| How does BC resolve a binding whose namespace does not match its part? Decides whether that finding is a warning or an error. | **B42** |
| What OOXML do the BC add-in's conditional controls produce? Blocks the largest missing capability. | **B24** |
| Is the LibreOffice path good enough to be the non-Windows experience? | **B4** |
| Should `refresh_xml_part` re-point foreign-namespace bindings, or only report them? | **B41** |
