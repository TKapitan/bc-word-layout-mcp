# Backlog — BC Word Layout MCP Server

What is deferred, why, and what each item is blocked on. Implemented work is not tracked here — the
release notes in [`CHANGELOG.md`](../CHANGELOG.md) and git history are the record.

**Ordering principle:** the tool is planned for **publication as an open-source community project**, and
every priority below is set against that goal. The runtime architecture stays as it is — single-user,
local, stdio; synchronous handlers, per-path locks, static design. What the publication goal changes is
the maintainer/environment model: unknown contributors, machines without Word, untrusted repos as inputs.

Priorities:

- **P0 — Public-launch blockers.** Publishing the repo without these either breaks the law/IP or ships a
  tool whose core promise is silently unverified.
- **P1 — Pre-publication hardening.** Do before accepting external PRs; each gets materially harder (or
  quietly erodes) once outside contributors are in the code.
- **P2 — Cleanup.** Real defects and hygiene, none launch-critical; batch opportunistically.
- **P3 — Feature roadmap.** Post-launch functionality, ordered by expected community value.
- **P4 — Standing verification debt.** Permanent manual checks; not closable once, tracked so no release
  silently skips them.

Two kinds of item live here. Some are **blocked on something this repo cannot provide**: a LibreOffice
install (B4), a BC sandbox (B42, B27), the real BC add-in (B24), a Word-authored reference table (B25), a
Word behaviour check (B37), or a maintainer decision (B41). The rest are **ordinary code work
waiting to be scheduled** (B43, B44, B45) — all three found by rendering this tool's output in a real BC
sandbox on 2026-08-01.

---

## P0 — Public-launch blockers

**None open.** The last one (B8, threat-model recalibration under the community model) closed
2026-08-02 — the resulting threat model and its accepted residual surfaces are documented in
[`SECURITY.md`](../SECURITY.md). The release process lives in [`RELEASING.md`](RELEASING.md).

## P1 — Pre-publication hardening

Wanted before external contributors arrive. The BC-sandbox fidelity validation that dominated this
section is **done** (2026-08-01 — all five pack items passed; the recorded result and what it measured
live in [`FIDELITY-CHECKLIST.md`](FIDELITY-CHECKLIST.md)), which leaves one open question it did not
settle.

## P2 — Cleanup

Real but non-blocking. Batch these as one or two cleanup passes.

| Id | Item | Why / what to do |
|---|---|---|

## P3 — Feature roadmap (post-launch)

Ordered by expected community value.

| Id | Item | Summary |
|---|---|---|
| B25 | Remaining table-structure ops — **`w:vMerge` and `w:gridBefore` only** | `w:gridAfter` is supported: it turned out to be a mainstream base-app shape, including on the line-items repeater table of `StandardSalesInvoiceVatSpec.docx`, and every column operation now carries the skipped-run offset. `w:gridBefore` falls out of the same arithmetic and is supported too, though still unseen in the wild — so it is covered only by synthetic tests. What remains is **`w:vMerge`** (vertical merges): zero occurrences across every layout reviewed, so there is no fixture and it is still rejected by every op. Author a reference vertical-merge table in Word first, then implement against it. |
| B27 | Repeater tables in headers/footers | Creation-side only (read/validate already handles them, with a warning). Mechanically small — `LocationResolver` already resolves header/footer roots and `LayoutEditor.InsertRepeaterTable`'s body-only check is a deliberate scope guard — but blocked on a DECISION, not code: `LayoutValidator` actively warns this shape is unsupported, and nothing confirms BC renders a repeating section in a header. The 2026-08-01 sandbox round did NOT cover this — it proved BC renders bound fields in header and footer parts, which is a different question from whether a repeating SECTION works there. Needs its own probe; pair it with B42's, since both want one sandbox session. |
| B29 | Page-position-conditional content | "Company name in the footer on the last page only" — a pure Word construct (`IF PAGE = NUMPAGES` field code) no tool emits. Confirm add-in-compatible OOXML first, then a dedicated tool — or document hand-edit + `validate_layout` as the path. |
| B31 | Parked ideas | Broader cosmetic formatting beyond the shipped knobs (bold/alignment/size on `set_cell_text`/`insert_field`/`insert_label`; per-column alignments) — e.g. colors, italics, fonts: hand-edit + validate remains the supported path. RDL→Word assisted conversion; Excel layouts; BC tenant upload (no public API exists — revisit only if Microsoft ships one). |

## P4 — Standing verification debt

Not closable once — tracked so no release silently skips them.

- **Per-release manual fidelity dimensions** (`FIDELITY-CHECKLIST.md`): fonts, pagination/page-break
  placement, number/date/locale formatting. Permanent manual-comparison line items — measured once
  (2026-08-01) is not measured forever, and every change to `MergeEngine`, `SampleDataGenerator` or a
  converter can move them.
- **Re-run the sandbox pack when the tool surface changes.** `tools/e2e/sandbox_pack.py` +
  `tools/e2e/bc_compare.py` make a full round mechanical; the standing cost is one sandbox session per
  release, not a project. Answering "is preview fidelity good enough?" once does not keep it answered.
