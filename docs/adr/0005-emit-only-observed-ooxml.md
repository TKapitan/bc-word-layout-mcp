# ADR-0005: Emit only OOXML that has been observed in a real BC layout

**Status:** Accepted (v1.0.0) · **Source:** SOLUTION-DESIGN §6.4, §7 (corpus)

## Description

The tools only *emit* OOXML shapes that have been seen in a real Business Central layout (the corpus
in `tests/corpus/`, plus captures reviewed during development), and only *accept for structural
edits* the shapes they can reason about. Shapes absent from every reviewed layout — the standing
example is vertical merges (`w:vMerge`) — are **refused**, not implemented from the OOXML
specification.

## Why

"Valid per the schema" and "what the BC Word add-in produces / what BC's renderer handles" are not
the same thing, and the gap is exactly where silent corruption lives (a control nested inside a
plain-text control validates cleanly and still makes Word declare the file corrupt). The corpus is
the evidence base; a shape with no witness in it cannot be tested against reality, only against the
spec — and the spec is not the implementation BC runs.

The corollary binds in both directions: **when a shape turns out to be real, the refusal is a bug.**
`w:gridAfter` was refused for months on the belief it was exotic, while sitting on the line-items
table of a stock Microsoft sales-invoice layout. Refusals are cheap to lift once evidence exists;
guesses are expensive to un-ship.

## Consequences

- A feature request for a new shape (e.g. hide-if-empty conditional controls, issue [#8](https://github.com/TKapitan/bc-word-layout-mcp/issues/8)) is
  research-gated: author it with the real BC add-in, capture the OOXML, add the fixture — then
  implement against evidence.
- PRs implementing a shape "per the ECMA-376 spec" without a captured BC-produced fixture will be
  declined, however correct the XML.
- Corpus additions follow the rules in `CONTRIBUTING.md` (scrub script + `PROVENANCE.md` row).
