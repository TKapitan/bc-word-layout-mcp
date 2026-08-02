# ADR-0006: The dataset schema is transplanted, never synthesized

**Status:** Accepted (v1.0.0, recorded retroactively) · **Source:** SOLUTION-DESIGN §6.8

## Description

The BC dataset custom XML part that every binding in a layout resolves against always comes from a
BC-produced artifact. `create_layout`'s `schemaSource` and `refresh_xml_part`'s `newSchemaSource`
accept an existing `.docx` layout (its BC part is copied byte-for-byte) or a standalone exported
schema `.xml` (validated, then used as-is). Nothing in this codebase constructs that XML from an AL
report's source code, and no tool will be added that does.

The workflow consequence: for a brand-new layout, the schema artifact is obtained *outside* this
server first — one AL build (the compiler creates the `.docx` referenced by `WordLayout`/`rendering`
if it does not exist, dataset part included) or one export from BC (a stock built-in layout, or a
schema `.xml`). The skill (§1, "Where the schema comes from") documents this as the entry point of
the workflow.

## Why

The AL compiler owns the dataset-to-schema mapping: element-name derivation from column names,
`IncludeCaption` → `_Lbl` label columns, the `Labels` data item, namespace construction from report
id/name. Reimplementing that mapping here would create a second source of truth that must track the
compiler release-by-release; any drift produces layouts whose bindings resolve against *our* part
but not against the part BC regenerates on the next build — orphaned controls discovered in a
sandbox, the most expensive place to discover them.

Transplanting instead makes the equivalence hold **by construction**: whatever `alc` (or BC's own
layout update) would write is exactly what gets written, because the source of truth *is* a
compiler/BC output. There is nothing to regression-test against compiler behaviour, because there is
no independent implementation to drift. This is the schema-part sibling of
[ADR-0005](0005-emit-only-observed-ooxml.md): where ADR-0005 refuses to *emit unobserved OOXML
shapes*, this ADR refuses to *author the schema part's content* at all.

What the tools do author is deliberately confined to the wrapper around the transplant:
`refresh_xml_part` preserves the existing `ds:itemID` so bindings keep linking to the same part,
rewrites the schema references, and on a namespace change remaps `w:prefixMappings`/`w:tag` — see
`LayoutRefresher` and issues [#1](https://github.com/TKapitan/bc-word-layout-mcp/issues/1) /
[#2](https://github.com/TKapitan/bc-word-layout-mcp/issues/2) for the open questions on that wrapper
behaviour.

## Consequences

- "Create a new layout for report X" is never blocked by this rule, but it does start with one AL
  build or one BC export. A report that exists only as uncompiled AL source with no BC artifact
  anywhere has no schema to bind against — the missing step is a build, not a hand-authored part.
- A feature request for "generate the schema from `.al` source" (zero compiler/BC round trips) is a
  proposal to reverse this ADR: it must demonstrate how the generated part is kept provably identical
  to `alc` output across compiler versions — the fidelity-by-construction property it would forfeit.
- PRs that hand-construct `NavWordReportXmlPart` content (in tools, fixtures presented as real
  captures, or docs examples presented as canonical) will be declined; corpus-derived or
  compiler-generated artifacts are the evidence base, as with ADR-0005.
