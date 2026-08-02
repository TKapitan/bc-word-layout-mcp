# Architecture Decision Records

Short records of the deliberate decisions an outside contributor would otherwise reasonably "fix".
Each states its context, the decision, and what a proposal to reverse it must demonstrate. The
long-form rationale lives in [`SOLUTION-DESIGN.md`](../SOLUTION-DESIGN.md) §6 — where an ADR and the
code disagree, one of the two is a bug.

| ADR | Decision |
|---|---|
| [0001](0001-synchronous-handlers-no-di.md) | Synchronous tool handlers, no DI container |
| [0002](0002-never-open-the-original-writable.md) | Never open the original layout writable |
| [0003](0003-refuse-rather-than-guess.md) | Refuse rather than guess |
| [0004](0004-uniform-response-envelope.md) | One `{ok, data, error}` envelope; nothing throws across the MCP boundary |
| [0005](0005-emit-only-observed-ooxml.md) | Emit only OOXML observed in a real BC layout |
| [0006](0006-schema-transplanted-never-synthesized.md) | The dataset schema is transplanted from a BC-produced artifact, never synthesized |
