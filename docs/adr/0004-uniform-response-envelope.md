# ADR-0004: One `{ok, data, error}` envelope; nothing throws across the MCP boundary

**Status:** Accepted (v1.0.0) · **Source:** SOLUTION-DESIGN §4, §6.6; README "Response envelope"

## Description

Every tool returns `{ok, data, error}`. On failure, `error` carries `code`, `message`, and a `hint`
that is **never empty** — it names the argument to fix and its valid values, or the inspection tool
to call first. No exception crosses the MCP boundary; the shared `Guard` choreography translates
exception *types* (never message text) into codes and hints. One tool extends (never replaces) the
shape: `render_preview_pages` returns the envelope as its first text content block, followed by MCP
image blocks.

The one documented boundary: the MCP SDK's own **argument-binding step** runs before any tool body,
so a malformed/wrong-typed argument produces the SDK's generic error, not this envelope. Calling
agents should treat a non-parseable `isError` result as exactly that case.

## Why

The consumer is an agent, not a human reading stderr. A uniform, machine-parseable failure shape
with a guaranteed next step is what makes self-correction loops work. Distinct exception types (e.g.
the dedicated `NotFoundException`) exist so a code like `not_found` cannot be produced accidentally
by an unrelated error — hints keyed on message text rot silently.

## Consequences

- **Envelope shape, tool names, parameters, and error codes are public API from v1.0.0** — changing
  or removing any of them is a semver-major event; additive fields and new codes are minor.
- New tools must route through the shared guards, not hand-roll try/catch.
- A PR that lets an exception escape a tool body, returns a bare payload, or emits an empty `hint`
  is a contract break, not a style issue.
