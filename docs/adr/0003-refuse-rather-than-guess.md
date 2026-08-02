# ADR-0003: Refuse rather than guess

**Status:** Accepted (v1.0.0) · **Source:** SOLUTION-DESIGN §6.3

## Description

When an edit has more than one defensible outcome, the tool **refuses and says why** (an error with
a non-empty `hint`), rather than picking one. Standing examples: `insert_column` at a position that
falls inside a spanned cell of a row that must carry the new content (widening silently drops the
requested field; splitting makes a layout decision nobody asked for); `create_layout` with a
`templatePath` that already carries bound controls; creating nested repeaters in a header/footer.

## Why

The caller is an AI agent. An agent can react to a clear refusal — adjust the argument, pick another
tool, ask its user — but it cannot react to a silently wrong guess, because nothing tells it a guess
was made. The wrong guess surfaces days later as a corrupt or wrong deliverable in a BC sandbox,
attributed to nobody. The cost of a refusal is one round-trip with a hint; the cost of a guess is
trust.

## Consequences

- Some edits take two calls (inspect, then edit with explicit choices). That is the intended shape.
- Error `hint`s are load-bearing: every refusal names the argument to fix, its valid values, or the
  inspection tool to call first.
- A PR replacing a refusal with a heuristic ("just widen the spanned cell") must show the choice has
  become genuinely unambiguous — not merely common.
