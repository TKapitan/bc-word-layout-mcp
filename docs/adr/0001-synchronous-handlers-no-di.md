# ADR-0001: Synchronous tool handlers, no DI container

**Status:** Accepted (v1.0.0) · **Source:** SOLUTION-DESIGN §3 (technology choices), §6.7

## Context

This is a single-user, local, stdio MCP server: one process, one agent, file-sized work items. The
natural .NET reflexes — `async` handlers end to end, a DI container, hosted-service abstractions —
exist to serve concurrency and substitutability this process does not have.

## Decision

Tool handlers are synchronous. There is no DI container; construction is direct. Concurrency is
handled where it actually exists: a per-path in-process lock serializes concurrent edits to the same
file, and a `CrossProcessLock` (named mutex, acquired *inside* the in-process lock) serializes
separate host processes — e.g. two IDE windows — touching the same path. The one seam that needed
substitution, PDF converter selection, is an explicit parameter (`converter=auto|word|libreoffice`),
not an injected interface graph.

## Consequences

- The call path is readable top to bottom; a stack trace names the actual work.
- Tests construct objects directly; there is no container configuration to drift.
- Async-ification or introducing a container is *not* an improvement here. A PR proposing either
  must demonstrate a user-visible problem it solves (e.g. a real workload where blocking a tool call
  hurts an MCP client) — "it's more idiomatic" does not clear the bar.
