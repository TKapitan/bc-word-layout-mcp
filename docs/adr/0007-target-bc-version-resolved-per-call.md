# ADR-0007: The target BC version is resolved per call, never stored

**Status:** Accepted (GitHub issue #152) · **Source:** SOLUTION-DESIGN §6.10

## Description

Some layout features need a newer Business Central: report themes and themable formatting need BC 29,
and later features will need 30 or 31. The **target** is the *minimum* BC version a layout must work on,
and every call resolves it again; the first match wins:

1. The nearest AL `app.json` above the layout (above `outputPath` for `create_layout`): its
   `application`, e.g. `29.0.0.0`. The walk stops at the first `app.json`, at a folder containing `.git`,
   or at the drive root. An `app.json` that cannot be read, or names something that is not a version,
   makes the target unknown; one with no `application` falls through.
2. `BCWL_BC_VERSION` (e.g. `29` or `29.1`) in the server's `env` block, for layouts outside an AL project.
3. Unknown, with the reason.

The server keeps no target between calls. There is no init tool and no per-call parameter.
`get_layout_info` and `validate_layout` report the target they used, with its source.

What the target controls:

- A tool that would **author** a catalog feature (`BcFeatures`) refuses with `bc_version_unsupported`
  below the feature's version, and when the target is unknown.
- `validate_layout`'s `bc-version` check reports a catalog feature a layout **already has**: a header/footer
  layout bound to Tenant Report Defaults is an Error below 29 and a Warning for an unknown target. Theme
  colours and paragraph styles already in a layout are not flagged: BC 28 prints them, and Microsoft's own
  BC 28 layouts carry them.
- Read tools never refuse.

## Why

- **The project already states the version.** `application` in `app.json` is exactly the minimum version an
  AL app supports. A version stored in the server, set by an init tool or passed per call, could disagree with
  it, and nothing would tell the agent which one won. One server also serves several projects at once.
- **Per call, not at startup.** Editing `app.json` takes effect on the next call, with no restart, and two
  layouts in two projects get their own targets.
- **No per-call parameter.** Every mutating response carries a validation summary that depends on the target,
  so the parameter would have to be on all 28 tools, and an agent would have to pass it consistently.
- **Unknown refuses (ADR-0003).** Allowing a gated feature for an unknown target is a guess that surfaces as a
  layout BC cannot use. Refusing costs one question to the user.
- **major.minor.** BC behaviour can change within a major version: 28.0 and 28.3 accepted an external attached
  template, and a later 28.x refused it.

## Consequences

- An agent reads `target` first. When it is unknown and the task needs a gated feature, the agent asks the
  user which BC version the layout is for, and suggests setting `application` or `BCWL_BC_VERSION`.
- Every `BcFeatures` entry states its evidence, and a capability starts at the first version *confirmed* to have
  it. The README's version table lists the catalog, and a test keeps the two equal.
- A layout outside any AL project with no `BCWL_BC_VERSION` has an unknown target: header/footer layouts get a
  warning, and gated features are refused.
- A proposal to store the target, or to take it per call, must show how it stays consistent with the project's
  own `app.json`.
