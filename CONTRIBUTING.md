# Contributing

Thanks for considering a contribution. This project has a small, deliberate design — the fastest
way to land a change is to work *with* that design, so please read this page (2 minutes) before
opening a PR.

## Getting started

Prerequisites: the **.NET 10 SDK**. Nothing else is required — Microsoft Word / LibreOffice are
optional (they only affect `preview_layout`'s PDF step, and the converter tests pass either way).

```pwsh
dotnet build
dotnet test
```

The full suite must be green before and after your change. Tests locate the corpus via
`AppContext.BaseDirectory`, so they run from any directory. Optional Python 3 tooling under
`tools/e2e/` (not required for a PR, but a good smoke test for tool-surface changes):
`scenarios.py` drives 24 multi-step edit journeys through the real stdio server;
`call.py` makes a one-shot tool call from the CLI (handy while debugging);
`author_order_confirmation.py` is the worked example that authors a complete layout from scratch
through the tools alone; `sandbox_pack.py`/`bc_compare.py` build and re-check the BC-sandbox
validation pack (see `docs/FIDELITY-CHECKLIST.md`). The Python tools have their own standard-library
tests, which CI runs on its Linux leg: `python -m unittest discover -s tools/tests`.

### Running your build as an MCP server

The repo-root [`.mcp.json`](.mcp.json) launches the server from your working tree with
`dotnet run --project src/BcWordLayout.McpHost --artifacts-path .mcp-run`, so there is no publish step
while iterating. The separate `.mcp-run/` folder matters on Windows: a running server locks the DLLs it
loaded, and building into the normal `bin/` folders would block `dotnet build` and `dotnet test` while a
session is open.

Claude Code starts every `.mcp.json` server when a session opens, so the committed
[`.claude/settings.json`](.claude/settings.json) disables this one (`disabledMcpjsonServers`). To use
your build, add it under your own local scope, which that setting does not cover:

```pwsh
claude mcp add --scope local bc-word-layout-dev -- dotnet run --project src/BcWordLayout.McpHost --artifacts-path .mcp-run
```

Remove it with `claude mcp remove --scope local bc-word-layout-dev`. The server runs the code as built
when the session started; restart the session to pick up later edits.

To check a build by hand: the server speaks newline-delimited JSON-RPC on stdout and logs to stderr.
Run `dotnet run --project src/BcWordLayout.McpHost` and send one object per line:

```json
{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"smoke","version":"1.0"}}}
{"jsonrpc":"2.0","method":"notifications/initialized"}
{"jsonrpc":"2.0","id":2,"method":"tools/list","params":{}}
```

### Publishing and packing

- `dotnet publish src/BcWordLayout.McpHost -c Release -r win-x64 --self-contained` builds a
  self-contained deployment in `src/BcWordLayout.McpHost/bin/Release/net10.0/win-x64/publish/`. Ship
  the whole folder: the `.exe` depends on the `.dll`s beside it.
- `dotnet pack src/BcWordLayout.McpHost -c Release` builds the top-level `BcWordLayout.Mcp` package
  (with the `.mcp/server.json` manifest nuget.org shows as install config) and one package per Windows
  RID; they are always pushed together.
- `scripts/build-release-zips.ps1` builds the per-RID zips and checksums a GitHub Release carries, and
  `python tools/release/verify_server.py -- <command>` smoke-tests any install channel.
- Releases follow [`docs/RELEASING.md`](docs/RELEASING.md); `PackagingMetadataTests` fails the suite
  when a version bump misses one of the manifests.

### Project layout

| Project | Role |
|---|---|
| `src/BcWordLayout.Domain` | Reading, validating and editing layouts: schema parsing, the control inventory, quick validation, the edit and table engines, formatting, `create_layout` and `refresh_xml_part`. |
| `src/BcWordLayout.Merge` | Sample data, the merge engine behind the preview, and the dry-run merge behind `validate_layout level=full`. |
| `src/BcWordLayout.Render` | PDF conversion: Word COM (primary) and LibreOffice (untested fallback). |
| `src/BcWordLayout.McpHost` | The MCP stdio server: the tool definitions (`Tools/*Tools.cs`, sharing `Tools/ToolGuards.cs`), their arguments, the `{ok,data,error}` envelope (`ToolContracts.cs`), and the call filter that keeps malformed calls inside it (`ToolCallBoundary.cs`). |
| `tests/BcWordLayout.Tests` | xUnit tests over the real corpus layouts in `tests/corpus/` (see its `PROVENANCE.md`) and synthetic fixtures, including Word COM conversion tests that pass whether or not Word is installed. |

## The design constraints (read before proposing)

[`docs/SOLUTION-DESIGN.md`](docs/SOLUTION-DESIGN.md) is the source of truth for design decisions —
where it and the code disagree, one of the two is a bug. The decisions most often "fixed" by
well-meaning PRs are written down as ADRs in [`docs/adr/`](docs/adr/); PRs that reverse one without
first making the case in an issue will be declined, however clean the code:

1. **Synchronous handlers, no DI container** — single-user local process; simplicity is the feature.
2. **Never open the original writable** — every mutation stages a copy and commits atomically.
3. **Refuse rather than guess** — ambiguous edits return an error with a hint, not a best effort.
4. **One `{ok, data, error}` envelope** — nothing throws across the MCP boundary; error `hint`s are
   never empty. Tool names, parameters, and error codes are a public API (semver).
5. **Emit only OOXML that has been observed in a real BC layout** — "valid per the OOXML schema" is
   not the bar; a new shape needs a captured fixture first.
6. **Transplant the dataset schema, never synthesize it** — the BC dataset part always comes from a
   BC-produced artifact (a layout or an exported schema), never from AL source or by hand.

## Corpus rules

The `.docx` files in [`tests/corpus/`](tests/corpus/) are real Business Central layout captures.
They are evidence, not fixtures of convenience:

- **Never re-save them** through Word or any tool that rewrites the package — their value is being
  byte-genuine BC output.
- **Adding a capture?** Run `python tools/scrub_corpus_metadata.py` before committing (it strips
  personal/organizational metadata), and add a row to
  [`tests/corpus/PROVENANCE.md`](tests/corpus/PROVENANCE.md) saying what the file is and where it
  came from.
- **Never commit a layout containing customer, client, or otherwise confidential material** — the
  same applies to issue attachments.

## Tests

- New behavior needs tests asserting the resulting OOXML, not just the response envelope.
- Validator checks come in pairs: each defect fixture is matched by a valid sibling that must **not**
  trip the same check — a check that fires indiscriminately is worse than none.
- Snapshot changes must be re-baselined deliberately: review each diff, never blanket-accept.

## Pull requests

- Branch from `main` (`feature/…`, `fix/…`, `docs/…`, `chore/…`) — branch naming, the change
  lifecycle, and how PRs merge are in [`docs/BRANCHING.md`](docs/BRANCHING.md).
- One functionality per PR; keep the build warning-free.
- Update [`CHANGELOG.md`](CHANGELOG.md) (add an *Unreleased* section if there isn't one) for anything
  user-visible.
- If your change alters a decision recorded in `docs/SOLUTION-DESIGN.md` or an ADR, update that
  document in the same PR — that is the contract stated at the top of the design doc.

## Licensing

This project is MIT-licensed. By contributing you agree your contribution is licensed under the
same terms (inbound = outbound). There is no CLA.
