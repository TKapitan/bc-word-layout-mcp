# Branching & Change Process

How a change travels from idea to `main`, and the guardrails on the way. This page covers the
*mechanics* — [`CONTRIBUTING.md`](../CONTRIBUTING.md) covers what a good change looks like, and
[`RELEASING.md`](RELEASING.md) covers how a release ships once the change is on `main`.

## Branch structure

**`main` is the only long-lived branch**, and it must stay releasable at all times: full suite
green, zero warnings (CI enforces both on every PR). Releases are annotated `v*` tags on `main`.
There is no `develop`, no release branches, no environment branches — with a single maintainer,
tag-driven releases, and immutable NuGet versions, extra long-lived branches would add ceremony
without adding safety.

Everything else is a **short-lived topic branch**:

| Prefix     | Use for                                    |
| ---------- | ------------------------------------------ |
| `feature/` | new tools, parameters, behavior            |
| `fix/`     | bug fixes                                  |
| `docs/`    | documentation-only changes                 |
| `chore/`   | build, CI, dependencies, internal tooling  |

Branch from up-to-date `main`, keep the slug short and imperative, and lead with the issue number
when one exists: `fix/23-nested-table-anchor`, `feature/31-header-image-tool`. Topic branches merge
via PR and are deleted on merge. A topic branch more than a couple of weeks old is a smell —
rebase it onto `main` or split the work.

### The one exception: hotfix branches

If the latest release is broken **and** `main` already carries unreleased work that should not
ship, branch from the release tag instead of `main`:

```pwsh
git switch -c fix/1.2.4-broken-preview v1.2.3
```

Apply the fix, bump the version per [`RELEASING.md`](RELEASING.md) §1, and tag on the branch —
[`release.yml`](../.github/workflows/release.yml) triggers on the tag push, not on the branch, so
the release publishes from the hotfix branch exactly as it would from `main`. Then merge the
branch back into `main` via a normal PR so the fix and version bump aren't lost. Until this
situation actually occurs, no standing hotfix machinery exists — that is deliberate.

## Handling a new change

1. **Start from an issue.** The backlog lives in
   [GitHub issues](https://github.com/TKapitan/bc-word-layout-mcp/issues); anything non-trivial
   (new behavior, design changes, corpus additions) gets an issue first so scope is agreed before
   code exists. Typo-grade fixes may go straight to a PR.
2. **Branch** from fresh `main` using the naming above.
3. **Make the change** following [`CONTRIBUTING.md`](../CONTRIBUTING.md): one functionality per
   branch, tests asserting the resulting OOXML, `CHANGELOG.md` *Unreleased* entry for anything
   user-visible, and design-doc/ADR updates in the same change when a recorded decision moves.
4. **Open a PR against `main`.** The template checklist is the review contract. The PR title
   becomes the commit subject on `main` (squash merge), so write it in the imperative mood, like
   the existing history.
5. **Gates.** CI must be green on both legs (`build-and-test (windows-latest)` and
   `build-and-test (ubuntu-latest)`), the branch must be current with `main`, and every review
   conversation must be resolved. A Markdown-only change skips both build legs by design
   ([`ci.yml`](../.github/workflows/ci.yml)'s `changes` job): the required checks report as
   *skipped*, which GitHub counts as satisfied — documentation PRs merge without spending
   build minutes. External PRs are reviewed by the code owner
   ([`CODEOWNERS`](../.github/CODEOWNERS)); maintainer PRs merge on green CI — the required-check
   gate, not self-review theater, is what protects `main` while there is one maintainer.
6. **Squash-merge.** One PR → one commit on `main`. The head branch is deleted automatically.
7. **Release** whenever `main` is worth shipping, per [`RELEASING.md`](RELEASING.md).

**Dependabot PRs** ([`dependabot.yml`](../.github/dependabot.yml)) follow the same path: CI gates
them, the maintainer merges them. No auto-merge — a dependency bump that changes emitted OOXML or
converter behavior deserves eyes.

## Branch protection rules

Protection is defined as two GitHub **rulesets**, kept as import-ready JSON under
[`.github/rulesets/`](../.github/rulesets/) so the intended configuration is reviewable in the
repo. The JSON files are documentation-plus-import-source — GitHub does not read them from the
repo; apply them via the appendix below or Settings → Rules → Rulesets → Import.

### Ruleset `main` ([`branch-main.json`](../.github/rulesets/branch-main.json)) — targets the default branch

| Rule | Setting | Why |
| --- | --- | --- |
| Require a pull request | required approvals: **0** | A solo maintainer cannot approve their own PR; requiring 1 would dead-lock every merge. Raise to 1 + require code-owner review the day a second maintainer joins. |
| Allowed merge methods | **squash only** | One functionality per PR → one commit per functionality on `main`; PR titles feed `--generate-notes` in the release workflow. |
| Required status checks | `build-and-test (windows-latest)`, `build-and-test (ubuntu-latest)`, **strict** (branch must be up to date) | Both CI legs are load-bearing: Windows is the supported platform, Linux is the rot-guard (see [`ci.yml`](../.github/workflows/ci.yml)). Strict mode means the merged result is exactly what CI tested. Markdown-only changes skip both legs via a job-level `if:` — a skipped required job counts as passing, so docs PRs still merge. |
| Require conversation resolution | on | Review threads are part of the record. |
| Require linear history | on | Squash-only already guarantees it; the rule keeps it true if merge settings ever drift. |
| Block force pushes / deletions | on | `main` is the release source of truth. |
| Bypass actors | **none** | Direct pushes to `main` stop, maintainer included. The escape hatch is editing the ruleset (admin-only, audit-logged), not a standing exemption that silently erodes the process. |

### Ruleset `release-tags` ([`tags-releases.json`](../.github/rulesets/tags-releases.json)) — targets `v*` tags

| Rule | Setting | Why |
| --- | --- | --- |
| Restrict creation | on, bypass: **repository admin** | Pushing a `v*` tag publishes immutable, unreclaimable NuGet packages. Only the maintainer may pull that trigger. |
| Block updates / deletions / force pushes | on | Release tags are permanent record; `gh release create --verify-tag` and the tag-vs-version guard both assume the tag never moves. |

### Consequence for the release runbook

With direct pushes to `main` blocked, [`RELEASING.md`](RELEASING.md) §1–3 changes shape slightly:
the version bump + changelog roll happen on a `chore/release-vX.Y.Z` branch and merge via PR like
any other change; the maintainer then tags the resulting squash commit on `main` and pushes **only
the tag** (`git push origin v1.2.3` — the `git push origin main --follow-tags` form would be
rejected). The tag push passes the `release-tags` ruleset via the admin bypass.

## Repository settings that back this up

Rulesets do not cover merge-button behavior; set these once in Settings → General:

- **Allow squash merging** only — disable merge commits and rebase merging.
- Default squash commit message: **pull request title and description**.
- **Automatically delete head branches** — on.

## Appendix: applying the configuration

```pwsh
# Merge-button behavior + auto-delete (Settings → General equivalents)
gh api -X PATCH repos/TKapitan/bc-word-layout-mcp `
  -F allow_merge_commit=false -F allow_rebase_merge=false -F allow_squash_merge=true `
  -F delete_branch_on_merge=true `
  -f squash_merge_commit_title=PR_TITLE -f squash_merge_commit_message=PR_BODY

# The two rulesets
gh api -X POST repos/TKapitan/bc-word-layout-mcp/rulesets --input .github/rulesets/branch-main.json
gh api -X POST repos/TKapitan/bc-word-layout-mcp/rulesets --input .github/rulesets/tags-releases.json
```

To verify: `gh api repos/TKapitan/bc-word-layout-mcp/rulesets` should list both rulesets as
`active`, and a test push of a commit directly to `main` should be rejected.
