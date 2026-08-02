# Releasing

The maintainer runbook for cutting a release. CI does the heavy lifting on a tag push
([`release.yml`](../.github/workflows/release.yml)); the manual part is the version bump, the
changelog, and the post-publish verification.

## 0. Preflight (from a clean clone of `main`)

```pwsh
dotnet build -c Release                    # 0 errors — warnings fail the build by design
dotnet test  -c Release                    # full suite green
gitleaks detect --log-opts=--no-textconv   # MUST report "commits scanned" > 0; expect no findings
```

The `--log-opts=--no-textconv` flag matters: Git-for-Windows' system-level `astextplain` textconv
chokes on the corpus `.docx` diffs and silently makes gitleaks scan **0 commits** while still
reporting "no leaks found" — always check the "commits scanned" line before trusting the result.

## 1. Version bump — every release, release candidates included

Bump the version in **all of these** (the suite fails via `PackagingMetadataTests` if any is
missed):

1. `Directory.Build.props` — `<Version>` and `<InformationalVersion>` (keep `AssemblyVersion` /
   `FileVersion` at the numeric `X.Y.Z.0` — they cannot carry a `-rc.N` suffix).
2. `src/BcWordLayout.McpHost/.mcp/server.json` — `version`, and `packages[0].version`.
3. `.claude-plugin/plugin.json` **and** root `plugin.json` — `version`, and the pinned
   `BcWordLayout.Mcp@<version>` dnx reference in `mcpServers` (both manifests).

**Release candidates are not exempt.** An RC gets a real bump to `X.Y.Z-rc.N` in all the files
above, exactly like a final release — the packed NuGet version comes from
`Directory.Build.props`, never from the tag, and NuGet versions are immutable once pushed. The
workflow's first step fails the run if the tag and `<Version>` disagree, so a stale version can
never burn the wrong (unreclaimable) package version. After the last RC, bump back to the plain
`X.Y.Z` for the final release (the RC suffix versions stay consumed; the plain version is still
free).

## 2. Changelog

Move the *Unreleased* section of [`CHANGELOG.md`](../CHANGELOG.md) to the new version and date it.

## 3. Tag

```pwsh
git tag -a v1.2.3 -m "v1.2.3"
git push origin main --follow-tags
```

The tag must be `v` + the exact `<Version>` from `Directory.Build.props` — `release.yml` verifies
this before doing anything else. A tag containing `-` (e.g. `v1.2.3-rc.1`) is published as a
**pre-release**.

## 4. CI takes over

`release.yml` runs the full test gate, packs and pushes **every** produced `.nupkg` (the top-level
`BcWordLayout.Mcp` manifest package plus one per RID — they must ship together), builds the per-RID
zips + `SHA256SUMS.txt` via `scripts/build-release-zips.ps1`, and creates the GitHub Release with
the zips, checksums, and `install.ps1` as assets.

NuGet publishing uses **Trusted Publishing (OIDC)** — no stored API key. Two pieces of setup, once:
a Trusted Publishing policy on nuget.org (account → Trusted Publishing → repository owner
`TKapitan`, repository `bc-word-layout-mcp`, workflow file `release.yml`), and the `NUGET_USER`
repo secret holding the nuget.org profile name. A newly created policy is only pending-active for
7 days until its first successful publish, so if releases are far apart and the policy shows as
inactive, restart its window on nuget.org before tagging.

## 5. Post-publish verification — on a machine that is not yours

```pwsh
# Channel A: dnx fetches from NuGet and the server answers with all 23 tools
python tools/release/verify_server.py -- dnx BcWordLayout.Mcp@1.2.3 --yes

# Channel B: the installer downloads, checksum-verifies, unpacks, and the installed exe works
irm https://github.com/TKapitan/bc-word-layout-mcp/releases/latest/download/install.ps1 | iex

# Verifying an RC: `releases/latest` excludes pre-releases, so download install.ps1 from the RC
# release's own asset URL and pass -Version explicitly, e.g.
#   & ([scriptblock]::Create((irm https://github.com/TKapitan/bc-word-layout-mcp/releases/download/v1.2.3-rc.1/install.ps1))) -Version 1.2.3-rc.1
python tools/release/verify_server.py -- "$env:LOCALAPPDATA\bc-word-layout-mcp\current\BcWordLayout.McpHost.exe"

# A real tool call against a sample layout
#   get_layout_info on any corpus layout → ok:true
```

## Rollback

NuGet packages can be **unlisted, never deleted**; GitHub Releases can be deleted. A bad build is
recoverable — a leaked file is not. That asymmetry is why the corpus rules in
[`CONTRIBUTING.md`](../CONTRIBUTING.md) and the scrub script gate everything that enters
`tests/corpus/`.
