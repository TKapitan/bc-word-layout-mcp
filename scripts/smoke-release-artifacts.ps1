# Smoke-tests the release artifacts BEFORE anything is published (GitHub issue #110): the NuGet packages
# `dotnet pack` produced and the zips scripts/build-release-zips.ps1 produced must hold what they ship,
# and the server must start from both channels and report THIS version. release.yml runs it between
# building the artifacts and pushing to NuGet, whose versions are immutable; CI runs it on every pull
# request, so a packaging problem surfaces long before a tag could burn a version on it.
#
#   pwsh scripts/smoke-release-artifacts.ps1 [-PackageDir artifacts/nupkg] [-ZipDir artifacts/release]
#
# Needs Python 3 (tools/release/verify_server.py) and the .NET 10 SDK (dnx). Every check runs; the
# script fails at the end with the full list of what is wrong.

[CmdletBinding()]
param(
    [string]$PackageDir = 'artifacts/nupkg',
    [string]$ZipDir = 'artifacts/release'
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$rids = @('win-x64', 'win-arm64')

function Resolve-RepoPath([string]$Path) {
    if ([IO.Path]::IsPathRooted($Path)) { return $Path }
    return Join-Path $repoRoot $Path
}

[xml]$props = Get-Content (Join-Path $repoRoot 'Directory.Build.props')
$version = ($props.Project.PropertyGroup | Where-Object { $_.Version } | Select-Object -First 1).Version
if (-not $version) { throw 'Could not read <Version> from Directory.Build.props' }
$packageDir = Resolve-RepoPath $PackageDir
$zipDir = Resolve-RepoPath $ZipDir
Write-Host "Smoke-testing the $version release artifacts"

Add-Type -AssemblyName System.IO.Compression.FileSystem
$failures = [System.Collections.Generic.List[string]]::new()

function Get-ArchiveEntries([string]$Path) {
    $archive = [IO.Compression.ZipFile]::OpenRead($Path)
    try { return @($archive.Entries | ForEach-Object { $_.FullName }) }
    finally { $archive.Dispose() }
}

function Assert-Contains([string]$Archive, [string[]]$Entries, [string[]]$Required) {
    foreach ($entry in $Required) {
        if ($Entries -notcontains $entry) { $failures.Add("$(Split-Path $Archive -Leaf) does not contain $entry") }
    }
}

# 1. NuGet: the manifest package and one per RID ship together, each carrying the license and the
#    third-party notices (GitHub issue #108), and each RID package the runnable tool.
$expectedPackages = @("BcWordLayout.Mcp.$version.nupkg") + ($rids | ForEach-Object { "BcWordLayout.Mcp.$_.$version.nupkg" })
$packages = @(Get-ChildItem -LiteralPath $packageDir -Filter '*.nupkg' -File)
$unexpected = @($packages.Name | Where-Object { $expectedPackages -notcontains $_ })
$missing = @($expectedPackages | Where-Object { $packages.Name -notcontains $_ })
if ($unexpected.Count -or $missing.Count) {
    $failures.Add("$packageDir should hold exactly $($expectedPackages -join ', '); missing: $($missing -join ', '); unexpected: $($unexpected -join ', ')")
}
foreach ($package in $packages | Where-Object { $expectedPackages -contains $_.Name }) {
    $entries = Get-ArchiveEntries $package.FullName
    Assert-Contains $package.FullName $entries @('README.md', 'LICENSE', 'THIRD-PARTY-NOTICES.md', '.mcp/server.json')
    $rid = $rids | Where-Object { $package.Name -eq "BcWordLayout.Mcp.$_.$version.nupkg" }
    if ($rid) {
        $toolDir = "tools/net10.0/$rid"
        Assert-Contains $package.FullName $entries @("$toolDir/BcWordLayout.McpHost.dll", "$toolDir/LICENSE", "$toolDir/THIRD-PARTY-NOTICES.md")
    }
}

# 2. Zips: one self-contained zip per RID with the exe and the notices at its root, every one listed in
#    SHA256SUMS.txt with its actual hash, and the installer beside them.
$sumsPath = Join-Path $zipDir 'SHA256SUMS.txt'
$sums = @{}
if (Test-Path -LiteralPath $sumsPath) {
    foreach ($line in Get-Content -LiteralPath $sumsPath) {
        if ($line -match '^([0-9a-f]{64})\s+(\S+)$') { $sums[$Matches[2]] = $Matches[1] }
    }
}
else { $failures.Add("$zipDir has no SHA256SUMS.txt") }
if (-not (Test-Path -LiteralPath (Join-Path $zipDir 'install.ps1'))) { $failures.Add("$zipDir has no install.ps1") }
foreach ($rid in $rids) {
    $zipName = "bc-word-layout-mcp-$version-$rid.zip"
    $zipPath = Join-Path $zipDir $zipName
    if (-not (Test-Path -LiteralPath $zipPath)) { $failures.Add("$zipDir has no $zipName"); continue }
    Assert-Contains $zipPath (Get-ArchiveEntries $zipPath) @('BcWordLayout.McpHost.exe', 'LICENSE', 'THIRD-PARTY-NOTICES.md')
    $hash = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($sums[$zipName] -ne $hash) { $failures.Add("SHA256SUMS.txt gives $zipName as '$($sums[$zipName])', but its hash is $hash") }
}
$unlisted = @($sums.Keys | Where-Object { $_ -notmatch "^bc-word-layout-mcp-$([regex]::Escape($version))-(win-x64|win-arm64)\.zip$" })
if ($unlisted.Count) { $failures.Add("SHA256SUMS.txt lists files this release does not ship: $($unlisted -join ', ')") }

# 3. Start the server from both channels: it must answer the handshake with this version and list every
#    tool. Only this machine's own architecture can run, which on the windows-latest runner is win-x64.
# `python` first: on Windows, `python3` is often only the Microsoft Store stub.
$python = (Get-Command python -ErrorAction SilentlyContinue) ?? (Get-Command python3 -ErrorAction SilentlyContinue)
if (-not $python) { throw 'Python 3 is required for tools/release/verify_server.py.' }
$verify = Join-Path $repoRoot 'tools/release/verify_server.py'
$nativeRid = if ([Runtime.InteropServices.RuntimeInformation]::OSArchitecture -eq 'Arm64') { 'win-arm64' } else { 'win-x64' }
$scratch = Join-Path ([IO.Path]::GetTempPath()) "bcwl-smoke-$([guid]::NewGuid().ToString('N'))"
$savedNuGetPackages = $env:NUGET_PACKAGES
try {
    $nativeZip = Join-Path $zipDir "bc-word-layout-mcp-$version-$nativeRid.zip"
    if (Test-Path -LiteralPath $nativeZip) {
        Write-Host "== $nativeRid zip =="
        $unzipped = Join-Path $scratch 'zip'
        Expand-Archive -LiteralPath $nativeZip -DestinationPath $unzipped
        & $python.Source $verify --expect-version $version -- (Join-Path $unzipped 'BcWordLayout.McpHost.exe')
        if ($LASTEXITCODE -ne 0) { $failures.Add("the server from $(Split-Path $nativeZip -Leaf) failed verify_server.py (exit $LASTEXITCODE)") }
    }

    # A private, empty package cache: neither an earlier install of this version nor the machine's own
    # cache can answer, so dnx must take the packages from $packageDir - the ones about to be pushed.
    Write-Host '== dnx from the local packages =='
    $env:NUGET_PACKAGES = Join-Path $scratch 'nuget-packages'
    & $python.Source $verify --expect-version $version -- dnx "BcWordLayout.Mcp@$version" --yes --source $packageDir
    if ($LASTEXITCODE -ne 0) { $failures.Add("the server from the local NuGet packages (dnx) failed verify_server.py (exit $LASTEXITCODE)") }
}
finally {
    $env:NUGET_PACKAGES = $savedNuGetPackages
    if (Test-Path -LiteralPath $scratch) { Remove-Item -LiteralPath $scratch -Recurse -Force -ErrorAction SilentlyContinue }
}

if ($failures.Count) {
    $failures | ForEach-Object { Write-Host "FAIL: $_" }
    throw "$($failures.Count) release artifact check(s) failed; nothing may be published."
}
Write-Host "OK: the $version packages and zips are complete and both channels start the server."
