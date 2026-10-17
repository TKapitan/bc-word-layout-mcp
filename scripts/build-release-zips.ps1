# Builds the Channel-B release artifacts (release plan §2): one self-contained zip per Windows RID,
# a SHA256SUMS.txt covering them, and a copy of install.ps1 — the exact set a GitHub Release carries.
# CI and release.yml run this, then scripts/smoke-release-artifacts.ps1 checks what it built. Running it
# locally builds the same files for rehearsal; the zips are not byte-identical between runs, because
# Compress-Archive stamps each entry with its file's time.
#
#   pwsh scripts/build-release-zips.ps1 [-Rids win-x64,win-arm64] [-OutDir artifacts/release]

[CmdletBinding()]
param(
    [string[]]$Rids = @('win-x64', 'win-arm64'),
    [string]$OutDir = 'artifacts/release',
    [string]$Configuration = 'Release'
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$project = Join-Path $repoRoot 'src/BcWordLayout.McpHost'

[xml]$props = Get-Content (Join-Path $repoRoot 'Directory.Build.props')
$version = ($props.Project.PropertyGroup | Where-Object { $_.Version } | Select-Object -First 1).Version
if (-not $version) { throw 'Could not read <Version> from Directory.Build.props' }

$out = Join-Path $repoRoot $OutDir
New-Item -ItemType Directory -Force $out | Out-Null
$sums = @()

foreach ($rid in $Rids) {
    Write-Host "== publish $rid =="

    # A folder of its own, emptied first: the default bin/<config>/net10.0/<rid>/publish is shared with
    # the framework-dependent publish `dotnet pack` runs for the tool packages, so files from either could
    # end up in the other's output.
    $publishDir = Join-Path $repoRoot "artifacts/publish/$rid"
    if (Test-Path $publishDir) { Remove-Item $publishDir -Recurse -Force }

    # Quiet on success; on failure the publish output is the only record of the cause, so print it.
    $publishOutput = dotnet publish $project -c $Configuration -r $rid --self-contained -o $publishDir 2>&1
    if ($LASTEXITCODE -ne 0) {
        $publishOutput | ForEach-Object { Write-Host $_ }
        throw "dotnet publish failed for $rid (exit $LASTEXITCODE)"
    }

    $zipName = "bc-word-layout-mcp-$version-$rid.zip"
    $zipPath = Join-Path $out $zipName
    if (Test-Path $zipPath) { Remove-Item $zipPath -Force }

    # Zip the CONTENTS of the publish folder (exe at zip root) — the whole folder must travel together,
    # LICENSE and THIRD-PARTY-NOTICES.md included (the host project copies both into it).
    Compress-Archive -Path (Join-Path $publishDir '*') -DestinationPath $zipPath
    $hash = (Get-FileHash $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $sums += "$hash  $zipName"
    Write-Host "   $zipName  $([math]::Round((Get-Item $zipPath).Length / 1MB, 1)) MB"
}

Set-Content -Path (Join-Path $out 'SHA256SUMS.txt') -Value ($sums -join "`n") -NoNewline
Copy-Item (Join-Path $PSScriptRoot 'install.ps1') $out -Force
Write-Host "== artifacts in $out =="
Get-ChildItem $out | ForEach-Object { Write-Host "   $($_.Name)" }
