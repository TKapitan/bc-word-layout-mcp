# Builds the Channel-B release artifacts (release plan §2): one self-contained zip per Windows RID,
# a SHA256SUMS.txt covering them, and a copy of install.ps1 — the exact set a GitHub Release carries.
# CI runs this on tag; running it locally produces byte-equivalent artifacts for rehearsal.
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
    dotnet publish $project -c $Configuration -r $rid --self-contained | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "dotnet publish failed for $rid" }

    $publishDir = Join-Path $project "bin/$Configuration/net10.0/$rid/publish"
    $zipName = "bc-word-layout-mcp-$version-$rid.zip"
    $zipPath = Join-Path $out $zipName
    if (Test-Path $zipPath) { Remove-Item $zipPath -Force }

    # Zip the CONTENTS of publish/ (exe at zip root) — the whole folder must travel together.
    Compress-Archive -Path (Join-Path $publishDir '*') -DestinationPath $zipPath
    $hash = (Get-FileHash $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $sums += "$hash  $zipName"
    Write-Host "   $zipName  $([math]::Round((Get-Item $zipPath).Length / 1MB, 1)) MB"
}

Set-Content -Path (Join-Path $out 'SHA256SUMS.txt') -Value ($sums -join "`n") -NoNewline
Copy-Item (Join-Path $PSScriptRoot 'install.ps1') $out -Force
Write-Host "== artifacts in $out =="
Get-ChildItem $out | ForEach-Object { Write-Host "   $($_.Name)" }
