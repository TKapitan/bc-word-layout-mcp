# Installs the bc-word-layout-mcp server from a GitHub Release to a STABLE path, so the documented
# MCP config block never depends on where a zip was unpacked (release plan §2 Channel B).
#
#   irm https://github.com/TKapitan/bc-word-layout-mcp/releases/latest/download/install.ps1 | iex
#
# or, pinned / offline:
#
#   pwsh scripts/install.ps1 [-Version 1.0.0] [-FromDirectory <dir with zips + SHA256SUMS.txt>]
#
# What it does — deliberately nothing more (D5: keep Channel B dumb):
#   1. picks the zip for this machine's architecture (win-x64 / win-arm64),
#   2. downloads it + SHA256SUMS.txt (or reads them from -FromDirectory) and VERIFIES the checksum,
#   3. refuses to proceed while an installed copy is running,
#   4. unpacks to %LOCALAPPDATA%\bc-word-layout-mcp\<version>\ and points ...\current at it,
#   5. prints the ready-to-paste MCP config block. No auto-update, no PATH edits, no telemetry.

[CmdletBinding()]
param(
    [string]$Version = 'latest',
    [string]$Repo = 'TKapitan/bc-word-layout-mcp',
    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA 'bc-word-layout-mcp'),
    [string]$FromDirectory = ''
)

$ErrorActionPreference = 'Stop'

$rid = if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') { 'win-arm64' } else { 'win-x64' }

if ($FromDirectory) {
    $zip = Get-ChildItem $FromDirectory -Filter "bc-word-layout-mcp-*-$rid.zip" | Select-Object -First 1
    if (-not $zip) { throw "No bc-word-layout-mcp-*-$rid.zip in $FromDirectory" }
    if ($zip.Name -notmatch "bc-word-layout-mcp-(.+)-$rid\.zip") { throw "Unexpected zip name $($zip.Name)" }
    $resolvedVersion = $Matches[1]
    $zipPath = $zip.FullName
    $sumsPath = Join-Path $FromDirectory 'SHA256SUMS.txt'
}
else {
    if ($Version -eq 'latest') {
        $release = Invoke-RestMethod "https://api.github.com/repos/$Repo/releases/latest"
        $resolvedVersion = $release.tag_name.TrimStart('v')
    }
    else {
        $resolvedVersion = $Version.TrimStart('v')
    }
    $zipName = "bc-word-layout-mcp-$resolvedVersion-$rid.zip"
    $base = "https://github.com/$Repo/releases/download/v$resolvedVersion"
    $tempDir = Join-Path ([IO.Path]::GetTempPath()) "bcwl-install-$([guid]::NewGuid().ToString('N'))"
    New-Item -ItemType Directory $tempDir | Out-Null
    $zipPath = Join-Path $tempDir $zipName
    $sumsPath = Join-Path $tempDir 'SHA256SUMS.txt'
    Write-Host "Downloading $zipName ..."
    Invoke-WebRequest "$base/$zipName" -OutFile $zipPath
    Invoke-WebRequest "$base/SHA256SUMS.txt" -OutFile $sumsPath
}

# 2. checksum — refuse anything that does not match the published sum.
$expectedLine = (Get-Content $sumsPath) | Where-Object { $_ -match [regex]::Escape((Split-Path $zipPath -Leaf)) }
if (-not $expectedLine) { throw "SHA256SUMS.txt has no entry for $(Split-Path $zipPath -Leaf)" }
$expected = ($expectedLine -split '\s+')[0].ToLowerInvariant()
$actual = (Get-FileHash $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
if ($actual -ne $expected) { throw "Checksum mismatch for $(Split-Path $zipPath -Leaf): expected $expected, got $actual" }
Write-Host 'Checksum OK.'

# 3. never modify an install that is currently running.
$exePath = Join-Path $InstallRoot 'current\BcWordLayout.McpHost.exe'
$running = Get-Process -Name 'BcWordLayout.McpHost' -ErrorAction SilentlyContinue |
    Where-Object { $_.Path -and $_.Path.StartsWith($InstallRoot, [System.StringComparison]::OrdinalIgnoreCase) }
if ($running) {
    throw "bc-word-layout-mcp is currently running (PID $($running.Id -join ', ')) from $InstallRoot. Close the MCP clients using it and re-run."
}

# 4. unpack to a versioned directory and flip 'current' to it.
$versionDir = Join-Path $InstallRoot $resolvedVersion
if (Test-Path $versionDir) { Remove-Item $versionDir -Recurse -Force }
New-Item -ItemType Directory -Force $versionDir | Out-Null
Expand-Archive $zipPath -DestinationPath $versionDir
if (-not (Test-Path (Join-Path $versionDir 'BcWordLayout.McpHost.exe'))) {
    throw "The zip did not contain BcWordLayout.McpHost.exe at its root - refusing to activate it."
}

$current = Join-Path $InstallRoot 'current'
if (Test-Path $current) { (Get-Item $current).Delete() }   # deletes the junction only, never a target dir
New-Item -ItemType Junction -Path $current -Target $versionDir | Out-Null

# 5. the block users paste — the whole point of the stable path.
Write-Host ''
Write-Host "Installed bc-word-layout-mcp $resolvedVersion -> $current"
Write-Host ''
Write-Host 'VS Code (.vscode/mcp.json or user mcp.json):'
Write-Host @"
{
  "servers": {
    "bc-word-layout": {
      "type": "stdio",
      "command": "$($exePath -replace '\\', '\\')"
    }
  }
}
"@
Write-Host 'Claude Code (.mcp.json, or: claude mcp add bc-word-layout -- <path below>):'
Write-Host @"
{
  "mcpServers": {
    "bc-word-layout": {
      "command": "$($exePath -replace '\\', '\\')"
    }
  }
}
"@
