# Installs the bc-word-layout-mcp server from a GitHub Release to a STABLE path, so the documented
# MCP config block never depends on where a zip was unpacked (release plan §2 Channel B).
#
#   irm https://github.com/TKapitan/bc-word-layout-mcp/releases/latest/download/install.ps1 | iex
#
# or a specific version (pre-releases included), with that release's own copy of this script:
#
#   & ([scriptblock]::Create((irm https://github.com/TKapitan/bc-word-layout-mcp/releases/download/v1.2.3/install.ps1))) -Version 1.2.3
#
# or pinned / offline, from a file:
#
#   pwsh scripts/install.ps1 [-Version 1.0.0] [-FromDirectory <dir with zips + SHA256SUMS.txt>]
#
# -FromDirectory installs exactly the -Version zip from that folder or, without -Version, the
# highest SemVer version among the folder's zips for this architecture.
#
# What it does — deliberately nothing more (D5: keep Channel B dumb):
#   1. picks the zip for this machine's NATIVE architecture (win-x64 / win-arm64),
#   2. downloads it + SHA256SUMS.txt (or reads them from -FromDirectory) and VERIFIES the checksum,
#   3. refuses to proceed while an installed copy is running,
#   4. unpacks to %LOCALAPPDATA%\bc-word-layout-mcp\<version>\ and points ...\current at it,
#   5. prints the ready-to-paste MCP config block. No auto-update, no PATH edits, no telemetry.
# A download lands in a %TEMP%\bcwl-install-* folder that is removed again however the run ends.
#
# Why the whole body is one scriptblock that holds the param block: `irm | iex` runs a script's top
# level IN THE CALLER'S SCOPE, so a script-level param block would overwrite the caller's own
# $Version, and $ErrorActionPreference = 'Stop' and every variable would stay behind in their
# session. `& { param(...) ... } @args` gives the body a scope of its own however it is started, and
# @args hands it the arguments of a file run or of the scriptblock form above. For the same reason
# the body never calls `exit` (under iex that closes the caller's shell; it throws instead) and never
# relies on $PSScriptRoot or its own file path (there is none under iex).

& {
    [CmdletBinding()]
    param(
        [string]$Version = 'latest',
        [string]$Repo = 'TKapitan/bc-word-layout-mcp',
        [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA 'bc-word-layout-mcp'),
        [string]$FromDirectory = ''
    )

    $ErrorActionPreference = 'Stop'
    # Windows PowerShell 5.1 redraws Invoke-WebRequest's progress bar so often that it slows the
    # ~50 MB download down many times over.
    $ProgressPreference = 'SilentlyContinue'

    # SemVer 2.0 precedence, for picking the highest -FromDirectory zip: Windows PowerShell 5.1 has no
    # [semver], and [version] cannot parse '1.2.0-rc.1'. ConvertTo-SemVer returns $null for anything
    # that is not MAJOR.MINOR.PATCH[-PRERELEASE][+BUILD]; build metadata never affects the order.
    function ConvertTo-SemVer([string]$Text) {
        $id = '(?:0|[1-9][0-9]*|[0-9]*[A-Za-z-][0-9A-Za-z-]*)'
        if ($Text -notmatch "^(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)(?:-($id(?:\.$id)*))?(?:\+[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?$") {
            return $null
        }
        $pre = @()
        if ($Matches[4]) { $pre = $Matches[4].Split('.') }
        [pscustomobject]@{ Core = @($Matches[1], $Matches[2], $Matches[3]); Pre = $pre }
    }

    # -1, 0 or 1. Numeric identifiers compare as numbers (SemVer forbids leading zeros, so the longer
    # one is the larger), the others ordinally, and a numeric one sorts below any other.
    function Compare-Identifier([string]$A, [string]$B) {
        $aIsNumber = $A -match '^[0-9]+$'
        $bIsNumber = $B -match '^[0-9]+$'
        if ($aIsNumber -ne $bIsNumber) { if ($aIsNumber) { return -1 } return 1 }
        if ($aIsNumber -and $A.Length -ne $B.Length) { return [math]::Sign($A.Length - $B.Length) }
        [math]::Sign([string]::CompareOrdinal($A, $B))
    }

    # -1, 0 or 1. A pre-release sorts below its release; pre-release identifiers compare left to
    # right, and when all the shared ones are equal the shorter list sorts first.
    function Compare-SemVer($A, $B) {
        for ($i = 0; $i -lt 3; $i++) {
            $order = Compare-Identifier $A.Core[$i] $B.Core[$i]
            if ($order) { return $order }
        }
        if (-not $A.Pre.Count -or -not $B.Pre.Count) { return [math]::Sign($B.Pre.Count - $A.Pre.Count) }
        for ($i = 0; $i -lt [math]::Min($A.Pre.Count, $B.Pre.Count); $i++) {
            $order = Compare-Identifier $A.Pre[$i] $B.Pre[$i]
            if ($order) { return $order }
        }
        [math]::Sign($A.Pre.Count - $B.Pre.Count)
    }

    # 1. the NATIVE architecture. x86 PowerShell finds it in PROCESSOR_ARCHITEW6432, but x64 PowerShell
    # emulated on an ARM64 machine is told AMD64, so Win32_Processor (Architecture 12 = ARM64) has the
    # last word whenever the query works.
    $arch = if ($env:PROCESSOR_ARCHITEW6432) { $env:PROCESSOR_ARCHITEW6432 } else { $env:PROCESSOR_ARCHITECTURE }
    $isArm64 = $arch -eq 'ARM64'
    try {
        $cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
        if ($cpu) { $isArm64 = $cpu.Architecture -eq 12 }
    }
    catch { }   # no CIM (e.g. WMI locked down): the environment variables decide
    $rid = if ($isArm64) { 'win-arm64' } else { 'win-x64' }

    # Set up front: an unset $tempDir in the finally below would be looked up in the CALLER's scope.
    $tempDir = $null
    try {
        if ($FromDirectory) {
            # The folder's zips for this architecture, each with the version its name carries.
            $found = @(foreach ($zip in Get-ChildItem -LiteralPath $FromDirectory -Filter "bc-word-layout-mcp-*-$rid.zip" -File) {
                if ($zip.Name -match "^bc-word-layout-mcp-(.+)-$rid\.zip$") {
                    [pscustomobject]@{ Version = $Matches[1]; SemVer = ConvertTo-SemVer $Matches[1]; Zip = $zip }
                }
            })
            # The SemVer ones in ascending order (a folder holds a handful, so an insertion sort will do).
            $ordered = New-Object System.Collections.Generic.List[object]
            foreach ($candidate in $found | Where-Object SemVer) {
                $at = 0
                while ($at -lt $ordered.Count -and (Compare-SemVer $ordered[$at].SemVer $candidate.SemVer) -le 0) { $at++ }
                $ordered.Insert($at, $candidate)
            }

            if ($Version -eq 'latest') {
                if (-not $found) { throw "No bc-word-layout-mcp-*-$rid.zip in $FromDirectory" }
                if (-not $ordered.Count) {
                    throw "No bc-word-layout-mcp-*-$rid.zip in $FromDirectory has a SemVer version to pick the latest from (found: $(($found | ForEach-Object Version) -join ', ')). Pass -Version to install one of them by name."
                }
                $pick = $ordered[$ordered.Count - 1]
            }
            else {
                $wanted = $Version.TrimStart('v')
                $pick = $found | Where-Object { $_.Version -eq $wanted } | Select-Object -First 1
                if (-not $pick) {
                    $present = @($ordered | ForEach-Object Version) + @($found | Where-Object { -not $_.SemVer } | ForEach-Object Version)
                    throw "No bc-word-layout-mcp-$wanted-$rid.zip in $FromDirectory. Versions there for ${rid}: $(if ($present) { $present -join ', ' } else { 'none' })."
                }
            }
            $resolvedVersion = $pick.Version
            $zipPath = $pick.Zip.FullName
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
    }
    finally {
        # Never leave the ~50 MB download behind, whether the install succeeded or threw.
        if ($tempDir -and (Test-Path -LiteralPath $tempDir)) {
            try { Remove-Item -LiteralPath $tempDir -Recurse -Force }
            catch { Write-Warning "Could not remove the download folder $tempDir ($($_.Exception.Message)); it is safe to delete." }
        }
    }
} @args
