using System.Diagnostics;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace BcWordLayout.Tests;

/// <summary>
/// GitHub issue #109: <c>scripts/install.ps1</c>, the GitHub Release installer, run in real child shells -
/// Windows PowerShell 5.1 (where the README's <c>irm … | iex</c> one-liner usually runs) and pwsh when it is on
/// PATH:
/// <list type="bullet">
/// <item>Under <c>iex</c> the script runs in a scope of its own: the caller's <c>$Version</c> and preference
/// variables are untouched, and none of its variables or functions stay behind.</item>
/// <item>The temp download folder is removed whether the install throws or succeeds.</item>
/// <item><c>-FromDirectory</c> installs exactly the <c>-Version</c> asked for, or else the highest SemVer
/// version (not the first file name), and a missing version names the ones that are there.</item>
/// <item>The zip matches the machine's NATIVE architecture, also from x64 PowerShell emulated on ARM64 (the
/// decision is driven by stand-ins, so no ARM64 machine is needed).</item>
/// </list>
/// Nothing touches the network or the real install: caller-scope functions stand in for
/// <c>Invoke-RestMethod</c>/<c>Invoke-WebRequest</c> (a function wins over a cmdlet in command lookup) and,
/// except in the one file run, for <c>Get-CimInstance</c>; every install goes under this class's temp folder,
/// and every child shell gets a TEMP of its own. The drivers run under <c>Set-StrictMode -Version Latest</c>,
/// which a user's profile may turn on and iex'd code inherits; it also fails any read of a variable the
/// script never set, which would otherwise silently pick up the caller's variable of that name.
/// </summary>
public sealed class InstallScriptTests : IDisposable
{
    private static readonly string Installer = Path.Combine(RepoRoot(), "scripts", "install.ps1");

    // Ascending SemVer 2.0 precedence. Neighbours pin the comparer's rules: a release above its pre-releases,
    // a numeric identifier below an alphanumeric one (1.2.0-1 < 1.2.0-rc), the shorter list first when the
    // shared identifiers tie (rc < rc.2), numeric identifiers numerically (rc.2 < rc.10), and 1.10.0 above
    // 1.2.0, which any ordering by file name gets wrong.
    private static readonly string[] OrderedVersions =
        ["1.0.0", "1.1.0-rc.1", "1.1.0", "1.2.0-1", "1.2.0-rc", "1.2.0-rc.2", "1.2.0-rc.10", "1.10.0-rc.1", "1.10.0"];

    private readonly string workDir = Directory.CreateTempSubdirectory("bcwl-installer-tests-").FullName;

    public void Dispose()
    {
        try
        {
            // Unlink every 'current' junction first, so the recursive delete never walks through one.
            foreach (var current in Directory.EnumerateDirectories(workDir, "current", SearchOption.AllDirectories)
                         .Where(path => new DirectoryInfo(path).LinkTarget is not null).ToList())
            {
                Directory.Delete(current);
            }

            Directory.Delete(workDir, recursive: true);
        }
        catch (IOException)
        {
            // Best-effort temp cleanup (a virus scanner may still hold a just-unpacked file).
        }
    }

    [WindowsOnlyFact]
    public void Under_iex_the_script_leaves_the_callers_session_as_it_was_and_removes_its_download_folder()
    {
        // The stand-ins let the run get as far as the download, so the temp folder exists, and fail there.
        // iex cannot pass -InstallRoot, so the default it derives from LOCALAPPDATA is pointed here as well.
        const string driver = """
            $env:LOCALAPPDATA = $env:BCWL_TEST_INSTALLS
            $ErrorActionPreference = 'Continue'
            $ProgressPreference = 'Continue'
            $Version = '9.9.9'
            $downloadDir = $null
            $failure = $null
            $before = $null
            function Invoke-RestMethod { param($Uri) [pscustomobject]@{ tag_name = 'v1.2.3' } }
            function Invoke-WebRequest { param($Uri, $OutFile) $global:downloadDir = Split-Path -Parent $OutFile; throw 'offline' }
            function Get-CimInstance { [pscustomobject]@{ Architecture = 9 } }
            function Get-SessionNames { @(Get-Variable | ForEach-Object Name) + @(Get-ChildItem function: | ForEach-Object Name) }
            $before = Get-SessionNames
            try { Invoke-Expression (Get-Content -Raw -LiteralPath $env:BCWL_TEST_INSTALLER) } catch { $failure = $_.Exception.Message }
            [ordered]@{
                Failure = $failure
                Version = $Version
                ErrorActionPreference = "$ErrorActionPreference"
                ProgressPreference = "$ProgressPreference"
                # $_ and $PSItem are the catch block's own
                LeftBehind = @(Get-SessionNames | Where-Object { $before -notcontains $_ -and '_', 'PSItem' -notcontains $_ })
                DownloadDir = $downloadDir
                DownloadDirExists = [bool]$downloadDir -and (Test-Path -LiteralPath $downloadDir)
            } | ConvertTo-Json | Set-Content -LiteralPath $env:BCWL_TEST_RESULT -Encoding UTF8
            """;

        foreach (var shell in Shells())
        {
            var name = Path.GetFileName(shell);
            var temp = ShellDirectory(shell, "temp");
            var installs = ShellDirectory(shell, "installs");

            var result = RunDriver(shell, driver, temp, new() { ["BCWL_TEST_INSTALLS"] = installs });

            Assert.True(Text(result, "Failure") == "offline", $"{name}: the run should have stopped at the download stand-in: {result}");
            Assert.True(Text(result, "Version") == "9.9.9", $"{name}: the caller's $Version changed: {result}");
            Assert.True(Text(result, "ErrorActionPreference") == "Continue", $"{name}: the caller's $ErrorActionPreference changed: {result}");
            Assert.True(Text(result, "ProgressPreference") == "Continue", $"{name}: the caller's $ProgressPreference changed: {result}");
            Assert.True(result.GetProperty("LeftBehind").GetArrayLength() == 0, $"{name}: variables or functions left in the caller's session: {result}");
            AssertDownloadFolderRemoved(result, temp, name);
            Assert.False(Directory.Exists(Path.Combine(installs, "bc-word-layout-mcp")), $"{name}: a failed download still installed something");
        }
    }

    [WindowsOnlyFact]
    public void A_successful_download_installs_the_latest_release_and_removes_its_download_folder()
    {
        CreateReleaseFolder("releases", "1.10.0");
        const string driver = """
            $requested = New-Object System.Collections.Generic.List[string]
            $downloadDir = $null
            $failure = $null
            function Invoke-RestMethod { param($Uri) $requested.Add($Uri); [pscustomobject]@{ tag_name = 'v1.10.0' } }
            function Invoke-WebRequest {
                param($Uri, $OutFile)
                $requested.Add($Uri)
                $global:downloadDir = Split-Path -Parent $OutFile
                Copy-Item -LiteralPath (Join-Path $env:BCWL_TEST_FIXTURES "releases\$(($Uri -split '/')[-1])") -Destination $OutFile
            }
            function Get-CimInstance { [pscustomobject]@{ Architecture = 9 } }
            try { & ([scriptblock]::Create((Get-Content -Raw -LiteralPath $env:BCWL_TEST_INSTALLER))) -InstallRoot $env:BCWL_TEST_INSTALLS }
            catch { $failure = $_.Exception.Message }
            [ordered]@{
                Failure = $failure
                Requested = @($requested)
                DownloadDir = $downloadDir
                DownloadDirExists = [bool]$downloadDir -and (Test-Path -LiteralPath $downloadDir)
            } | ConvertTo-Json | Set-Content -LiteralPath $env:BCWL_TEST_RESULT -Encoding UTF8
            """;

        foreach (var shell in Shells())
        {
            var name = Path.GetFileName(shell);
            var temp = ShellDirectory(shell, "temp");
            var installRoot = ShellDirectory(shell, "installs");

            var result = RunDriver(shell, driver, temp, new() { ["BCWL_TEST_INSTALLS"] = installRoot });

            Assert.True(result.GetProperty("Failure").ValueKind == JsonValueKind.Null, $"{name}: the install failed: {result}");
            Assert.Equal(
                new[]
                {
                    "https://api.github.com/repos/TKapitan/bc-word-layout-mcp/releases/latest",
                    "https://github.com/TKapitan/bc-word-layout-mcp/releases/download/v1.10.0/bc-word-layout-mcp-1.10.0-win-x64.zip",
                    "https://github.com/TKapitan/bc-word-layout-mcp/releases/download/v1.10.0/SHA256SUMS.txt",
                },
                result.GetProperty("Requested").EnumerateArray().Select(uri => uri.GetString()!).ToArray());
            AssertInstalled(installRoot, "1.10.0", "win-x64", name);
            AssertDownloadFolderRemoved(result, temp, name);
        }
    }

    [WindowsOnlyFact]
    public void FromDirectory_installs_exactly_the_requested_version_or_else_the_highest_SemVer_one()
    {
        CreateReleaseFolder("releases", [.. OrderedVersions, "nightly"]);
        CreateReleaseFolder("nightly-only", "nightly");
        const string driver = """
            function Get-CimInstance { [pscustomobject]@{ Architecture = 9 } }
            $installer = [scriptblock]::Create((Get-Content -Raw -LiteralPath $env:BCWL_TEST_INSTALLER))
            $results = [ordered]@{}
            function Install([string]$Case, [string]$Folder, [hashtable]$Arguments) {
                try {
                    & $installer -FromDirectory (Join-Path $env:BCWL_TEST_FIXTURES $Folder) -InstallRoot (Join-Path $env:BCWL_TEST_INSTALLS $Case) @Arguments
                    $results[$Case] = 'installed'
                }
                catch { $results[$Case] = $_.Exception.Message }
            }
            Install 'latest' 'releases' @{}
            Install 'prerelease' 'releases' @{ Version = '1.1.0-rc.1' }
            Install 'v-prefixed' 'releases' @{ Version = 'v1.0.0' }
            Install 'not-semver' 'releases' @{ Version = 'nightly' }
            Install 'missing' 'releases' @{ Version = '9.9.9' }
            Install 'latest-without-semver' 'nightly-only' @{}
            $results | ConvertTo-Json | Set-Content -LiteralPath $env:BCWL_TEST_RESULT -Encoding UTF8
            """;

        foreach (var shell in Shells())
        {
            var name = Path.GetFileName(shell);
            var installs = ShellDirectory(shell, "installs");

            var result = RunDriver(shell, driver, ShellDirectory(shell, "temp"), new() { ["BCWL_TEST_INSTALLS"] = installs });

            AssertCaseInstalled(result, installs, "latest", "1.10.0", "win-x64", name);
            AssertCaseInstalled(result, installs, "prerelease", "1.1.0-rc.1", "win-x64", name);
            AssertCaseInstalled(result, installs, "v-prefixed", "1.0.0", "win-x64", name);
            // Asked for by name, a zip installs even when its version is not SemVer.
            AssertCaseInstalled(result, installs, "not-semver", "nightly", "win-x64", name);

            var missing = Text(result, "missing");
            Assert.True(
                missing.StartsWith("No bc-word-layout-mcp-9.9.9-win-x64.zip in ", StringComparison.Ordinal)
                && missing.EndsWith($". Versions there for win-x64: {string.Join(", ", OrderedVersions)}, nightly.", StringComparison.Ordinal),
                $"{name}: a missing -Version should list the versions present, in SemVer order: {missing}");
            Assert.False(Directory.Exists(Path.Combine(installs, "missing")), $"{name}: a missing -Version still installed something");

            var unpickable = Text(result, "latest-without-semver");
            Assert.True(
                unpickable.Contains("SemVer", StringComparison.Ordinal) && unpickable.Contains("(found: nightly)", StringComparison.Ordinal)
                && unpickable.Contains("-Version", StringComparison.Ordinal),
                $"{name}: with no SemVer zip, the error should say so and point at -Version: {unpickable}");
            Assert.False(Directory.Exists(Path.Combine(installs, "latest-without-semver")), $"{name}: an unpickable folder still installed something");
        }
    }

    [WindowsOnlyFact]
    public void The_zip_matches_the_native_architecture_even_from_emulated_x64_PowerShell()
    {
        CreateReleaseFolder("releases", "1.0.0");
        // $cimArchitecture is what the Get-CimInstance stand-in reports as Win32_Processor.Architecture
        // (9 = x64, 12 = ARM64); $null makes it fail the way a locked-down WMI does.
        const string driver = """
            $installer = [scriptblock]::Create((Get-Content -Raw -LiteralPath $env:BCWL_TEST_INSTALLER))
            $cimArchitecture = $null
            function Get-CimInstance {
                if ($null -eq $global:cimArchitecture) { throw 'WMI is unavailable' }
                [pscustomobject]@{ Architecture = $global:cimArchitecture }
            }
            $results = [ordered]@{}
            function Install([string]$Case, $Cim, [string]$Wow64Native, [string]$Process) {
                $global:cimArchitecture = $Cim
                $env:PROCESSOR_ARCHITEW6432 = $Wow64Native
                $env:PROCESSOR_ARCHITECTURE = $Process
                try {
                    & $installer -FromDirectory (Join-Path $env:BCWL_TEST_FIXTURES 'releases') -InstallRoot (Join-Path $env:BCWL_TEST_INSTALLS $Case) -Version 1.0.0
                    $results[$Case] = 'installed'
                }
                catch { $results[$Case] = $_.Exception.Message }
            }
            Install 'emulated-x64' 12 '' 'AMD64'
            Install 'wow64-without-cim' $null 'ARM64' 'x86'
            Install 'arm64-without-cim' $null '' 'ARM64'
            Install 'x64-without-cim' $null '' 'AMD64'
            $results | ConvertTo-Json | Set-Content -LiteralPath $env:BCWL_TEST_RESULT -Encoding UTF8
            """;

        foreach (var shell in Shells())
        {
            var name = Path.GetFileName(shell);
            var installs = ShellDirectory(shell, "installs");

            var result = RunDriver(shell, driver, ShellDirectory(shell, "temp"), new() { ["BCWL_TEST_INSTALLS"] = installs });

            // x64 PowerShell emulated on ARM64: PROCESSOR_ARCHITECTURE says AMD64, but Win32_Processor says ARM64.
            AssertCaseInstalled(result, installs, "emulated-x64", "1.0.0", "win-arm64", name);
            // Without CIM, x86 PowerShell on ARM64 (WOW64) finds the native architecture in PROCESSOR_ARCHITEW6432.
            AssertCaseInstalled(result, installs, "wow64-without-cim", "1.0.0", "win-arm64", name);
            AssertCaseInstalled(result, installs, "arm64-without-cim", "1.0.0", "win-arm64", name);
            AssertCaseInstalled(result, installs, "x64-without-cim", "1.0.0", "win-x64", name);
        }
    }

    [WindowsOnlyFact]
    public void A_file_run_takes_its_parameters_by_name_and_detects_this_machines_architecture()
    {
        // The one run without stand-ins: parameters reach the scriptblock through @args, and the real
        // Win32_Processor query picks the zip.
        var releases = CreateReleaseFolder("releases", OrderedVersions);
        var nativeRid = RuntimeInformation.OSArchitecture == Architecture.Arm64 ? "win-arm64" : "win-x64";

        foreach (var shell in Shells())
        {
            var name = Path.GetFileName(shell);
            var installRoot = ShellDirectory(shell, "installs");

            var run = Run(shell, ShellDirectory(shell, "temp"), new Dictionary<string, string>(),
                "-File", Installer, "-FromDirectory", releases, "-InstallRoot", installRoot, "-Version", "1.2.0-rc.2");

            Assert.True(run.ExitCode == 0, $"{name}: exit code {run.ExitCode}:\n{run.Output}");
            AssertInstalled(installRoot, "1.2.0-rc.2", nativeRid, name);
        }
    }

    /// <summary>
    /// A folder of stand-in release assets: for every version, one zip per RID holding a placeholder
    /// <c>BcWordLayout.McpHost.exe</c> (what the script checks for after unpacking) and a <c>release.txt</c>
    /// naming the version and RID, plus the <c>SHA256SUMS.txt</c> the script verifies them against.
    /// </summary>
    private string CreateReleaseFolder(string folderName, params string[] versions)
    {
        var folder = Directory.CreateDirectory(Path.Combine(workDir, folderName)).FullName;
        var sums = new List<string>();
        foreach (var version in versions)
        {
            foreach (var rid in (string[])["win-x64", "win-arm64"])
            {
                var zipName = $"bc-word-layout-mcp-{version}-{rid}.zip";
                var zipPath = Path.Combine(folder, zipName);
                using (var zip = ZipFile.Open(zipPath, ZipArchiveMode.Create))
                {
                    AddEntry(zip, "BcWordLayout.McpHost.exe", $"placeholder for {version} {rid}");
                    AddEntry(zip, "release.txt", $"{version} {rid}");
                }

                sums.Add($"{Convert.ToHexStringLower(SHA256.HashData(File.ReadAllBytes(zipPath)))}  {zipName}");
            }
        }

        File.WriteAllText(Path.Combine(folder, "SHA256SUMS.txt"), string.Join('\n', sums));
        return folder;
    }

    private static void AddEntry(ZipArchive zip, string entryName, string content)
    {
        using var writer = new StreamWriter(zip.CreateEntry(entryName).Open());
        writer.Write(content);
    }

    private string ShellDirectory(string shell, string purpose) =>
        Directory.CreateDirectory(Path.Combine(workDir, Path.GetFileNameWithoutExtension(shell), purpose)).FullName;

    /// <summary>A driver's <paramref name="caseName"/> run reported success and installed <paramref name="version"/>.</summary>
    private static void AssertCaseInstalled(JsonElement results, string installs, string caseName, string version, string rid, string shellName)
    {
        var outcome = Text(results, caseName);
        Assert.True(outcome == "installed", $"{shellName} {caseName}: {outcome}");
        AssertInstalled(Path.Combine(installs, caseName), version, rid, $"{shellName} {caseName}");
    }

    /// <summary>
    /// <paramref name="installRoot"/> holds <paramref name="version"/> unpacked from its <paramref name="rid"/>
    /// zip and nothing else, and its <c>current</c> junction leads there.
    /// </summary>
    private static void AssertInstalled(string installRoot, string version, string rid, string context)
    {
        var current = new DirectoryInfo(Path.Combine(installRoot, "current"));
        Assert.True(current.Exists && current.LinkTarget is not null, $"{context}: no 'current' junction in {installRoot}");
        Assert.True(
            Path.GetFileName(Path.TrimEndingDirectorySeparator(current.LinkTarget)) == version,
            $"{context}: 'current' points at {current.LinkTarget}, expected version {version}");

        var marker = File.ReadAllText(Path.Combine(current.FullName, "release.txt"));
        Assert.True(marker == $"{version} {rid}", $"{context}: installed '{marker}', expected '{version} {rid}'");

        var entries = Directory.GetFileSystemEntries(installRoot).Select(entry => Path.GetFileName(entry)).Order(StringComparer.Ordinal);
        Assert.Equal(new[] { "current", version }.Order(StringComparer.Ordinal), entries);
    }

    /// <summary>The run's download folder was a <c>bcwl-install-*</c> folder in the child's TEMP and is gone.</summary>
    private static void AssertDownloadFolderRemoved(JsonElement result, string temp, string context)
    {
        var downloadDir = Text(result, "DownloadDir");
        Assert.True(
            Path.GetFileName(downloadDir).StartsWith("bcwl-install-", StringComparison.Ordinal)
            && string.Equals(Path.GetDirectoryName(downloadDir), temp, StringComparison.OrdinalIgnoreCase),
            $"{context}: expected the download in a bcwl-install-* folder in {temp}: {result}");
        Assert.False(result.GetProperty("DownloadDirExists").GetBoolean(), $"{context}: the download folder outlived the run: {result}");
        Assert.Empty(Directory.EnumerateFileSystemEntries(temp, "bcwl-install-*"));
    }

    private static string Text(JsonElement result, string property) =>
        result.GetProperty(property).GetString() ?? throw new InvalidOperationException($"{property} is null in {result}");

    /// <summary>
    /// Runs <paramref name="driver"/> under strict mode through <c>-EncodedCommand</c>, so it runs at the global
    /// scope like a prompt, and returns the JSON it writes to <c>$env:BCWL_TEST_RESULT</c>. Paths reach it as
    /// environment variables, never quoted into its code.
    /// </summary>
    private JsonElement RunDriver(string shell, string driver, string temp, Dictionary<string, string> environment)
    {
        var resultPath = Path.Combine(workDir, $"result-{Guid.NewGuid():N}.json");
        environment["BCWL_TEST_RESULT"] = resultPath;
        environment["BCWL_TEST_INSTALLER"] = Installer;
        environment["BCWL_TEST_FIXTURES"] = workDir;

        var command = "Set-StrictMode -Version Latest\n" + driver;
        var run = Run(shell, temp, environment, "-EncodedCommand", Convert.ToBase64String(Encoding.Unicode.GetBytes(command)));

        Assert.True(File.Exists(resultPath), $"{Path.GetFileName(shell)} wrote no result (exit code {run.ExitCode}):\n{run.Output}");
        using var document = JsonDocument.Parse(File.ReadAllText(resultPath));
        return document.RootElement.Clone();
    }

    private sealed record ShellRun(int ExitCode, string Output);

    /// <summary>
    /// Runs <paramref name="shell"/> with <paramref name="arguments"/>, TEMP/TMP set to <paramref name="temp"/>
    /// and <paramref name="environment"/> added.
    /// </summary>
    private static ShellRun Run(string shell, string temp, IReadOnlyDictionary<string, string> environment, params string[] arguments)
    {
        var startInfo = new ProcessStartInfo(shell)
        {
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
        };
        foreach (var argument in (string[])["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", .. arguments])
        {
            startInfo.ArgumentList.Add(argument);
        }

        startInfo.Environment["TEMP"] = temp;
        startInfo.Environment["TMP"] = temp;
        // A PSModulePath inherited from a pwsh parent would point Windows PowerShell at pwsh's own modules.
        startInfo.Environment.Remove("PSModulePath");
        foreach (var (name, value) in environment)
        {
            startInfo.Environment[name] = value;
        }

        var output = new StringBuilder();
        using var process = new Process { StartInfo = startInfo };
        process.OutputDataReceived += (_, e) => Append(e.Data);
        process.ErrorDataReceived += (_, e) => Append(e.Data);
        process.Start();
        process.BeginOutputReadLine();
        process.BeginErrorReadLine();
        if (!process.WaitForExit(TimeSpan.FromMinutes(2)))
        {
            process.Kill(entireProcessTree: true);
            throw new TimeoutException($"{shell} did not finish within two minutes.");
        }

        process.WaitForExit(); // let the redirected streams drain
        lock (output)
        {
            return new ShellRun(process.ExitCode, output.ToString());
        }

        void Append(string? line)
        {
            if (line is not null)
            {
                lock (output)
                {
                    output.AppendLine(line);
                }
            }
        }
    }

    /// <summary>Windows PowerShell 5.1, plus pwsh when it is on PATH (left out silently when it is not).</summary>
    private static List<string> Shells()
    {
        var shells = new List<string> { Path.Combine(Environment.SystemDirectory, "WindowsPowerShell", "v1.0", "powershell.exe") };
        var pwsh = (Environment.GetEnvironmentVariable("PATH") ?? "")
            .Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(dir => Path.Combine(dir.Trim('"'), "pwsh.exe"))
            .FirstOrDefault(File.Exists);
        if (pwsh is not null)
        {
            shells.Add(pwsh);
        }

        return shells;
    }

    private static string RepoRoot()
    {
        for (var dir = new DirectoryInfo(AppContext.BaseDirectory); dir is not null; dir = dir.Parent)
        {
            if (File.Exists(Path.Combine(dir.FullName, "bc-word-layout-mcp.sln")))
            {
                return dir.FullName;
            }
        }

        throw new DirectoryNotFoundException("Could not find the repository root above the test assembly.");
    }
}
