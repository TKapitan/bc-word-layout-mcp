using System.Diagnostics;

namespace BcWordLayout.Tests;

/// <summary>
/// Builds the other spellings of a path that GitHub issue #92 is about, for Windows-only tests: a directory
/// junction (any account can make one, unlike a symlink), a symlink to a file when the account may create one,
/// and the 8.3 short name when the volume generates them.
/// </summary>
internal static class FileSystemLinks
{
    /// <summary>Creates a junction at <paramref name="link"/> pointing at the existing folder <paramref name="target"/>.</summary>
    public static void CreateJunction(string link, string target)
    {
        var (exitCode, output) = RunCmd($"mklink /J \"{link}\" \"{target}\"");
        Assert.True(exitCode == 0, $"mklink /J failed with exit code {exitCode}: {output}");
    }

    /// <summary>Removes the junction itself, never the contents of the folder it points at.</summary>
    public static void DeleteJunction(string link)
    {
        if (Directory.Exists(link))
        {
            Directory.Delete(link);
        }
    }

    /// <summary>
    /// Creates a symlink at <paramref name="link"/> to the file <paramref name="target"/>, or returns false when
    /// the account may not create symlinks (Windows without Developer Mode or elevation).
    /// </summary>
    public static bool TryCreateFileSymlink(string link, string target)
    {
        try
        {
            File.CreateSymbolicLink(link, target);
            return true;
        }
        catch (IOException)
        {
            return false;
        }
        catch (UnauthorizedAccessException)
        {
            return false;
        }
    }

    /// <summary>
    /// The 8.3 short spelling of the existing <paramref name="path"/>, or null when it has none that differs
    /// (8.3 name generation is off for the volume, or every segment is already short).
    /// </summary>
    public static string? ShortPathOrNull(string path)
    {
        var (exitCode, output) = RunCmd($"for %I in (\"{path}\") do @echo %~sI");
        var shortPath = output.Trim();
        return exitCode == 0 && shortPath.Length > 0 && !string.Equals(shortPath, path, StringComparison.OrdinalIgnoreCase)
            ? shortPath
            : null;
    }

    private static (int ExitCode, string Output) RunCmd(string command)
    {
        // A raw argument string, not ArgumentList: cmd.exe does not understand the backslash-escaped quotes
        // ArgumentList would produce. The command never starts with a quote, so /c keeps its quotes as written.
        var startInfo = new ProcessStartInfo("cmd.exe", "/d /c " + command)
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
        };

        using var process = Process.Start(startInfo)!;
        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        Assert.True(process.WaitForExit(TimeSpan.FromSeconds(30)), $"cmd.exe did not finish: {command}");
        return (process.ExitCode, stdout.Result + stderr.Result);
    }
}
