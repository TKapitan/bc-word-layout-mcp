using System.Runtime.CompilerServices;
using System.Text;

namespace BcWordLayout.Tests;

/// <summary>
/// Approval-testing ("snapshot") helper: compares freshly produced text against a checked-in approved
/// file, failing with an actionable diff when they differ. The approved files live under
/// <c>Snapshots/</c> next to this source file and are resolved via <see cref="CallerFilePathAttribute"/>
/// against THIS file's own compile-time path rather than <see cref="AppContext.BaseDirectory"/> — the
/// helper must read (and, when regenerating, write) the real source-tree files, not a copy-to-output
/// build artifact, so a snapshot update actually lands somewhere `git status` will see it.
/// </summary>
public static class SnapshotAssert
{
    private static readonly UTF8Encoding Utf8NoBom = new(encoderShouldEmitUTF8Identifier: false);

    private static readonly string SnapshotsDirectory = ResolveSnapshotsDirectory();

    /// <summary>
    /// Resolves <c>Snapshots/</c> relative to this file's own directory. The <see cref="CallerFilePathAttribute"/>
    /// default is filled in by the compiler at THIS call site (the static field initializer above, which
    /// lives in SnapshotAssert.cs itself) — not at each test's call site — so it reliably yields
    /// <c>tests/BcWordLayout.Tests/Snapshots</c> regardless of which test class calls <see cref="Match"/>.
    /// </summary>
    private static string ResolveSnapshotsDirectory([CallerFilePath] string thisFilePath = "") =>
        Path.Combine(Path.GetDirectoryName(thisFilePath)!, "Snapshots");

    /// <summary>
    /// Compares <paramref name="actual"/> against the approved file <c>Snapshots/&lt;snapshotName&gt;</c>,
    /// normalizing both sides' line endings to <c>\n</c> first so the comparison (and the stored
    /// snapshots) are portable across machines and git's <c>autocrlf</c>.
    /// </summary>
    /// <remarks>
    /// When the approved file does not exist yet, or the <c>UPDATE_SNAPSHOTS</c> environment variable is
    /// set to a truthy value (<c>1</c>/<c>true</c>/<c>yes</c>/<c>on</c>), this (re)creates the approved file from <paramref name="actual"/> and
    /// passes — that is how snapshots are generated and deliberately updated. Otherwise, a mismatch writes
    /// a sibling <c>&lt;snapshotName&gt;.received</c> file for inspection and fails with the 1-based line
    /// number and content of the first differing line. A match deletes any stale <c>.received</c> sibling
    /// left over from a previous failing run.
    /// </remarks>
    public static void Match(string snapshotName, string actual)
    {
        Directory.CreateDirectory(SnapshotsDirectory);

        var normalizedActual = Normalize(actual);
        var approvedPath = Path.Combine(SnapshotsDirectory, snapshotName);
        var receivedPath = approvedPath + ".received";

        if (!File.Exists(approvedPath) || HasUpdateSnapshotsFlag())
        {
            File.WriteAllText(approvedPath, normalizedActual, Utf8NoBom);
            DeleteIfExists(receivedPath);
            return;
        }

        var approved = Normalize(File.ReadAllText(approvedPath, Encoding.UTF8));
        if (string.Equals(approved, normalizedActual, StringComparison.Ordinal))
        {
            DeleteIfExists(receivedPath);
            return;
        }

        File.WriteAllText(receivedPath, normalizedActual, Utf8NoBom);

        var (lineNumber, approvedLine, actualLine) = FirstDifferingLine(approved, normalizedActual);
        Assert.Fail(
            $"Snapshot mismatch for '{snapshotName}'.{Environment.NewLine}"
            + $"Approved: {approvedPath}{Environment.NewLine}"
            + $"Received: {receivedPath}{Environment.NewLine}"
            + $"First differing line {lineNumber}:{Environment.NewLine}"
            + $"  approved: {approvedLine}{Environment.NewLine}"
            + $"  actual:   {actualLine}{Environment.NewLine}"
            + "If this change is expected, set UPDATE_SNAPSHOTS=1 and re-run to accept it.");
    }

    // Only genuinely truthy values enable regeneration, so that UPDATE_SNAPSHOTS=0 / false / no
    // (a developer trying to DISABLE it) does not silently rewrite every snapshot and pass.
    private static bool HasUpdateSnapshotsFlag()
    {
        var value = Environment.GetEnvironmentVariable("UPDATE_SNAPSHOTS")?.Trim();
        return value is "1"
            || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "yes", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "on", StringComparison.OrdinalIgnoreCase);
    }

    private static string Normalize(string text) => text.Replace("\r\n", "\n").Replace("\r", "\n");

    private static void DeleteIfExists(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }

    /// <summary>
    /// Returns the 1-based line number of the first line at which <paramref name="approved"/> and
    /// <paramref name="actual"/> differ, plus each side's content at that line — an out-of-range side
    /// (one text simply has fewer lines) reports a placeholder instead of throwing.
    /// </summary>
    private static (int LineNumber, string Approved, string Actual) FirstDifferingLine(string approved, string actual)
    {
        var approvedLines = approved.Split('\n');
        var actualLines = actual.Split('\n');
        var lineCount = Math.Max(approvedLines.Length, actualLines.Length);

        for (var i = 0; i < lineCount; i++)
        {
            var approvedLine = i < approvedLines.Length ? approvedLines[i] : "<approved ended>";
            var actualLine = i < actualLines.Length ? actualLines[i] : "<actual ended>";
            if (!string.Equals(approvedLine, actualLine, StringComparison.Ordinal))
            {
                return (i + 1, approvedLine, actualLine);
            }
        }

        // Unreachable in practice: callers only invoke this after confirming the full strings differ.
        return (lineCount, "<identical>", "<identical>");
    }
}
