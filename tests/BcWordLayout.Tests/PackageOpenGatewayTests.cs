using System.Text.RegularExpressions;
using BcWordLayout.Domain;

namespace BcWordLayout.Tests;

/// <summary>
/// Keeps <see cref="LayoutPackage.Open"/> the ONLY place production code opens a <c>.docx</c>: a direct
/// <c>WordprocessingDocument.Open</c> would skip <see cref="ResourceLimits.EnsurePackageWithinLimits"/> and
/// hand an untrusted package straight to the SDK's recursive part loads — the uncatchable stack overflow
/// the gate exists to prevent. Mechanical rather than a review convention (SOLUTION-DESIGN §6, principle 6).
/// </summary>
public class PackageOpenGatewayTests
{
    private static readonly Regex DirectOpen = new(@"\bWordprocessingDocument\s*\.\s*Open\s*\(", RegexOptions.Compiled);

    [Fact]
    public void No_production_source_opens_a_package_except_through_LayoutPackage()
    {
        var srcDir = Path.Combine(RepoRoot(), "src");
        var gateway = Path.Combine(srcDir, "BcWordLayout.Domain", "LayoutPackage.cs");

        var offenders = Directory.EnumerateFiles(srcDir, "*.cs", SearchOption.AllDirectories)
            .Where(f => !IsBuildOutput(f) && !string.Equals(f, gateway, StringComparison.OrdinalIgnoreCase))
            .SelectMany(f => File.ReadLines(f)
                .Select((line, i) => (File: f, Line: i + 1, Text: line))
                .Where(l => !l.Text.TrimStart().StartsWith("//", StringComparison.Ordinal) && DirectOpen.IsMatch(l.Text)))
            .Select(l => $"{Path.GetRelativePath(srcDir, l.File)}:{l.Line}")
            .ToList();

        Assert.True(offenders.Count == 0,
            "Open packages through LayoutPackage.Open so the resource gate runs first: " + string.Join(", ", offenders));
    }

    private static bool IsBuildOutput(string path) =>
        path.Contains($"{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}", StringComparison.Ordinal)
        || path.Contains($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}", StringComparison.Ordinal);

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
