using System.Text.Json;
using System.Text.RegularExpressions;

namespace BcWordLayout.Tests;

/// <summary>
/// GitHub issue #108: <c>THIRD-PARTY-NOTICES.md</c> is the notice the Apache-2.0 and BSD components require
/// to travel with the binaries, and its version table had drifted from every Dependabot bump since 1.0.0. It
/// now names components without versions, and these tests keep it honest against what the server actually
/// ships: the package closure of <c>BcWordLayout.McpHost.deps.json</c> (copied next to the test assembly,
/// so the check needs no repository checkout). A version bump never fails them; a new shipped dependency
/// without a row does.
/// </summary>
public class ThirdPartyNoticesTests
{
    private static string PackagingFile(string name) => Path.Combine(AppContext.BaseDirectory, "packaging", name);

    /// <summary>The first cell of every row of the notices file's runtime-dependencies table.</summary>
    private static List<string> RuntimeComponents()
    {
        var text = File.ReadAllText(PackagingFile("THIRD-PARTY-NOTICES.md"));
        var start = text.IndexOf("## Runtime dependencies", StringComparison.Ordinal);
        Assert.True(start >= 0, "THIRD-PARTY-NOTICES.md has no '## Runtime dependencies' section.");
        var end = text.IndexOf("\n## ", start + 1, StringComparison.Ordinal);
        var section = end < 0 ? text[start..] : text[start..end];

        return section.Split('\n')
            .Select(line => line.Trim())
            .Where(line => line.StartsWith('|') && !line.StartsWith("|---", StringComparison.Ordinal))
            .Select(line => line.Split('|', StringSplitOptions.TrimEntries)[1])
            .Where(cell => cell != "Component")
            .ToList();
    }

    /// <summary>
    /// The packages the server ships, by id: everything reachable from the host project in its deps.json
    /// that carries a managed, native or Windows-RID asset (the closure is RID-less, so it also lists the
    /// Linux/macOS native packages the win-x64/win-arm64 builds never include). Plus, separately, the
    /// packages the shipped projects reference directly.
    /// </summary>
    private static (HashSet<string> Shipped, HashSet<string> Direct) HostPackages()
    {
        using var deps = JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "BcWordLayout.McpHost.deps.json")));
        var root = deps.RootElement;
        var target = root.GetProperty("targets").EnumerateObject().First().Value;
        var libraries = root.GetProperty("libraries");
        string TypeOf(string key) => libraries.TryGetProperty(key, out var lib) ? lib.GetProperty("type").GetString()! : "";

        var host = target.EnumerateObject().First(e => e.Name.StartsWith("BcWordLayout.McpHost/", StringComparison.Ordinal)).Name;
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var direct = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var pending = new Stack<string>([host]);
        while (pending.TryPop(out var key))
        {
            if (!seen.Add(key) || !target.TryGetProperty(key, out var entry)
                || !entry.TryGetProperty("dependencies", out var dependencies))
            {
                continue;
            }

            foreach (var dependency in dependencies.EnumerateObject())
            {
                var dependencyKey = $"{dependency.Name}/{dependency.Value.GetString()}";
                if (TypeOf(key) == "project" && TypeOf(dependencyKey) == "package")
                {
                    direct.Add(dependency.Name);
                }

                pending.Push(dependencyKey);
            }
        }

        var shipped = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var key in seen.Where(k => TypeOf(k) == "package"))
        {
            var entry = target.GetProperty(key);
            var windowsOrAnyRid = entry.TryGetProperty("runtime", out _) || entry.TryGetProperty("native", out _)
                || (entry.TryGetProperty("runtimeTargets", out var runtimeTargets)
                    && runtimeTargets.EnumerateObject().Any(t => t.Value.GetProperty("rid").GetString()!.StartsWith("win", StringComparison.Ordinal)));
            if (windowsOrAnyRid)
            {
                shipped.Add(key[..key.IndexOf('/', StringComparison.Ordinal)]);
            }
        }

        return (shipped, direct);
    }

    /// <summary>
    /// True when a row names <paramref name="packageId"/>: the row's leading package id is the package itself,
    /// or the family it belongs to (<c>SkiaSharp</c> covers <c>SkiaSharp.NativeAssets.Win32</c>, as the row
    /// "SkiaSharp (+ <c>SkiaSharp.NativeAssets.*</c>)" says).
    /// </summary>
    private static bool Names(string component, string packageId)
    {
        var leading = component.Split(' ', 2)[0];
        return string.Equals(leading, packageId, StringComparison.OrdinalIgnoreCase)
            || packageId.StartsWith(leading + ".", StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Every_package_the_server_ships_has_a_row_in_the_runtime_table()
    {
        var components = RuntimeComponents();
        var (shipped, direct) = HostPackages();
        Assert.NotEmpty(shipped);

        // The .NET runtime's own packages (Microsoft.*, System.*) are covered by the rows for the .NET
        // libraries that bring them; anything the shipped projects reference directly must be named.
        var mustBeNamed = shipped
            .Where(p => direct.Contains(p) || !(p.StartsWith("Microsoft.", StringComparison.OrdinalIgnoreCase)
                                               || p.StartsWith("System.", StringComparison.OrdinalIgnoreCase)))
            .OrderBy(p => p, StringComparer.Ordinal)
            .ToList();
        var missing = mustBeNamed.Where(p => !components.Any(c => Names(c, p))).ToList();

        Assert.True(
            missing.Count == 0,
            "THIRD-PARTY-NOTICES.md's runtime table has no row for these shipped packages (add the component, "
            + "its license and its project link): " + string.Join(", ", missing));
    }

    [Fact]
    public void Every_row_of_the_runtime_table_is_still_shipped()
    {
        var (shipped, _) = HostPackages();
        var stale = RuntimeComponents()
            .Where(c => !c.StartsWith(".NET runtime", StringComparison.Ordinal))
            .Where(c => !shipped.Any(p => Names(c, p)))
            .ToList();

        Assert.True(stale.Count == 0, "THIRD-PARTY-NOTICES.md names components the server no longer ships: " + string.Join(", ", stale));
    }

    [Fact]
    public void The_notices_name_no_versions()
    {
        // Versions drifted from every Dependabot bump (a 1.4.1 row next to a shipped 2.2.0); the project files
        // are the one place that pins them.
        var tableRows = File.ReadAllLines(PackagingFile("THIRD-PARTY-NOTICES.md")).Where(l => l.TrimStart().StartsWith('|'));
        var versioned = tableRows.Where(l => Regex.IsMatch(l, @"\b\d+\.\d+(\.\d+|\.x)\b")).ToList();

        Assert.True(versioned.Count == 0, "THIRD-PARTY-NOTICES.md tables should name no versions: " + string.Join(" | ", versioned));
    }
}
