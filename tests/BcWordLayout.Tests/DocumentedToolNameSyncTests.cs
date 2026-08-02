using System.Reflection;
using System.Text.RegularExpressions;
using BcWordLayout.McpHost.Tools;
using ModelContextProtocol.Server;

namespace BcWordLayout.Tests;

/// <summary>
/// The README's tool table and the skill's tool documentation both duplicate the live tool surface,
/// and both have drifted before (the README carried a wrong tool count until 2026-08-01). This is the
/// mechanical guard: the names documented in those files must match the <c>[McpServerTool]</c> names
/// exactly — a tool added, removed, or renamed without updating both docs fails the suite. The docs
/// are copied next to the test assembly at build time (see the csproj), keeping the tests hermetic.
/// </summary>
public class DocumentedToolNameSyncTests
{
    /// <summary>Table rows whose first cell is a single backticked tool name, e.g. <c>| `insert_field` | …</c>.</summary>
    private static readonly Regex FirstCellToolName =
        new(@"^\|\s*`([a-z][a-z0-9_]*)`\s*\|", RegexOptions.Compiled | RegexOptions.Multiline);

    private static string DocText(string fileName)
    {
        var full = Path.Combine(AppContext.BaseDirectory, "docsync", fileName);
        Assert.True(File.Exists(full), $"'{full}' was not copied next to the test assembly - check the test csproj's docsync ItemGroup.");
        return File.ReadAllText(full);
    }

    private static HashSet<string> LiveToolNames()
    {
        var names = typeof(ToolGuards).Assembly.GetTypes()
            .SelectMany(t => t.GetMethods(
                BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.Instance))
            .Select(m => m.GetCustomAttribute<McpServerToolAttribute>())
            .Where(a => a?.Name is not null)
            .Select(a => a!.Name!)
            .ToHashSet(StringComparer.Ordinal);
        Assert.NotEmpty(names);
        return names;
    }

    [Fact]
    public void Readme_tools_table_matches_the_live_tool_surface_exactly()
    {
        var readme = DocText("README.md");
        var toolsSection = ExtractSection(readme, "## Tools");
        // The header row's first cell is the literal word "Tool", not a backticked name, so the regex
        // collects data rows only.
        var documented = FirstCellToolName.Matches(toolsSection).Select(m => m.Groups[1].Value)
            .ToHashSet(StringComparer.Ordinal);
        var live = LiveToolNames();

        var missingFromDoc = live.Except(documented).Order().ToArray();
        var ghostsInDoc = documented.Except(live).Order().ToArray();
        Assert.True(missingFromDoc.Length == 0 && ghostsInDoc.Length == 0,
            $"README ## Tools table is out of sync with [McpServerTool] names. "
            + $"Missing from the table: [{string.Join(", ", missingFromDoc)}]; "
            + $"documented but not a live tool: [{string.Join(", ", ghostsInDoc)}].");
    }

    [Fact]
    public void Skill_mentions_every_live_tool_and_documents_no_ghost_tools()
    {
        var skill = DocText("SKILL.md");
        var live = LiveToolNames();

        // Every live tool must be mentioned (backticked) somewhere in the skill.
        var unmentioned = live.Where(n => !skill.Contains($"`{n}`", StringComparison.Ordinal))
            .Order().ToArray();

        // Every tool-shaped FIRST table cell in the skill must be a live tool (catches a renamed or
        // removed tool leaving a stale row behind).
        var ghosts = FirstCellToolName.Matches(skill).Select(m => m.Groups[1].Value)
            .Where(n => !live.Contains(n)).Distinct().Order().ToArray();

        Assert.True(unmentioned.Length == 0 && ghosts.Length == 0,
            $"skills/al-word-layout/SKILL.md is out of sync with [McpServerTool] names. "
            + $"Live tools never mentioned: [{string.Join(", ", unmentioned)}]; "
            + $"table rows for non-existent tools: [{string.Join(", ", ghosts)}].");
    }

    private static string ExtractSection(string markdown, string heading)
    {
        var start = markdown.IndexOf(heading, StringComparison.Ordinal);
        Assert.True(start >= 0, $"Heading '{heading}' not found.");
        var next = markdown.IndexOf("\n## ", start + heading.Length, StringComparison.Ordinal);
        return next < 0 ? markdown[start..] : markdown[start..next];
    }
}
