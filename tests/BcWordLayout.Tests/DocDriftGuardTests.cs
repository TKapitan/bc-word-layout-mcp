using System.Reflection;
using System.Text.Json;
using System.Text.RegularExpressions;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using ModelContextProtocol.Server;

namespace BcWordLayout.Tests;

/// <summary>
/// GitHub issue #114: the doc-drift guards checked tool NAMES only (<see cref="DocumentedToolNameSyncTests"/>)
/// plus the packaging identity (<see cref="PackagingMetadataTests"/>), so every count, list, parameter, error
/// code and manifest description written in prose drifted unnoticed (#100, #113). Each test here derives one
/// fact from the code and requires the docs to state it: the tool and mutating-tool counts, which tools are
/// mutating, every tool's parameters in the skill's tool reference, the names the design skill uses, the
/// error codes, and the skills each plugin and marketplace entry names. They read the checkout
/// (<see cref="RepoFiles"/>), like the other source-scan tests. The host-level behaviour #114 also asked for is
/// pinned by <c>McpHostPipelineTests</c>, and the shipped server by <c>scripts/smoke-release-artifacts.ps1</c>.
/// </summary>
public class DocDriftGuardTests
{
    // ---- facts from the code ----

    /// <summary>Every live tool, by name, with the method the SDK binds it to.</summary>
    private static readonly Lazy<IReadOnlyDictionary<string, MethodInfo>> LiveTools = new(() =>
        typeof(ToolGuards).Assembly.GetTypes()
            .Where(t => t.GetCustomAttribute<McpServerToolTypeAttribute>() is not null)
            .SelectMany(t => t.GetMethods(BindingFlags.Public | BindingFlags.Static))
            .Select(m => (Method: m, Name: m.GetCustomAttribute<McpServerToolAttribute>()?.Name))
            .Where(x => x.Name is not null)
            .ToDictionary(x => x.Name!, x => x.Method, StringComparer.Ordinal));

    private static IEnumerable<string> ParametersOf(string tool) =>
        LiveTools.Value[tool].GetParameters().Select(p => p.Name!);

    /// <summary>
    /// The mutating tools, derived from the code that actually saves: every tool whose body writes through one
    /// of <see cref="ToolGuards"/>' four <c>GuardMutate</c> instantiations. Nothing marks a tool as mutating by
    /// hand, so there is no marker to forget.
    /// </summary>
    private static readonly Lazy<IReadOnlySet<string>> MutatingTools = new(() =>
    {
        string[] writeWrappers = ["GuardEdit(", "GuardCellEdit(", "GuardTableEdit(", "GuardRefresh("];
        var mutating = new HashSet<string>(StringComparer.Ordinal);
        foreach (var file in RepoFiles.SourceFiles("src/BcWordLayout.McpHost/Tools"))
        {
            var text = File.ReadAllText(file);
            var tools = Regex.Matches(text, @"\[McpServerTool\(Name = ""(?<name>[a-z_]+)""\)\]");
            for (var i = 0; i < tools.Count; i++)
            {
                var end = i + 1 < tools.Count ? tools[i + 1].Index : text.Length;
                if (writeWrappers.Any(w => text.AsSpan(tools[i].Index, end - tools[i].Index).Contains(w, StringComparison.Ordinal)))
                {
                    mutating.Add(tools[i].Groups["name"].Value);
                }
            }
        }

        return mutating;
    });

    [Fact]
    public void The_derived_tool_facts_are_plausible()
    {
        // Pins the derivations above, so a broken scan cannot make every later test pass vacuously.
        Assert.Equal(27, LiveTools.Value.Count);
        Assert.Contains("insert_field", MutatingTools.Value);
        Assert.Contains("set_cell_borders", MutatingTools.Value);
        Assert.Contains("refresh_xml_part", MutatingTools.Value);
        Assert.DoesNotContain("get_layout_info", MutatingTools.Value);
        Assert.DoesNotContain("preview_layout", MutatingTools.Value);
        Assert.DoesNotContain("create_layout", MutatingTools.Value);
        Assert.Subset(LiveTools.Value.Keys.ToHashSet(), MutatingTools.Value.ToHashSet());
    }

    // ---- counts written in prose ----

    private static readonly string[] NumberWords =
    [
        "zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven",
        "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen", "twenty",
        "twenty-one", "twenty-two", "twenty-three", "twenty-four", "twenty-five", "twenty-six", "twenty-seven",
        "twenty-eight", "twenty-nine", "thirty",
    ];

    // "26 tools", "26 typed MCP tools", "26 deterministic, typed tools", "the 26 tool definitions",
    // "20 mutating tools". A count of some other group ("the 7 table-structure tools") does not match.
    private static readonly Regex ToolCount = new(
        @"\b(?<n>\d+|" + string.Join("|", NumberWords.OrderByDescending(w => w.Length)) + @")"
        + @"(?<adjectives>(?:[ ,]+(?:deterministic|typed|MCP|live|mutating))*)[ ]+(?:tools|tool definitions)\b",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>Every document and manifest that states how many tools there are.</summary>
    private static readonly string[] CountedFiles =
    [
        "README.md", "CONTRIBUTING.md", "docs/SOLUTION-DESIGN.md", "docs/RELEASING.md", "docs/BRANCHING.md",
        "docs/FIDELITY-CHECKLIST.md", "skills/al-word-layout/SKILL.md", "skills/al-word-layout-design/SKILL.md",
        "plugin.json", ".claude-plugin/plugin.json", ".claude-plugin/marketplace.json",
        ".github/plugin/marketplace.json", "src/BcWordLayout.McpHost/.mcp/server.json",
        "tools/release/verify_server.py",
    ];

    private static int Number(string text) =>
        int.TryParse(text, out var n) ? n : Array.FindIndex(NumberWords, w => w.Equals(text, StringComparison.OrdinalIgnoreCase));

    [Fact]
    public void Every_tool_count_in_the_docs_and_manifests_matches_the_code()
    {
        var total = LiveTools.Value.Count;
        var mutating = MutatingTools.Value.Count;
        var sources = CountedFiles.Select(f => (Name: f, Text: RepoFiles.Read(f))).ToList();

        // The NuGet description (the csproj <Description>, what nuget.org shows) reaches the assembly attribute.
        var description = typeof(ToolGuards).Assembly.GetCustomAttribute<AssemblyDescriptionAttribute>()?.Description;
        Assert.False(string.IsNullOrEmpty(description), "the McpHost assembly carries no description");
        sources.Add(("BcWordLayout.McpHost.csproj <Description>", description!));

        var checkedCount = 0;
        var wrong = new List<string>();
        foreach (var (name, text) in sources)
        {
            foreach (Match match in ToolCount.Matches(text))
            {
                checkedCount++;
                var isMutating = match.Groups["adjectives"].Value.Contains("mutating", StringComparison.OrdinalIgnoreCase);
                var expected = isMutating ? mutating : total;
                if (Number(match.Groups["n"].Value) != expected)
                {
                    var line = text.AsSpan(0, match.Index).Count('\n') + 1;
                    wrong.Add($"{name}:{line}: \"{match.Value}\" - the code has {expected}");
                }
            }
        }

        var verifyDefault = Regex.Match(RepoFiles.Read("tools/release/verify_server.py"), @"^EXPECTED_TOOLS_DEFAULT = (\d+)\r?$", RegexOptions.Multiline);
        Assert.True(verifyDefault.Success, "tools/release/verify_server.py has no 'EXPECTED_TOOLS_DEFAULT = <n>' line");
        if (int.Parse(verifyDefault.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture) != total)
        {
            wrong.Add($"tools/release/verify_server.py: EXPECTED_TOOLS_DEFAULT = {verifyDefault.Groups[1].Value} - the code has {total}");
        }

        Assert.True(checkedCount >= 8, $"only {checkedCount} tool counts found - has the phrasing changed?");
        Assert.True(wrong.Count == 0, "Tool counts that disagree with the code:" + Environment.NewLine + string.Join(Environment.NewLine, wrong));
    }

    // ---- the mutating-tool lists ----

    [Fact]
    public void The_readme_lists_exactly_the_mutating_tools()
    {
        var readme = RepoFiles.Read("README.md");
        var list = Regex.Match(readme, @"mutating tools\s*\((?<list>[^)]*)\)");
        Assert.True(list.Success, "README.md has no '… mutating tools (`a`, `b`, …)' list");
        var documented = Regex.Matches(list.Groups["list"].Value, "`([a-z_]+)`").Select(m => m.Groups[1].Value).ToHashSet();

        Assert.Equal(MutatingTools.Value.OrderBy(t => t), documented.OrderBy(t => t));
    }

    [Fact]
    public void The_skills_mutating_tables_hold_exactly_the_mutating_tools()
    {
        // The skill groups the mutating tools as its Edit and Table structure tables, plus refresh_xml_part,
        // which sits in the Lifecycle table beside create_layout (that one writes a NEW file, not through
        // the save-or-reject choreography).
        var skill = RepoFiles.Read("skills/al-word-layout/SKILL.md");
        var tabled = ToolRows(Section(skill, "### Edit (mutating")).Concat(ToolRows(Section(skill, "### Table structure (mutating")))
            .Select(r => r.Tool).ToHashSet();

        Assert.Equal(MutatingTools.Value.Except(["refresh_xml_part"]).OrderBy(t => t), tabled.OrderBy(t => t));
        Assert.Contains("refresh_xml_part", MutatingTools.Value);
    }

    // ---- the skill's tool reference ----

    /// <summary>
    /// Every live tool's row in the skill names each of its parameters, and every <c>`name`=…</c> default it
    /// states belongs to a real parameter. The location parameters are named once, by "`locationType` + its
    /// required params" (the shared location model), and the table-structure tools' <c>layoutPart</c>/
    /// <c>partName</c> once, in that table's introduction.
    /// </summary>
    [Fact]
    public void Every_tool_row_in_the_skill_names_each_parameter_and_only_real_ones()
    {
        var skill = RepoFiles.Read("skills/al-word-layout/SKILL.md");
        var tableStructure = ToolRows(Section(skill, "### Table structure (mutating")).Select(r => r.Tool).ToHashSet();
        string[] locationParameters = ["controlId", "tableIndex", "row", "col", "searchText"];

        var rows = ToolRows(skill).Where(r => LiveTools.Value.ContainsKey(r.Tool)).ToList();
        Assert.Equal(LiveTools.Value.Keys.OrderBy(t => t), rows.Select(r => r.Tool).Distinct().OrderBy(t => t));

        var problems = new List<string>();
        foreach (var (tool, keyParams) in rows)
        {
            var named = Regex.Matches(keyParams, "`([A-Za-z][A-Za-z0-9]*)`").Select(m => m.Groups[1].Value).ToHashSet(StringComparer.Ordinal);
            if (keyParams.Contains("`locationType` + its required params", StringComparison.Ordinal))
            {
                named.UnionWith(locationParameters);
            }

            if (tableStructure.Contains(tool))
            {
                named.UnionWith(["layoutPart", "partName"]);
            }

            var live = ParametersOf(tool).ToHashSet(StringComparer.Ordinal);
            problems.AddRange(live.Except(named).Select(p => $"{tool}: parameter `{p}` is not named in its Key params cell"));

            var withDefaults = Regex.Matches(keyParams, @"`([A-Za-z][A-Za-z0-9]*)`\s*=\s*`").Select(m => m.Groups[1].Value);
            problems.AddRange(withDefaults.Where(p => !live.Contains(p)).Select(p => $"{tool}: `{p}`=… is not one of its parameters"));
        }

        Assert.True(problems.Count == 0, string.Join(Environment.NewLine, problems));
    }

    [Fact]
    public void The_design_skill_names_only_live_tools_and_parameters()
    {
        var design = RepoFiles.Read("skills/al-word-layout-design/SKILL.md");
        var allParameters = LiveTools.Value.Keys.SelectMany(ParametersOf).ToHashSet(StringComparer.Ordinal);

        var tools = Regex.Matches(design, @"\b(?:insert|remove|set|clear|get|list|validate|preview|render|create|refresh|merge|split)_[a-z_]+\b")
            .Select(m => m.Value).Distinct();
        var unknownTools = tools.Where(t => !LiveTools.Value.ContainsKey(t)).ToList();

        // name=value arguments inside code spans; a namespace-prefixed OOXML attribute (w:val=) is not one.
        var arguments = Regex.Matches(design, "`[^`]+`").SelectMany(span => Regex.Matches(span.Value, @"(?<![:\w])([a-z][A-Za-z0-9]*)=")
            .Select(m => m.Groups[1].Value)).Distinct();
        var unknownArguments = arguments.Where(a => !allParameters.Contains(a)).ToList();

        Assert.True(unknownTools.Count == 0, "The design skill names tools that do not exist: " + string.Join(", ", unknownTools));
        Assert.True(unknownArguments.Count == 0, "The design skill passes arguments no tool has: " + string.Join(", ", unknownArguments));
    }

    // ---- error codes ----

    private static HashSet<string> LiveErrorCodes() =>
        typeof(ErrorCodes).GetFields(BindingFlags.NonPublic | BindingFlags.Static)
            .Where(f => f.IsLiteral)
            .Select(f => (string)f.GetRawConstantValue()!)
            .ToHashSet(StringComparer.Ordinal);

    [Fact]
    public void The_skill_lists_exactly_the_error_codes_the_server_returns()
    {
        var section = Section(RepoFiles.Read("skills/al-word-layout/SKILL.md"), "## 5.");
        var start = section.IndexOf("`code` is one of", StringComparison.Ordinal);
        var end = section.IndexOf("`hint` is always", start + 1, StringComparison.Ordinal);
        Assert.True(start >= 0 && end > start, "SKILL.md §5 has no \"`code` is one of … `hint` is always\" passage");

        var documented = Regex.Matches(section[start..end], "`([a-z]+(?:_[a-z]+)+)`")
            .Select(m => m.Groups[1].Value)
            .Where(token => !LiveTools.Value.ContainsKey(token)) // "`create_layout`-only" names a tool, not a code
            .ToHashSet(StringComparer.Ordinal);

        Assert.Equal(LiveErrorCodes().OrderBy(c => c), documented.OrderBy(c => c));
    }

    [Fact]
    public void No_envelope_spells_its_error_code_as_a_literal()
    {
        var literal = new Regex("\"(" + string.Join("|", LiveErrorCodes().Select(Regex.Escape)) + ")\"");
        var offenders = RepoFiles.SourceFiles("src")
            .Where(f => !f.EndsWith($"{Path.DirectorySeparatorChar}ErrorCodes.cs", StringComparison.Ordinal))
            .SelectMany(f =>
            {
                var text = File.ReadAllText(f);
                return literal.Matches(text).Select(m => $"{RepoFiles.Relative(f)}:{text.AsSpan(0, m.Index).Count('\n') + 1}: {m.Value}");
            })
            .ToList();

        Assert.True(offenders.Count == 0, "Use the ErrorCodes constant instead: " + string.Join(", ", offenders));
    }

    // ---- plugin and marketplace manifests ----

    [Fact]
    public void Every_plugin_and_marketplace_entry_names_every_skill()
    {
        var skills = Directory.GetDirectories(RepoFiles.PathOf("skills")).Select(Path.GetFileName).ToList();
        Assert.NotEmpty(skills);

        var descriptions = new List<(string Where, string Text)>();
        foreach (var manifest in new[] { "plugin.json", ".claude-plugin/plugin.json" })
        {
            using var json = JsonDocument.Parse(RepoFiles.Read(manifest));
            descriptions.Add((manifest, json.RootElement.GetProperty("description").GetString()!));
        }

        foreach (var marketplace in new[] { ".claude-plugin/marketplace.json", ".github/plugin/marketplace.json" })
        {
            using var json = JsonDocument.Parse(RepoFiles.Read(marketplace));
            foreach (var plugin in json.RootElement.GetProperty("plugins").EnumerateArray())
            {
                descriptions.Add(($"{marketplace} plugin '{plugin.GetProperty("name").GetString()}'", plugin.GetProperty("description").GetString()!));
            }
        }

        var missing = descriptions
            .SelectMany(d => skills.Where(s => !Regex.IsMatch(d.Text, $@"(?<![\w-]){Regex.Escape(s!)}(?![\w-])")).Select(s => $"{d.Where} does not name {s}"))
            .ToList();

        Assert.True(missing.Count == 0, string.Join(Environment.NewLine, missing));
    }

    // ---- Markdown helpers ----

    /// <summary>The text from the heading line starting with <paramref name="heading"/> to the next heading of the same level.</summary>
    private static string Section(string markdown, string heading)
    {
        var start = markdown.IndexOf("\n" + heading, StringComparison.Ordinal);
        Assert.True(start >= 0, $"no '{heading}' heading");
        var level = heading.TakeWhile(c => c == '#').Count();
        var next = Regex.Match(markdown[(start + 1)..], $@"\n#{{1,{level}}} ");
        return next.Success ? markdown.Substring(start + 1, next.Index) : markdown[(start + 1)..];
    }

    /// <summary>Every table row whose first cell is a backticked tool name, with its second cell.</summary>
    private static IEnumerable<(string Tool, string KeyParams)> ToolRows(string markdown) =>
        markdown.Split('\n')
            .Select(line => line.TrimEnd('\r'))
            .Select(line => Regex.Match(line, @"^\|\s*`(?<tool>[a-z_]+)`\s*\|(?<cells>.*)$"))
            .Where(m => m.Success)
            .Select(m => (m.Groups["tool"].Value, Regex.Split(m.Groups["cells"].Value, @"(?<!\\)\|")[0]));
}
