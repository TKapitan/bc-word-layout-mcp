using System.Reflection;
using System.Text.Json;
using BcWordLayout.McpHost.Tools;

namespace BcWordLayout.Tests;

/// <summary>
/// Guards the version/identity coupling between <c>Directory.Build.props</c> (the single version
/// source), the NuGet MCP manifest (<c>src/BcWordLayout.McpHost/.mcp/server.json</c> — what NuGet.org
/// renders as the copy-paste install config), and the Claude Code plugin manifest
/// (<c>.claude-plugin/plugin.json</c> — whose pinned <c>dnx</c> reference is what plugin users
/// actually launch). A release bump that misses one of the three files ships a manifest pointing at
/// the wrong package version; these tests turn that into a suite failure instead. The manifests are
/// copied next to the test assembly at build time (see the csproj), so the tests stay hermetic — no
/// repo-root probing at runtime.
/// </summary>
public class PackagingMetadataTests
{
    /// <summary>The package id decided in the release plan (D3) and baked into every install snippet.</summary>
    private const string ExpectedPackageId = "BcWordLayout.Mcp";

    private static string PackagingFile(string fileName)
    {
        var full = Path.Combine(AppContext.BaseDirectory, "packaging", fileName);
        Assert.True(File.Exists(full), $"'{full}' was not copied next to the test assembly - check the test csproj's packaging ItemGroup.");
        return full;
    }

    /// <summary>
    /// The product version as built, from the McpHost assembly's informational version with any
    /// SourceLink <c>+&lt;commit&gt;</c> suffix stripped — the exact value <c>&lt;Version&gt;</c> in
    /// <c>Directory.Build.props</c> declares.
    /// </summary>
    private static string ProductVersion()
    {
        var informational = typeof(ToolGuards).Assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()!.InformationalVersion;
        var plus = informational.IndexOf('+');
        return plus < 0 ? informational : informational[..plus];
    }

    [Fact]
    public void NuGet_server_manifest_matches_the_built_product_version_and_identity()
    {
        using var doc = JsonDocument.Parse(File.ReadAllText(PackagingFile("server.json")));
        var root = doc.RootElement;
        var version = ProductVersion();

        Assert.Equal("io.github.tkapitan/bc-word-layout-mcp", root.GetProperty("name").GetString());
        Assert.Equal(version, root.GetProperty("version").GetString());

        var package = Assert.Single(root.GetProperty("packages").EnumerateArray());
        Assert.Equal("nuget", package.GetProperty("registryType").GetString());
        Assert.Equal(ExpectedPackageId, package.GetProperty("identifier").GetString());
        Assert.Equal(version, package.GetProperty("version").GetString());
        Assert.Equal("stdio", package.GetProperty("transport").GetProperty("type").GetString());
    }

    /// <summary>
    /// The MCP registry's server.json schema (2025-10-17) caps every <c>description</c> at 100
    /// characters — the registry rejects a manifest over the limit, and the original manifest shipped
    /// three descriptions over it (found 2026-08-08 via the IDE's schema validation). Walks the whole
    /// document so a description added anywhere later (a new environment variable, say) is covered
    /// without this test knowing the paths.
    /// </summary>
    [Fact]
    public void NuGet_server_manifest_descriptions_fit_the_registry_schema_limit()
    {
        using var doc = JsonDocument.Parse(File.ReadAllText(PackagingFile("server.json")));

        var over = new List<string>();
        void Walk(JsonElement element, string path)
        {
            switch (element.ValueKind)
            {
                case JsonValueKind.Object:
                    foreach (var property in element.EnumerateObject())
                    {
                        if (property.Name == "description"
                            && property.Value.ValueKind == JsonValueKind.String
                            && property.Value.GetString()!.Length > 100)
                        {
                            over.Add($"{path}.description ({property.Value.GetString()!.Length} chars)");
                        }

                        Walk(property.Value, $"{path}.{property.Name}");
                    }

                    break;
                case JsonValueKind.Array:
                    var index = 0;
                    foreach (var item in element.EnumerateArray())
                    {
                        Walk(item, $"{path}[{index++}]");
                    }

                    break;
            }
        }

        Walk(doc.RootElement, "$");
        Assert.True(over.Count == 0,
            "server.json descriptions over the registry's 100-char cap: " + string.Join("; ", over));
    }

    /// <summary>
    /// Covers both plugin manifests: the Claude Code one (<c>.claude-plugin/plugin.json</c>) and the
    /// VS Code / GitHub Copilot one (root <c>plugin.json</c>). They pin the same dnx reference and
    /// must both follow a version bump.
    /// </summary>
    [Theory]
    [InlineData("plugin.json")]
    [InlineData("copilot-plugin.json")]
    public void Plugin_manifests_pin_the_dnx_reference_to_the_built_product_version(string fileName)
    {
        using var doc = JsonDocument.Parse(File.ReadAllText(PackagingFile(fileName)));
        var root = doc.RootElement;
        var version = ProductVersion();

        Assert.Equal(version, root.GetProperty("version").GetString());

        var server = root.GetProperty("mcpServers").GetProperty("bc-word-layout");
        Assert.Equal("dnx", server.GetProperty("command").GetString());
        var args = server.GetProperty("args").EnumerateArray().Select(a => a.GetString()).ToArray();
        Assert.Equal($"{ExpectedPackageId}@{version}", args[0]);
        Assert.Contains("--yes", args);
    }
}
