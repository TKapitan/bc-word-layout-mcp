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
