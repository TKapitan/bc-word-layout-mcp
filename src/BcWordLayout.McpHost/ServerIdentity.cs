using System.Reflection;
using Microsoft.Extensions.DependencyInjection;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace BcWordLayout.McpHost;

/// <summary>
/// What this server reports as <c>serverInfo</c> in the MCP initialize handshake: <see cref="Name"/> plus the
/// PRODUCT version (<see cref="Version"/>, e.g. <c>1.1.0</c> or <c>1.2.0-rc.1</c>).
/// </summary>
/// <remarks>
/// Left unset, the SDK falls back to the entry assembly's name and its four-part ASSEMBLY version, which
/// <c>Directory.Build.props</c> keeps numeric (<c>1.1.0.0</c>) because it cannot carry a <c>-rc.N</c> suffix. So
/// an RC and the release it became both answered <c>1.1.0.0</c>, and neither an MCP client nor the release smoke
/// test (<c>tools/release/verify_server.py --expect-version</c>) could tell which build it was talking to (GitHub
/// issue #112). The name stays exactly what that fallback reported, so a client keying on it sees no change.
/// </remarks>
internal static class ServerIdentity
{
    /// <summary>The <c>serverInfo.name</c> this server has always reported (the entry assembly's name).</summary>
    internal const string Name = "BcWordLayout.McpHost";

    /// <summary>
    /// The product version: this assembly's informational version (<c>&lt;InformationalVersion&gt;</c> in
    /// <c>Directory.Build.props</c>, bumped with <c>&lt;Version&gt;</c> for every release, RCs included) without the
    /// <c>+&lt;commit&gt;</c> build metadata the .NET SDK's source-revision support appends to it.
    /// </summary>
    internal static string Version { get; } = WithoutBuildMetadata(
        typeof(ServerIdentity).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
        ?? typeof(ServerIdentity).Assembly.GetName().Version?.ToString()
        ?? "0.0.0");

    /// <summary>A SemVer version without its build metadata: <c>1.2.0-rc.1+3f2a9c1</c> becomes <c>1.2.0-rc.1</c>.</summary>
    internal static string WithoutBuildMetadata(string version)
    {
        var plus = version.IndexOf('+', StringComparison.Ordinal);
        return plus < 0 ? version : version[..plus];
    }

    /// <summary>
    /// Reports <see cref="Name"/> and <see cref="Version"/> as <c>serverInfo</c>. Called by <c>Program.cs</c> and by
    /// the host-level tests' in-process server alike, so the tests see the identity the shipped server reports.
    /// </summary>
    internal static IMcpServerBuilder WithBcWordLayoutServerInfo(this IMcpServerBuilder builder)
    {
        builder.Services.Configure<McpServerOptions>(
            options => options.ServerInfo = new Implementation { Name = Name, Version = Version });
        return builder;
    }
}
