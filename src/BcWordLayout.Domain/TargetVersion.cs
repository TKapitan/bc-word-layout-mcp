namespace BcWordLayout.Domain;

/// <summary>Where a <see cref="TargetVersion"/> came from.</summary>
public enum TargetVersionSource
{
    /// <summary>The <c>application</c> property of the nearest AL <c>app.json</c> above the layout.</summary>
    AppJson,

    /// <summary>The <see cref="TargetVersionResolver.EnvironmentVariable"/> in the server's environment.</summary>
    Environment,

    /// <summary>Neither gave a version; <see cref="TargetVersion.Detail"/> says why.</summary>
    Unknown,
}

/// <summary>
/// The Business Central version a layout must work on, and where that came from (GitHub issue #152, ADR-0007).
/// It is the MINIMUM version, which is exactly what an AL project's <c>app.json</c> <c>application</c> states.
/// <see cref="TargetVersionResolver"/> resolves it for every call; the server never stores one.
/// </summary>
public sealed record TargetVersion
{
    private TargetVersion(BcVersion? version, TargetVersionSource source, string detail)
    {
        Version = version;
        Source = source;
        Detail = detail;
    }

    /// <summary>The target version; null when it is unknown.</summary>
    public BcVersion? Version { get; }

    public TargetVersionSource Source { get; }

    /// <summary>
    /// What the version was read from (<c>C:\Repo\app\app.json: application 29.0.0.0</c>,
    /// <c>BCWL_BC_VERSION=29</c>), or, for an unknown target, the reason there is none.
    /// </summary>
    public string Detail { get; }

    /// <summary><c>app.json</c>, <c>environment</c> or <c>unknown</c>: the <see cref="Source"/> as tools report it.</summary>
    public string SourceName => Source switch
    {
        TargetVersionSource.AppJson => "app.json",
        TargetVersionSource.Environment => "environment",
        _ => "unknown",
    };

    internal static TargetVersion FromAppJson(BcVersion version, string appJsonPath, string application) =>
        new(version, TargetVersionSource.AppJson, $"{appJsonPath}: application {application}");

    internal static TargetVersion FromEnvironment(BcVersion version, string rawValue, string? note) =>
        new(
            version,
            TargetVersionSource.Environment,
            $"{TargetVersionResolver.EnvironmentVariable}={rawValue.Trim()}" + (note is null ? string.Empty : $" ({note})"));

    /// <summary>An unknown target, with the <paramref name="reason"/> a caller can act on.</summary>
    public static TargetVersion Unknown(string reason) => new(null, TargetVersionSource.Unknown, reason);

    /// <summary>
    /// The target for a message: <c>BC 28.0 (2026 release wave 1), from app.json (C:\…\app.json: application
    /// 28.0.0.0)</c>, or <c>unknown (no app.json above the layout …)</c>.
    /// </summary>
    public string Describe() =>
        Version is { } version ? $"{version.Describe()}, from {SourceName} ({Detail})" : $"unknown ({Detail})";
}
