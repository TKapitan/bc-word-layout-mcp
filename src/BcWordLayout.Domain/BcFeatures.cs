namespace BcWordLayout.Domain;

/// <summary>Whether a <see cref="BcFeature"/> is something BC can do, or something it refuses.</summary>
public enum BcFeatureKind
{
    /// <summary>Something BC can do from <see cref="BcFeature.Since"/>: the first version confirmed to have it.</summary>
    Capability,

    /// <summary>Something BC refuses from <see cref="BcFeature.Since"/>: the first version that might enforce it.</summary>
    Restriction,
}

/// <summary>
/// One layout feature that depends on the Business Central version, with the evidence for its boundary. See
/// <see cref="BcFeatures"/>.
/// </summary>
/// <param name="Id">Stable id, e.g. <c>report-themes</c>; used in messages and the README's version table.</param>
/// <param name="Since">The first version it applies to.</param>
/// <param name="Until">The first version it no longer applies to, or null while it still does.</param>
/// <param name="What">What it covers, as a sentence fragment for messages.</param>
/// <param name="Evidence">Where the boundary comes from.</param>
public sealed record BcFeature(string Id, BcFeatureKind Kind, BcVersion Since, BcVersion? Until, string What, string Evidence)
{
    /// <summary>True when the feature applies to <paramref name="version"/>: from <see cref="Since"/>, and before <see cref="Until"/>.</summary>
    public bool AppliesTo(BcVersion version) => version >= Since && (Until is not { } until || version < until);
}

/// <summary>
/// The catalog of version-dependent layout features (GitHub issue #152). Tools gate what they AUTHOR with
/// <see cref="Require"/>; the validator's <c>bc-version</c> check reports catalog features a layout already
/// uses. The BC 27, 28, 30 and 31 features each become one more entry. The README's version table lists exactly
/// these entries, and a test keeps the two equal.
/// </summary>
public static class BcFeatures
{
    /// <summary>
    /// Report themes: a report theme <c>.dotx</c>, a header/footer layout bound to Tenant Report Defaults, and the
    /// planned theme tools (previewing under a theme, editing a theme, body/theme style checks).
    /// </summary>
    public static readonly BcFeature ReportThemes = new(
        "report-themes",
        BcFeatureKind.Capability,
        new BcVersion(29, 0),
        null,
        "report themes: a theme .dotx, and header/footer layouts bound to Tenant_Report_Defaults",
        "2026 release wave 2; a composite body + header/footer + theme set printed on 2026-10-17");

    /// <summary>
    /// Authoring themable formatting: theme-slot values for text colour, cell fill and border colour, and paragraph
    /// styles. Only what the tools write is gated: theme colours and paragraph styles already in a layout print on
    /// BC 28, and Microsoft's own BC 28 layouts carry them, so the validator does not flag them.
    /// </summary>
    public static readonly BcFeature ThemableFormatting = new(
        "themable-formatting",
        BcFeatureKind.Capability,
        new BcVersion(29, 0),
        null,
        "themable formatting: theme-slot colours for text, cell fills and borders, and paragraph styles",
        "decided with report themes, 2026-10-27 (GitHub issue #152)");

    /// <summary>Every entry, in the order the README's version table lists them.</summary>
    public static IReadOnlyList<BcFeature> All { get; } = [ReportThemes, ThemableFormatting];

    /// <summary>
    /// Refuses a call that would author <paramref name="capability"/> for a <paramref name="target"/> without it:
    /// below <see cref="BcFeature.Since"/>, from <see cref="BcFeature.Until"/> on, or unknown (ADR-0003: refuse
    /// rather than guess).
    /// </summary>
    /// <exception cref="ArgumentException"><paramref name="capability"/> is a restriction, which nothing requires.</exception>
    /// <exception cref="UnsupportedBcVersionException">The target lacks the feature, or is unknown.</exception>
    public static void Require(BcFeature capability, TargetVersion target)
    {
        ArgumentNullException.ThrowIfNull(capability);
        ArgumentNullException.ThrowIfNull(target);

        if (capability.Kind != BcFeatureKind.Capability)
        {
            throw new ArgumentException(
                $"'{capability.Id}' is a restriction; only a capability can be required.", nameof(capability));
        }

        if (target.Version is not { } version || !capability.AppliesTo(version))
        {
            throw new UnsupportedBcVersionException(capability, target);
        }
    }
}
