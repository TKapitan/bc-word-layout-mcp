using BcWordLayout.Domain;

namespace BcWordLayout.Tests;

/// <summary>
/// <see cref="BcFeatures"/>: the catalog of version-dependent features and its gate (GitHub issue #152). No shipped
/// tool requires a feature yet, so the gate is pinned here and its envelope in <c>BcVersionCheckTests</c>.
/// </summary>
public class BcFeaturesTests
{
    private static TargetVersion Known(int major, int minor) =>
        TargetVersion.FromEnvironment(new BcVersion(major, minor), $"{major}.{minor}", note: null);

    [Fact]
    public void Catalog_ids_are_unique_and_both_entries_start_at_29()
    {
        Assert.Equal(BcFeatures.All.Count, BcFeatures.All.Select(f => f.Id).Distinct(StringComparer.Ordinal).Count());
        Assert.Equal(["report-themes", "themable-formatting"], BcFeatures.All.Select(f => f.Id));
        Assert.All(BcFeatures.All, f =>
        {
            Assert.Equal(BcFeatureKind.Capability, f.Kind);
            Assert.Equal(new BcVersion(29, 0), f.Since);
            Assert.Null(f.Until);
            Assert.False(string.IsNullOrWhiteSpace(f.Evidence));
        });
    }

    [Theory]
    [InlineData(29, 0)]
    [InlineData(29, 1)]
    [InlineData(31, 0)]
    public void Require_passes_from_the_first_version_with_the_feature(int major, int minor)
    {
        BcFeatures.Require(BcFeatures.ThemableFormatting, Known(major, minor));
    }

    [Fact]
    public void Require_refuses_a_lower_target_and_names_both_versions()
    {
        var target = Known(28, 3);

        var ex = Assert.Throws<UnsupportedBcVersionException>(() => BcFeatures.Require(BcFeatures.ReportThemes, target));

        Assert.Same(BcFeatures.ReportThemes, ex.Feature);
        Assert.Same(target, ex.Target);
        Assert.Contains("('report-themes')", ex.Message, StringComparison.Ordinal);
        Assert.Contains("from BC 29.0 (2026 release wave 2)", ex.Message, StringComparison.Ordinal);
        Assert.Contains("the target is BC 28.3 (2026 release wave 1), from environment (BCWL_BC_VERSION=28.3)", ex.Message, StringComparison.Ordinal);
        Assert.EndsWith("Nothing was changed.", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Require_refuses_an_unknown_target_rather_than_guessing()
    {
        // ADR-0003: the tools cannot tell whether BC has the feature, so a gated call is refused.
        var ex = Assert.Throws<UnsupportedBcVersionException>(
            () => BcFeatures.Require(BcFeatures.ThemableFormatting, TargetVersion.Unknown("BCWL_BC_VERSION is not set")));

        Assert.Contains("the target BC version is unknown (BCWL_BC_VERSION is not set)", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_feature_ends_at_its_Until_version()
    {
        var retired = new BcFeature("retired", BcFeatureKind.Capability, new BcVersion(27, 0), new BcVersion(30, 0), "x", "test");

        Assert.False(retired.AppliesTo(new BcVersion(26, 5)));
        Assert.True(retired.AppliesTo(new BcVersion(29, 9)));
        Assert.False(retired.AppliesTo(new BcVersion(30, 0)));
        Assert.Contains(" until 30.0", new UnsupportedBcVersionException(retired, Known(30, 0)).Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Only_a_capability_can_be_required()
    {
        var restriction = new BcFeature("refused", BcFeatureKind.Restriction, new BcVersion(28, 0), null, "x", "test");

        Assert.Throws<ArgumentException>(() => BcFeatures.Require(restriction, Known(29, 0)));
    }
}
