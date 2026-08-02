using BcWordLayout.McpHost;

namespace BcWordLayout.Tests;

/// <summary>
/// Covers <see cref="EnvironmentConfig.ParseLabelSuffixes"/> and
/// <see cref="EnvironmentConfig.ParseLabelsDataItemName"/> — the pure parsing helpers behind
/// <c>BCWL_LABEL_SUFFIXES</c>/<c>BCWL_LABELS_DATA_ITEM</c>. Drives the helpers directly rather than
/// spawning the actual MCP host process: they touch no
/// process-wide static (unlike <see cref="LabelConventionConfigTests"/>), so this class needs no
/// seam-collection membership.
/// </summary>
public class EnvironmentConfigTests
{
    [Theory]
    [InlineData("Lbl", new[] { "Lbl" })]
    [InlineData("Lbl,Caption", new[] { "Lbl", "Caption" })]
    [InlineData(" Lbl , Caption ", new[] { "Lbl", "Caption" })]
    [InlineData("Lbl,Lbl,Caption", new[] { "Lbl", "Caption" })]
    public void ParseLabelSuffixes_parses_a_valid_value(string rawValue, string[] expected)
    {
        var result = EnvironmentConfig.ParseLabelSuffixes(rawValue);

        Assert.NotNull(result);
        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(",")]
    [InlineData(" , ,, ")]
    public void ParseLabelSuffixes_returns_null_for_an_invalid_or_unusable_value(string? rawValue)
    {
        Assert.Null(EnvironmentConfig.ParseLabelSuffixes(rawValue));
    }

    [Fact]
    public void LabelSuffixesVariable_is_the_documented_name()
    {
        Assert.Equal("BCWL_LABEL_SUFFIXES", EnvironmentConfig.LabelSuffixesVariable);
    }

    [Theory]
    [InlineData("Labels", "Labels")]
    [InlineData("  Labels  ", "Labels")]
    [InlineData("EtiquetasDeInforme", "EtiquetasDeInforme")]
    public void ParseLabelsDataItemName_parses_a_valid_value(string rawValue, string expected)
    {
        Assert.Equal(expected, EnvironmentConfig.ParseLabelsDataItemName(rawValue));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("/Report/Labels")] // a path, not a single data-item name - the rule matches one name only
    public void ParseLabelsDataItemName_returns_null_for_an_invalid_or_unusable_value(string? rawValue)
    {
        Assert.Null(EnvironmentConfig.ParseLabelsDataItemName(rawValue));
    }

    [Fact]
    public void LabelsDataItemVariable_is_the_documented_name()
    {
        Assert.Equal("BCWL_LABELS_DATA_ITEM", EnvironmentConfig.LabelsDataItemVariable);
    }
}
