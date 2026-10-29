using BcWordLayout.Domain;

namespace BcWordLayout.Tests;

/// <summary>
/// <see cref="BcVersion"/>: the major.minor form every target is compared in (GitHub issue #152). An AL
/// <c>app.json</c> states four parts; <c>BCWL_BC_VERSION</c> one or two.
/// </summary>
public class BcVersionTests
{
    [Theory]
    [InlineData("29", 29, 0)]
    [InlineData("29.1", 29, 1)]
    [InlineData("29.0.0.0", 29, 0)] // app.json's application
    [InlineData("28.3.12345.0", 28, 3)]
    [InlineData(" 29 ", 29, 0)]
    public void TryParse_reads_one_to_four_parts_as_major_minor(string text, int major, int minor)
    {
        Assert.True(BcVersion.TryParse(text, out var version));
        Assert.Equal(new BcVersion(major, minor), version);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("  ")]
    [InlineData("abc")]
    [InlineData("29.")]
    [InlineData(".1")]
    [InlineData("-29")]
    [InlineData("+29")]
    [InlineData("29 .1")]
    [InlineData("29,1")]
    [InlineData("29.0.0.0.0")]
    [InlineData("0")]
    public void TryParse_refuses_anything_else(string? text)
    {
        Assert.False(BcVersion.TryParse(text, out _));
    }

    [Fact]
    public void Versions_order_by_major_then_minor()
    {
        var v283 = new BcVersion(28, 3);
        var v290 = new BcVersion(29, 0);
        var v291 = new BcVersion(29, 1);

        Assert.True(v283 < v290);
        Assert.True(v290 < v291);
        Assert.True(v291 >= v290);
        Assert.True(v290 <= new BcVersion(29, 0));
        Assert.Equal(new[] { v283, v290, v291 }, new[] { v291, v283, v290 }.Order());
    }

    [Theory]
    [InlineData(29, "2026 release wave 2")]
    [InlineData(28, "2026 release wave 1")]
    [InlineData(27, "2025 release wave 2")]
    [InlineData(30, "2027 release wave 1")]
    [InlineData(14, "2019 release wave 1")]
    [InlineData(13, null)]
    public void ReleaseWave_names_the_twice_yearly_release(int major, string? wave)
    {
        Assert.Equal(wave, new BcVersion(major, 0).ReleaseWave);
    }

    [Fact]
    public void A_version_reads_as_major_minor_with_its_release_wave()
    {
        Assert.Equal("29.0", new BcVersion(29, 0).ToString());
        Assert.Equal("BC 29.0 (2026 release wave 2)", new BcVersion(29, 0).Describe());
        Assert.Equal("BC 13.0", new BcVersion(13, 0).Describe());
    }
}
