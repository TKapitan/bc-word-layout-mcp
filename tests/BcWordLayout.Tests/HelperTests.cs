using BcWordLayout.Domain;

namespace BcWordLayout.Tests;

public class LabelConventionTests
{
    [Theory]
    [InlineData("BilledTo_Lbl", true)]
    [InlineData("BillToContactEmailLbl", true)]
    [InlineData("CompanyLegalOffice_Lbl", true)]
    [InlineData("ItemNo_Line", false)]
    [InlineData("CustomerAddress1", false)]
    [InlineData("", false)]
    [InlineData(null, false)]
    public void IsLabelName_matches_convention(string? name, bool expected)
    {
        // Deliberately LabelConvention.Default (not .Current): this test pins the BC-STANDARD convention's
        // own matching rule regardless of whatever a parallel test class may have installed on .Current
        // (see LabelConventionSeamCollection for the process-wide-static isolation rule) - it needs no
        // seam-collection membership as a result.
        Assert.Equal(expected, LabelConvention.Default.IsLabelName(name));
    }
}

public class BindingXPathTests
{
    [Fact]
    public void Segments_strips_prefixes_and_indices()
    {
        var xpath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:ItemNo_Line[1]";
        Assert.Equal(new[] { "NavWordReportXmlPart", "Header", "Line", "ItemNo_Line" }, BindingXPath.Segments(xpath));
    }

    [Fact]
    public void Segments_handles_unindexed_final_step()
    {
        var xpath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line";
        Assert.Equal("Line", BindingXPath.LeafName(xpath));
    }

    [Fact]
    public void Segments_empty_for_null()
    {
        Assert.Empty(BindingXPath.Segments(null));
        Assert.Null(BindingXPath.LeafName(""));
    }
}
