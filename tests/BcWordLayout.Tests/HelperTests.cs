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

    [Fact]
    public void Prefixes_lists_each_step_prefix_once_in_first_use_order()
    {
        Assert.Equal(
            new[] { "ns0", "ns1" },
            BindingXPath.Prefixes("/ns0:NavWordReportXmlPart[1]/ns1:Header[1]/ns0:Line[1]/ns1:ItemNo_Line[1]"));
        Assert.Empty(BindingXPath.Prefixes("/NavWordReportXmlPart[1]/Header[1]"));
        Assert.Empty(BindingXPath.Prefixes(null));
    }

    [Fact]
    public void PrefixDeclarations_parses_xmlns_declarations_and_nothing_from_a_bare_uri()
    {
        Assert.Equal(
            new[] { ("ns0", "urn:a/"), ("ns1", "urn:b/") },
            BindingXPath.PrefixDeclarations("xmlns:ns0='urn:a/' xmlns:ns1=\"urn:b/\" xmlns:ns0='urn:ignored/'"));
        Assert.Empty(BindingXPath.PrefixDeclarations("urn:microsoft-dynamics-nav/reports/Mini_Sales_Invoice/1306/"));
        Assert.Empty(BindingXPath.PrefixDeclarations(null));
    }

    [Fact]
    public void IsBcDatasetBinding_is_decided_by_the_root_step_whatever_its_prefix()
    {
        Assert.True(BindingXPath.IsBcDatasetBinding("/ns0:NavWordReportXmlPart[1]/ns0:Header[1]"));
        Assert.True(BindingXPath.IsBcDatasetBinding("/x:NavWordReportXmlPart[1]"));
        Assert.False(BindingXPath.IsBcDatasetBinding("/ns1:coreProperties[1]/ns0:title[1]"));
        Assert.False(BindingXPath.IsBcDatasetBinding(null));
    }
}
