using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// Direct coverage of <see cref="SdtInspector"/> — the single classification/property-reading
/// implementation that replaced four independently-diverged <c>Classify</c>/<c>ClassifyKind</c> copies and
/// the copy-pasted <c>FindBinding</c>/<c>ReadControlId</c>/<c>FirstChild</c>/<c>HasChild</c>/<c>Attr</c>
/// helper family. Several tests below pin the exact divergences the
/// consolidation resolved or deliberately preserved — see <see cref="SdtInspector"/>'s own remarks.
/// </summary>
public class SdtInspectorTests
{
    private const string SomeItemId = SyntheticLayout.GoodItemId;

    private static SdtElement FirstSdt(string path)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        return doc.MainDocumentPart!.Document!.Body!.Descendants<SdtElement>().First();
    }

    private static SdtElement NthSdt(string path, int index)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        return doc.MainDocumentPart!.Document!.Body!.Descendants<SdtElement>().ElementAt(index);
    }

    // ---- Classify: branch order / marker precedence ----

    [Fact]
    public void Classify_returns_Unbound_when_sdtPr_is_null()
    {
        Assert.Equal(SdtInspector.Classification.Unbound, SdtInspector.Classify(null));
    }

    [Fact]
    public void Classify_recognizes_a_repeatingSection_even_though_it_also_carries_a_dataBinding()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.ProperRepeater(
            "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line", SomeItemId));
        try
        {
            var sdt = FirstSdt(path);
            Assert.Equal(SdtInspector.Classification.Repeater, SdtInspector.Classify(sdt.GetFirstChild<SdtProperties>()));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Classify_recognizes_a_repeatingSectionItem()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.ProperRepeater(
            "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line", SomeItemId));
        try
        {
            var itemSdt = NthSdt(path, 1); // the inner repeatingSectionItem sdt
            Assert.Equal(SdtInspector.Classification.RepeaterItem, SdtInspector.Classify(itemSdt.GetFirstChild<SdtProperties>()));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Classify_picture_marker_wins_over_a_dataBinding_on_the_same_sdtPr()
    {
        // Real BC picture controls (e.g. CompanyPicture) carry BOTH w:picture and w:dataBinding — the
        // picture marker must be checked first or the control would misclassify as a bound field.
        var body =
            "<w:sdt><w:sdtPr><w:id w:val=\"1\"/>"
            + $"<w:dataBinding w:xpath=\"/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:CompanyPicture[1]\" w:storeItemID=\"{SomeItemId}\"/>"
            + "<w:picture/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:sdtContent></w:sdt>";
        var path = SyntheticLayout.Create(body);
        try
        {
            var sdt = FirstSdt(path);
            Assert.Equal(SdtInspector.Classification.Picture, SdtInspector.Classify(sdt.GetFirstChild<SdtProperties>()));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Classify_dataBinding_only_is_Bound()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.BoundField(
            "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:CompanyName[1]", SomeItemId));
        try
        {
            var sdt = FirstSdt(path);
            Assert.Equal(SdtInspector.Classification.Bound, SdtInspector.Classify(sdt.GetFirstChild<SdtProperties>()));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Classify_no_marker_and_no_binding_is_Unbound()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.InlineControlWithId(1));
        try
        {
            var sdt = FirstSdt(path);
            Assert.Equal(SdtInspector.Classification.Unbound, SdtInspector.Classify(sdt.GetFirstChild<SdtProperties>()));
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- ClassifyControlKind: the higher-level ControlKind mapping ----

    [Fact]
    public void ClassifyControlKind_maps_a_label_suffixed_bound_leaf_to_Label()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.BoundField(
            "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:CompanyName_Lbl[1]", SomeItemId));
        try
        {
            var sdt = FirstSdt(path);
            Assert.Equal(ControlKind.Label, SdtInspector.ClassifyControlKind(sdt));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void ClassifyControlKind_maps_a_non_label_bound_leaf_to_Field()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.BoundField(
            "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:CompanyName[1]", SomeItemId));
        try
        {
            var sdt = FirstSdt(path);
            Assert.Equal(ControlKind.Field, SdtInspector.ClassifyControlKind(sdt));
        }
        finally
        {
            File.Delete(path);
        }
    }

    /// <summary>
    /// Pins the fix for a confirmed drift: <c>TableStructureReader.ClassifyKind</c> used to have no
    /// <c>repeatingSectionItem</c> branch, so an item wrapper that happened to carry its own binding would
    /// have fallen through to Field/Label there while every other copy mapped it to Unbound. The single
    /// <see cref="SdtInspector.ClassifyControlKind(SdtProperties?)"/> now applies the same branch order
    /// everywhere, so a bound repeatingSectionItem (never seen in the real corpus, but structurally
    /// possible) classifies as Unbound regardless of which caller asks.
    /// </summary>
    [Fact]
    public void ClassifyControlKind_maps_a_bound_repeatingSectionItem_to_Unbound_not_Field()
    {
        var body =
            "<w:sdt><w:sdtPr><w:id w:val=\"1\"/><w15:repeatingSectionItem/>"
            + $"<w:dataBinding w:xpath=\"/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]\" w:storeItemID=\"{SomeItemId}\"/>"
            + "</w:sdtPr><w:sdtContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:sdtContent></w:sdt>";
        var path = SyntheticLayout.Create(body);
        try
        {
            var sdt = FirstSdt(path);
            Assert.Equal(ControlKind.Unbound, SdtInspector.ClassifyControlKind(sdt));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void ClassifyControlKind_maps_Repeater_and_Picture_directly()
    {
        var repeaterPath = SyntheticLayout.Create(SyntheticLayout.ProperRepeater(
            "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line", SomeItemId));
        var picturePath = SyntheticLayout.Create(
            "<w:sdt><w:sdtPr><w:id w:val=\"1\"/>"
            + $"<w:dataBinding w:xpath=\"/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:CompanyPicture[1]\" w:storeItemID=\"{SomeItemId}\"/>"
            + "<w:picture/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:sdtContent></w:sdt>");
        try
        {
            Assert.Equal(ControlKind.Repeater, SdtInspector.ClassifyControlKind(FirstSdt(repeaterPath)));
            Assert.Equal(ControlKind.Picture, SdtInspector.ClassifyControlKind(FirstSdt(picturePath)));
        }
        finally
        {
            File.Delete(repeaterPath);
            File.Delete(picturePath);
        }
    }

    // ---- FindBinding vs FindRepeaterBinding: the deliberate divergence ----

    [Fact]
    public void FindBinding_prefers_legacy_w_dataBinding_over_w15_when_both_are_present()
    {
        var body =
            "<w:sdt><w:sdtPr><w:id w:val=\"1\"/>"
            + $"<w:dataBinding w:xpath=\"/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:CompanyName[1]\" w:storeItemID=\"{SomeItemId}\"/>"
            + $"<w15:dataBinding w:xpath=\"/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line\" w:storeItemID=\"{SomeItemId}\"/>"
            + "</w:sdtPr><w:sdtContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:sdtContent></w:sdt>";
        var path = SyntheticLayout.Create(body);
        try
        {
            var sdt = FirstSdt(path);
            Assert.Equal("/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:CompanyName[1]", SdtInspector.ReadXPath(sdt));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void FindBinding_falls_back_to_w15_when_no_legacy_binding_is_present()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.ProperRepeater(
            "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line", SomeItemId));
        try
        {
            var sdt = FirstSdt(path);
            Assert.Equal("/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line", SdtInspector.ReadXPath(sdt));
            Assert.True(SdtInspector.UsesW15Binding(sdt));
        }
        finally
        {
            File.Delete(path);
        }
    }

    /// <summary>
    /// Pins the deliberate divergence documented on <see cref="SdtInspector.FindRepeaterBinding"/>: unlike
    /// the general <see cref="SdtInspector.FindBinding(SdtProperties)"/> (legacy-preferred, w15 fallback),
    /// a repeater's OWN row binding is looked up via w15 ONLY — proven here by an sdtPr that carries BOTH,
    /// where the two methods must disagree.
    /// </summary>
    [Fact]
    public void FindRepeaterBinding_ignores_a_legacy_w_dataBinding_that_FindBinding_would_have_preferred()
    {
        var body =
            "<w:sdt><w:sdtPr><w:id w:val=\"1\"/>"
            + $"<w:dataBinding w:xpath=\"/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:CompanyName[1]\" w:storeItemID=\"{SomeItemId}\"/>"
            + $"<w15:dataBinding w:xpath=\"/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line\" w:storeItemID=\"{SomeItemId}\"/>"
            + "</w:sdtPr><w:sdtContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:sdtContent></w:sdt>";
        var path = SyntheticLayout.Create(body);
        try
        {
            var pr = FirstSdt(path).GetFirstChild<SdtProperties>()!;

            var general = SdtInspector.FindBinding(pr);
            var repeaterOnly = SdtInspector.FindRepeaterBinding(pr);

            Assert.Equal("/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:CompanyName[1]", SdtInspector.Attr(general!, "xpath", OoxmlNames.W));
            Assert.Equal("/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line", SdtInspector.Attr(repeaterOnly!, "xpath", OoxmlNames.W));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void FindRepeaterBinding_returns_null_when_only_a_legacy_binding_is_present()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.BoundField(
            "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:CompanyName[1]", SomeItemId));
        try
        {
            var pr = FirstSdt(path).GetFirstChild<SdtProperties>()!;
            Assert.Null(SdtInspector.FindRepeaterBinding(pr));
            Assert.NotNull(SdtInspector.FindBinding(pr));
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- property readback ----

    [Fact]
    public void ReadAlias_ReadTag_ReadControlId_ReadStoreItemId_round_trip()
    {
        var body =
            "<w:sdt><w:sdtPr>"
            + "<w:alias w:val=\"#Nav: /Header/CompanyName\"/>"
            + "<w:tag w:val=\"#Nav: TestReport/50000\"/>"
            + "<w:id w:val=\"-42\"/>"
            + $"<w:dataBinding w:xpath=\"/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:CompanyName[1]\" w:storeItemID=\"{SomeItemId}\"/>"
            + "</w:sdtPr><w:sdtContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:sdtContent></w:sdt>";
        var path = SyntheticLayout.Create(body);
        try
        {
            var sdt = FirstSdt(path);
            Assert.Equal("#Nav: /Header/CompanyName", SdtInspector.ReadAlias(sdt));
            Assert.Equal("#Nav: TestReport/50000", SdtInspector.ReadTag(sdt));
            Assert.Equal(-42, SdtInspector.ReadControlId(sdt));
            Assert.Equal(SomeItemId, SdtInspector.ReadStoreItemId(sdt));
            Assert.False(SdtInspector.UsesW15Binding(sdt));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void ReadControlId_returns_null_when_the_sdt_has_no_id()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.BoundField(
            "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:CompanyName[1]", SomeItemId));
        try
        {
            var sdt = FirstSdt(path);
            Assert.Null(SdtInspector.ReadControlId(sdt));
            Assert.Null(SdtInspector.ReadAlias(sdt));
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- IsRepeater / IsRepeaterItem / NearestRepeaterAncestor ----

    [Fact]
    public void IsRepeater_and_IsRepeaterItem_and_NearestRepeaterAncestor_agree_on_a_real_shaped_chain()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.ProperRepeater(
            "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line", SomeItemId));
        try
        {
            using var doc = WordprocessingDocument.Open(path, false);
            var sdts = doc.MainDocumentPart!.Document!.Body!.Descendants<SdtElement>().ToList();
            var repeater = sdts[0];
            var item = sdts[1];

            Assert.True(SdtInspector.IsRepeater(repeater));
            Assert.False(SdtInspector.IsRepeaterItem(repeater));

            Assert.True(SdtInspector.IsRepeaterItem(item));
            Assert.False(SdtInspector.IsRepeater(item));
            Assert.Same(repeater, SdtInspector.NearestRepeaterAncestor(item));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void NearestRepeaterAncestor_is_null_for_an_orphaned_repeatingSectionItem()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.OrphanRepeaterItem());
        try
        {
            var item = FirstSdt(path);
            Assert.True(SdtInspector.IsRepeaterItem(item));
            Assert.Null(SdtInspector.NearestRepeaterAncestor(item));
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- generic HasChild / FirstChild / Attr / ChildVal ----

    [Fact]
    public void ChildVal_reads_the_val_attribute_in_the_w_namespace_regardless_of_the_childs_own_namespace()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.ProperRepeater(
            "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line", SomeItemId));
        try
        {
            var repeater = NthSdt(path, 0);
            var pr = repeater.GetFirstChild<SdtProperties>()!;

            // w:id is a w:-namespaced child either way; ChildVal must read its w:val the same regardless.
            Assert.Equal("100", SdtInspector.ChildVal(pr, "id", OoxmlNames.W));
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- ExtractBcNamespace: prefixMappings comes in more shapes than Word's own writer emits ----

    [Theory]
    // The form Word writes: a quoted xmlns declaration, with the trailing space BC also emits.
    [InlineData("xmlns:ns0='urn:microsoft-dynamics-nav/reports/R/1/' ",
                "urn:microsoft-dynamics-nav/reports/R/1/")]
    // Double-quoted (escaped in the attribute) - same shape, other quote style.
    [InlineData("xmlns:ns0=\"urn:microsoft-dynamics-nav/reports/R/1/\"",
                "urn:microsoft-dynamics-nav/reports/R/1/")]
    // BARE uri, no xmlns declaration at all: a real base-app shape, see StandardSalesInvoiceVatSpec.docx.
    [InlineData("urn:microsoft-dynamics-nav/reports/R/1/",
                "urn:microsoft-dynamics-nav/reports/R/1/")]
    // Several prefixes declared; only the BC one is of interest and it is not first.
    [InlineData("xmlns:a='http://example.com/x' xmlns:ns0='urn:microsoft-dynamics-nav/reports/R/1/'",
                "urn:microsoft-dynamics-nav/reports/R/1/")]
    public void ExtractBcNamespace_finds_the_bc_uri_in_every_prefixMappings_shape(string mappings, string expected)
    {
        Assert.Equal(expected, SdtInspector.ExtractBcNamespace(mappings));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("xmlns:a='http://example.com/not-bc'")]
    public void ExtractBcNamespace_returns_null_when_no_bc_namespace_is_declared(string? mappings)
    {
        Assert.Null(SdtInspector.ExtractBcNamespace(mappings));
    }
}
