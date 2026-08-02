using BcWordLayout.Domain;

namespace BcWordLayout.Tests;

/// <remarks>Joins the label-convention-seam collection: <c>InventoryOrderDetails_identity_and_zero_labels</c>
/// pins <c>InventoryOrderDetails.docx</c>'s zero-<c>IsLabel</c>-column count, which a concurrently-swapped
/// <c>LabelConvention.Current</c> could disturb (see <see cref="LabelConventionSeamCollection"/>).</remarks>
[Collection("label-convention-seam")]
public class SchemaProviderTests
{
    [Fact]
    public void SalesInvoice_identity_and_labels()
    {
        var tree = SchemaProvider.FromLayout(Corpus.Path(Corpus.SalesInvoice));

        Assert.Equal("1306", tree.Report.ReportId);
        Assert.Contains("Standard_Sales_Invoice", tree.Report.Namespace);
        Assert.Equal("Standard_Sales_Invoice", tree.Report.ReportName);
        Assert.Equal("{AF7A6226-6056-400F-ADDA-E1ADA7C08250}", tree.Report.StoreItemId);

        // At least 100 label-suffixed columns somewhere in the schema tree (actual: 112).
        var labelColumns = tree.AllColumns().Count(c => c.IsLabel);
        Assert.True(labelColumns >= 100, $"expected >= 100 label columns, found {labelColumns}");
    }

    [Fact]
    public void InventoryOrderDetails_identity_and_zero_labels()
    {
        // This layout's label-like columns are suffixed "Caption"/"Label" (a BC label naming convention
        // LabelConvention.IsLabelName - which only recognizes the "*Lbl" suffix - does not recognize), so
        // every one of them is classified as a plain (non-label) column here; see the final report's note
        // on this production-code assumption.
        var tree = SchemaProvider.FromLayout(Corpus.Path(Corpus.InventoryOrderDetails));

        Assert.Equal("50002", tree.Report.ReportId);
        Assert.Contains("FS_YSR_InventoryOrderDetails", tree.Report.Namespace);
        Assert.Equal(0, tree.AllColumns().Count(c => c.IsLabel));
    }

    [Fact]
    public void StandardStatement_identity()
    {
        var tree = SchemaProvider.FromLayout(Corpus.Path(Corpus.StandardStatement));

        Assert.Equal("1316", tree.Report.ReportId);
        Assert.Contains("Standard_Statement", tree.Report.Namespace);
    }

    [Fact]
    public void Root_is_NavWordReportXmlPart_and_system_node_is_flagged()
    {
        var tree = SchemaProvider.FromLayout(Corpus.Path(Corpus.SalesInvoice));

        Assert.Equal("NavWordReportXmlPart", tree.Root.Name);

        // BCReportInformation must exist as a system node and be excluded from business columns.
        var system = tree.Root.Children.SingleOrDefault(c => c.Name == "BCReportInformation");
        Assert.NotNull(system);
        Assert.True(system!.IsSystem);
        Assert.DoesNotContain(tree.AllDataItems(), d => d.IsSystem);
    }

    // ---- FromSchemaXml: root LOCAL NAME alone is not enough, the NAMESPACE must be a BC one too ----

    [Fact]
    public void FromSchemaXml_with_correct_root_name_but_a_non_BC_namespace_throws_InvalidDataException()
    {
        // Right local name (NavWordReportXmlPart), WRONG namespace (not "urn:microsoft-dynamics-nav/reports/...").
        // Distinct from the existing "wrong root element name" negative path (LayoutBuilderTests'
        // Create_from_a_non_BC_xml_throws_InvalidDataException uses <SomeOtherRoot>, which already fails the
        // EARLIER local-name check and never reaches this one) - this proves the namespace check itself.
        var badXmlPath = Path.Combine(Path.GetTempPath(), $"bcwl-schemaprovider-badns-{Guid.NewGuid():N}.xml");
        try
        {
            File.WriteAllText(
                badXmlPath,
                "<NavWordReportXmlPart xmlns=\"urn:not-bc\"><Header><Foo>x</Foo></Header></NavWordReportXmlPart>");

            var ex = Assert.Throws<InvalidDataException>(() => SchemaProvider.FromSchemaXml(badXmlPath));
            Assert.Contains("urn:not-bc", ex.Message, StringComparison.Ordinal);
        }
        finally
        {
            if (File.Exists(badXmlPath))
            {
                File.Delete(badXmlPath);
            }
        }
    }
}
