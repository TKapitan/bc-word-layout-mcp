using System.Reflection;
using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

public class SdtFactoryTests
{
    // Real, verified paths from tests/corpus/SalesInvoiceForSubscriptionBilling.docx (Standard_Sales_Invoice/1306):
    // a top-level field, a top-level label, a nested field/label under the Line repeater, and the Line
    // data item itself (for the "must be a leaf column" rejection test).
    private const string FieldPath = "/Header/CustomerAddress1";
    private const string LabelPath = "/Header/Contact_Lbl";
    private const string NestedFieldPath = "/Header/Line/ItemNo_Line";
    private const string NestedLabelPath = "/Header/Line/ItemNo_Line_Lbl";
    private const string DataItemPath = "/Header/Line";
    private const string NonExistentPath = "/Header/ThisFieldDoesNotExistAnywhere";

    private static DatasetTree SalesInvoiceSchema() => SchemaProvider.FromLayout(Corpus.Path(Corpus.SalesInvoice));

    private static SdtAlias Alias(SdtRun sdt) => sdt.SdtProperties!.GetFirstChild<SdtAlias>()!;

    private static Tag Tag(SdtRun sdt) => sdt.SdtProperties!.GetFirstChild<Tag>()!;

    private static SdtId Id(SdtRun sdt) => sdt.SdtProperties!.GetFirstChild<SdtId>()!;

    private static DataBinding Binding(SdtRun sdt) => sdt.SdtProperties!.GetFirstChild<DataBinding>()!;

    private static string VisibleText(SdtRun sdt) =>
        sdt.SdtContentRun!.GetFirstChild<Run>()!.GetFirstChild<Text>()!.Text;

    // ---- shape of a built field ----

    [Fact]
    public void BuildField_valid_path_has_expected_alias_tag_xpath_storeItemId_and_prefixMappings()
    {
        var schema = SalesInvoiceSchema();

        var sdt = SdtFactory.BuildField(schema, FieldPath);

        Assert.Equal("#Nav: /Header/CustomerAddress1", Alias(sdt).Val!.Value);
        Assert.Equal("#Nav: Standard_Sales_Invoice/1306", Tag(sdt).Val!.Value);
        Assert.Equal(
            "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:CustomerAddress1[1]",
            Binding(sdt).XPath!.Value);
        Assert.Equal(schema.Report.StoreItemId, Binding(sdt).StoreItemId!.Value);
        Assert.Equal($"xmlns:ns0='{schema.Report.Namespace}'", Binding(sdt).PrefixMappings!.Value);

        // Plain-text marker present; sdtContent holds a run with placeholder text.
        Assert.NotNull(sdt.SdtProperties!.GetFirstChild<SdtContentText>());
        Assert.Equal("CustomerAddress1", VisibleText(sdt));
    }

    // OpenXmlValidator does NOT enforce CT_SdtPr child order (verified: a deliberately-reordered sdtPr
    // still validates clean), so BC byte-compatibility depends entirely on this factory's hardcoded
    // construction order matching the real corpus - nothing else would catch a reordering regression.
    [Fact]
    public void BuildField_sdtPr_child_order_matches_the_real_corpus_order()
    {
        var schema = SalesInvoiceSchema();

        var sdt = SdtFactory.BuildField(schema, FieldPath);

        var actualOrder = sdt.SdtProperties!.ChildElements.Select(e => e.LocalName).ToList();
        Assert.Equal(new[] { "alias", "tag", "id", "placeholder", "dataBinding", "text" }, actualOrder);
    }

    [Fact]
    public void BuildLabel_sdtPr_child_order_matches_the_real_corpus_order()
    {
        var schema = SalesInvoiceSchema();

        var sdt = SdtFactory.BuildLabel(schema, LabelPath);

        var actualOrder = sdt.SdtProperties!.ChildElements.Select(e => e.LocalName).ToList();
        Assert.Equal(new[] { "alias", "tag", "id", "placeholder", "dataBinding", "text" }, actualOrder);
    }

    [Fact]
    public void BuildField_computes_fully_indexed_xpath_for_a_nested_path()
    {
        var schema = SalesInvoiceSchema();

        var sdt = SdtFactory.BuildField(schema, NestedFieldPath);

        Assert.Equal(
            "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:ItemNo_Line[1]",
            Binding(sdt).XPath!.Value);
        Assert.Equal("#Nav: /Header/Line/ItemNo_Line", Alias(sdt).Val!.Value);
    }

    [Fact]
    public void BuildLabel_valid_label_path_has_expected_shape()
    {
        var schema = SalesInvoiceSchema();

        var sdt = SdtFactory.BuildLabel(schema, LabelPath);

        Assert.Equal("#Nav: /Header/Contact_Lbl", Alias(sdt).Val!.Value);
        Assert.Equal(
            "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Contact_Lbl[1]",
            Binding(sdt).XPath!.Value);
        Assert.Equal(schema.Report.StoreItemId, Binding(sdt).StoreItemId!.Value);
        Assert.Equal("Contact_Lbl", VisibleText(sdt));
    }

    [Fact]
    public void BuildLabel_nested_label_path_resolves_under_the_repeater_item()
    {
        var schema = SalesInvoiceSchema();

        var sdt = SdtFactory.BuildLabel(schema, NestedLabelPath);

        Assert.Equal(
            "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:ItemNo_Line_Lbl[1]",
            Binding(sdt).XPath!.Value);
    }

    // ---- placeholder text / id generation ----

    [Fact]
    public void BuildField_default_placeholder_text_is_the_leaf_segment_name()
    {
        var schema = SalesInvoiceSchema();

        var sdt = SdtFactory.BuildField(schema, FieldPath);

        Assert.Equal("CustomerAddress1", VisibleText(sdt));
    }

    [Fact]
    public void BuildField_custom_placeholder_text_is_used_verbatim()
    {
        var schema = SalesInvoiceSchema();

        var sdt = SdtFactory.BuildField(schema, FieldPath, placeholderText: "Enter address");

        Assert.Equal("Enter address", VisibleText(sdt));
    }

    [Fact]
    public void BuildField_explicit_id_is_used_verbatim()
    {
        var schema = SalesInvoiceSchema();

        var sdt = SdtFactory.BuildField(schema, FieldPath, id: 424242);

        Assert.Equal(424242, Id(sdt).Val!.Value);
    }

    [Fact]
    public void BuildField_omitted_id_generates_an_id_on_every_call()
    {
        // SdtFactory no longer tracks issued ids at all (see SdtFactory.ResolveId's remarks) - it
        // does not (and is not meant to) GUARANTEE distinctness across calls; doc-scoped uniqueness is the
        // caller's job (LayoutEditor.GenerateUniqueId/MakeIdGenerator, which every production call site
        // already uses - they check the real target document, which this factory has no access to). This
        // only proves the omitted-id path still assigns SOME id, on every call.
        var schema = SalesInvoiceSchema();

        var ids = Enumerable.Range(0, 5)
            .Select(_ => Id(SdtFactory.BuildField(schema, FieldPath)).Val)
            .ToList();

        Assert.All(ids, id => Assert.NotNull(id));
    }

    [Fact]
    public void SdtFactory_keeps_no_static_collection_state_that_could_grow_unboundedly()
    {
        // Earlier revisions kept a process-lifetime static HashSet<int> (plus a lock and a shared
        // Random) recording every w:id ever issued OR explicitly passed - unbounded growth over a long
        // host session, for a uniqueness scope (a whole process) w:id was never required to have. Lock
        // the fix down structurally rather than by field name: no static field of a mutable collection
        // type may exist on this type at all, so the anti-pattern cannot silently return under a
        // different name. (A stateless, immutable helper like a compiled Regex is fine and untouched by
        // this check - ICollection is specifically about accumulating/growable state.)
        var offendingFields = typeof(SdtFactory)
            .GetFields(BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Static)
            .Where(f => !f.IsLiteral && typeof(System.Collections.ICollection).IsAssignableFrom(f.FieldType))
            .Select(f => f.Name)
            .ToList();

        Assert.Empty(offendingFields);
    }

    // ---- naming-convention guard rails ----

    [Fact]
    public void BuildField_rejects_a_label_shaped_path()
    {
        var schema = SalesInvoiceSchema();

        var ex = Assert.Throws<ArgumentException>(() => SdtFactory.BuildField(schema, LabelPath));
        // Message must guide the caller toward binding it as a label, without leaking the internal method name.
        Assert.Contains("label", ex.Message);
        Assert.Contains("insert_label", ex.Message);
        Assert.DoesNotContain("BuildLabel", ex.Message);
    }

    [Fact]
    public void BuildLabel_rejects_a_non_label_shaped_path()
    {
        var schema = SalesInvoiceSchema();

        var ex = Assert.Throws<ArgumentException>(() => SdtFactory.BuildLabel(schema, FieldPath));
        // Message must guide the caller toward binding it as a field, without leaking the internal method name.
        Assert.Contains("field", ex.Message);
        Assert.Contains("insert_field", ex.Message);
        Assert.DoesNotContain("BuildField", ex.Message);
    }

    // ---- negative: paths that do not resolve, or resolve to the wrong kind of node ----

    [Fact]
    public void BuildField_rejects_a_nonexistent_dataset_path()
    {
        var schema = SalesInvoiceSchema();

        var ex = Assert.Throws<ArgumentException>(() => SdtFactory.BuildField(schema, NonExistentPath));
        Assert.Contains("does not resolve", ex.Message);
    }

    [Fact]
    public void BuildField_rejects_a_path_that_resolves_to_a_repeating_data_item_not_a_leaf_column()
    {
        var schema = SalesInvoiceSchema();

        var ex = Assert.Throws<ArgumentException>(() => SdtFactory.BuildField(schema, DataItemPath));
        Assert.Contains("repeating data item", ex.Message);
    }

    [Fact]
    public void BuildField_rejects_when_schema_report_has_no_storeItemId()
    {
        var schemaWithoutStoreItemId = new DatasetTree
        {
            Report = new ReportIdentity
            {
                ReportName = "Fake_Report",
                ReportId = "99999",
                Namespace = "urn:microsoft-dynamics-nav/reports/Fake_Report/99999/",
                StoreItemId = null,
            },
            Root = BuildFakeRoot(),
        };

        var ex = Assert.Throws<ArgumentException>(() => SdtFactory.BuildField(schemaWithoutStoreItemId, "/Header/SomeField"));
        Assert.Contains("StoreItemId", ex.Message);
    }

    private static DataItem BuildFakeRoot()
    {
        var root = new DataItem { Name = "NavWordReportXmlPart", Path = "/" };
        var header = new DataItem { Name = "Header", Path = "/Header" };
        header.Columns.Add(new DatasetColumn { Name = "SomeField", Path = "/Header/SomeField" });
        root.Children.Add(header);
        return root;
    }

    // ---- round-trip correctness: a factory-built control survives real OOXML/BC validation ----

    [Fact]
    public void BuildField_inserted_into_a_corpus_copy_passes_OpenXmlValidator_and_LayoutValidator_Quick()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-sdtfactory-field-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(Corpus.SalesInvoice), path, overwrite: true);
        string? expectedStoreItemId;

        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var schema = SchemaProvider.FromLayout(doc);
                expectedStoreItemId = schema.Report.StoreItemId;
                var field = SdtFactory.BuildField(schema, FieldPath, placeholderText: "New Field Control");

                var anchor = LocationResolver.Resolve(new Location { Type = LocationKind.DocumentEnd }, doc);
                anchor.InsertInline(field);

                doc.MainDocumentPart!.Document!.Save();
            }

            using (var reopened = WordprocessingDocument.Open(path, false))
            {
                var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened).ToList();
                Assert.Empty(openXmlErrors);

                var quick = LayoutValidator.Quick(reopened);
                Assert.True(quick.Passed, "errors: " + string.Join(" | ", quick.Findings
                    .Where(f => f.Severity == FindingSeverity.Error).Select(f => f.Message)));
                Assert.Equal(0, quick.ErrorCount);

                // The corpus already binds this same field elsewhere (header table), so there are now
                // two controls sharing this alias/xpath — assert our freshly-inserted one reads back
                // correctly as a Field with the expected xpath/storeItemID, and that inserting a second
                // control on the same field didn't somehow break either existing binding (still exactly
                // two, none unbound/misclassified).
                var inventory = LayoutReader.Read(reopened);
                var matching = inventory.Controls
                    .Where(c => c.XPath == "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:CustomerAddress1[1]")
                    .ToList();
                Assert.Equal(2, matching.Count);
                Assert.All(matching, c =>
                {
                    Assert.Equal(ControlKind.Field, c.Kind);
                    Assert.Equal(expectedStoreItemId, c.StoreItemId);
                });
            }
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void BuildLabel_inserted_into_a_corpus_copy_passes_OpenXmlValidator_and_LayoutValidator_Quick()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-sdtfactory-label-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(Corpus.SalesInvoice), path, overwrite: true);

        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var schema = SchemaProvider.FromLayout(doc);
                var label = SdtFactory.BuildLabel(schema, LabelPath);

                var anchor = LocationResolver.Resolve(new Location { Type = LocationKind.DocumentEnd }, doc);
                anchor.InsertInline(label);

                doc.MainDocumentPart!.Document!.Save();
            }

            using (var reopened = WordprocessingDocument.Open(path, false))
            {
                var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(reopened).ToList();
                Assert.Empty(openXmlErrors);

                var quick = LayoutValidator.Quick(reopened);
                Assert.Equal(0, quick.ErrorCount);

                var inventory = LayoutReader.Read(reopened);
                Assert.Contains(inventory.Controls, c =>
                    c.Kind == ControlKind.Label &&
                    c.XPath == "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Contact_Lbl[1]");
            }
        }
        finally
        {
            File.Delete(path);
        }
    }
}
