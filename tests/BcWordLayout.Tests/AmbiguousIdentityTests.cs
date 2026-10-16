using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Ds = DocumentFormat.OpenXml.CustomXmlDataProperties;

namespace BcWordLayout.Tests;

/// <summary>
/// GitHub issue #101: two identities the tools resolved by "first match" instead of refusing (ADR-0003) — a
/// package carrying two BC dataset parts (relationship order picked the schema) and a content-control id
/// carried by two controls (part order picked the control). Reproduced as the issue did: JobQuote.docx plus
/// report 1306's BC part, with its own itemProps, in either relationship order; and JobQuote.docx with
/// header2's control given the id of the body control <c>#Nav: /Job/BillToAddress3</c>.
/// </summary>
public class AmbiguousIdentityTests
{
    private const string JobQuotePart = "customXml/item1.xml";
    private const string ForeignItemId = "{449FD958-388F-4BB5-BF33-B32A63E27B5D}";
    private const string ForeignNamespace = "urn:microsoft-dynamics-nav/reports/Standard_Sales_Invoice/1306/";
    private const int BillToAddress3Id = 189959013;

    // ---- two BC dataset parts ----

    /// <summary>
    /// A copy of JobQuote.docx with StandardSalesInvoiceVatSpec.docx's report-1306 BC part (and an itemProps
    /// carrying that layout's own itemID) added beside JobQuote's own; relationship listed first or last.
    /// </summary>
    private static (string Path, string ForeignPart) JobQuoteWithForeignPart(bool foreignFirst)
    {
        var path = PackageSurgery.CopyOfCorpus(Corpus.JobQuote, "twoparts");
        var foreignBytes = PackageSurgery.ReadEntry(Corpus.Path(Corpus.SalesInvoiceVatSpec), "customXml/item2.xml");

        string foreignPart;
        using (var doc = WordprocessingDocument.Open(path, isEditable: true))
        {
            var part = doc.MainDocumentPart!.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            using (var stream = part.GetStream(FileMode.Create, FileAccess.Write))
            {
                stream.Write(foreignBytes, 0, foreignBytes.Length);
            }

            var props = part.AddNewPart<CustomXmlPropertiesPart>();
            props.DataStoreItem = new Ds.DataStoreItem
            {
                ItemId = ForeignItemId,
                SchemaReferences = new Ds.SchemaReferences(new Ds.SchemaReference { Uri = ForeignNamespace }),
            };
            props.DataStoreItem.Save();
            foreignPart = part.Uri.OriginalString.TrimStart('/');
        }

        if (foreignFirst)
        {
            PackageSurgery.MoveRelationshipFirst(path, Path.GetFileName(foreignPart));
        }

        // Precondition, so the two orders genuinely differ: the SDK enumerates parts in relationship order.
        using (var doc = LayoutPackage.Open(path, isEditable: false))
        {
            var first = SchemaProvider.Scan(doc.MainDocumentPart!).Parts[0].Name;
            Assert.Equal(foreignFirst ? foreignPart : JobQuotePart, first);
        }

        return (path, foreignPart);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Validate_measures_against_the_part_the_bindings_use_in_either_relationship_order(bool foreignFirst)
    {
        var (path, foreignPart) = JobQuoteWithForeignPart(foreignFirst);
        try
        {
            var result = LayoutValidator.Quick(path);

            // One error, the package defect itself - not 124 bogus ones measured against the foreign part.
            var finding = Assert.Single(result.Findings, f => f.Severity == FindingSeverity.Error);
            Assert.Equal("single-bc-part", finding.Check);
            Assert.StartsWith("Found 2 BC dataset custom XML parts; exactly one is expected: ", finding.Message);
            Assert.Contains($"'{JobQuotePart}' (Job_Quote/1016, itemID {{066C1F47-B699-4C82-A6FD-7B5BDFDF4CBA}})", finding.Message);
            Assert.Contains($"'{foreignPart}' (Standard_Sales_Invoice/1306, itemID {ForeignItemId})", finding.Message);
            Assert.Contains(
                $"The other checks measure the layout against '{JobQuotePart}': 41 of the layout's 41 dataset "
                + "binding(s) name its itemID as their storeItemID.",
                finding.Message);
            Assert.Contains("delete its customXml/itemN.xml", finding.Message);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void GetLayoutInfo_reports_the_part_the_bindings_use_in_either_relationship_order(bool foreignFirst)
    {
        var (path, _) = JobQuoteWithForeignPart(foreignFirst);
        try
        {
            var info = Assert.IsType<LayoutInfoDto>(ReadTools.GetLayoutInfo(path).Data);

            Assert.Equal("Job_Quote", info.Report.ReportName);
            Assert.Equal("1016", info.Report.ReportId);
            Assert.Equal("{066C1F47-B699-4C82-A6FD-7B5BDFDF4CBA}", info.Report.StoreItemId);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void RefreshXmlPart_refuses_a_layout_with_two_BC_parts_and_leaves_it_untouched(bool foreignFirst)
    {
        var (path, foreignPart) = JobQuoteWithForeignPart(foreignFirst);
        try
        {
            var before = File.ReadAllBytes(path);

            var response = LifecycleTools.RefreshXmlPart(path, Corpus.Path(Corpus.JobQuote));

            Assert.False(response.Ok);
            Assert.Equal("invalid_layout", response.Error!.Code);
            Assert.StartsWith("This layout carries 2 BC dataset custom XML parts (", response.Error.Message);
            Assert.Contains($"'{foreignPart}'", response.Error.Message);
            Assert.Contains($"The read-only tools report '{JobQuotePart}'", response.Error.Message);
            Assert.Contains("Remove the extra part from the .docx package", response.Error.Hint);
            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Tools_that_write_a_binding_refuse_but_a_row_with_no_bound_cell_does_not()
    {
        var (path, _) = JobQuoteWithForeignPart(foreignFirst: true);
        try
        {
            var field = EditTools.InsertField(path, "/Job/No_Job", "documentEnd");
            Assert.Equal("invalid_layout", field.Error!.Code);
            Assert.Contains("A layout carries exactly one BC dataset part", field.Error.Hint);

            var picture = EditTools.InsertPicture(path, "/Job/CompanyPicture", "documentEnd");
            Assert.Equal("invalid_layout", picture.Error!.Code);

            // An all-spacer row binds nothing, so which dataset part is meant cannot matter to it.
            var info = Assert.IsType<LayoutInfoDto>(ReadTools.GetLayoutInfo(path).Data);
            var table = info.Tables[0];
            var spacer = TableTools.InsertTableRow(path, table.TableIndex, $"{table.ColumnCount}:-");
            Assert.True(spacer.Ok, spacer.Error?.Message);

            var bound = TableTools.InsertTableRow(path, table.TableIndex, $"{table.ColumnCount}:/Job/No_Job");
            Assert.Equal("invalid_layout", bound.Error!.Code);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_two_part_schema_source_is_refused_naming_the_source()
    {
        var (source, _) = JobQuoteWithForeignPart(foreignFirst: false);
        var target = PackageSurgery.CopyOfCorpus(Corpus.JobQuote, "twoparts-target");
        var output = Path.Combine(Path.GetTempPath(), $"bcwl-twoparts-out-{Guid.NewGuid():N}.docx");
        try
        {
            var refresh = LifecycleTools.RefreshXmlPart(target, source);
            Assert.Equal("invalid_layout", refresh.Error!.Code);
            Assert.StartsWith($"newSchemaSource '{source}' carries 2 BC dataset custom XML parts", refresh.Error.Message);

            var create = LifecycleTools.CreateLayout(source, output);
            Assert.Equal("invalid_layout", create.Error!.Code);
            Assert.Contains("carries 2 BC dataset custom XML parts", create.Error.Message);
            Assert.False(File.Exists(output));
        }
        finally
        {
            File.Delete(source);
            File.Delete(target);
            File.Delete(output);
        }
    }

    // ---- the choice itself: storeItemID votes, then namespace votes, then package path ----

    private const string ReportA = "urn:microsoft-dynamics-nav/reports/ReportA/50000/";
    private const string ReportB = "urn:microsoft-dynamics-nav/reports/ReportB/50001/";
    private const string ItemA = "{AAAAAAAA-0000-0000-0000-000000000001}";
    private const string ItemB = "{BBBBBBBB-0000-0000-0000-000000000002}";
    private const string CompanyNameXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:CompanyName[1]";

    private static string Dataset(string ns) =>
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        + $"<NavWordReportXmlPart xmlns=\"{ns}\"><Header><CompanyName>Contoso</CompanyName></Header></NavWordReportXmlPart>";

    /// <summary>A synthetic layout with <paramref name="body"/> and one BC part per entry, in that relationship order.</summary>
    private static string LayoutWithParts(string body, params (string Namespace, string? ItemId)[] parts)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-choice-{Guid.NewGuid():N}.docx");
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        WriteUtf8(
            main.GetStream(FileMode.Create, FileAccess.Write),
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
            + $"<w:body>{body}<w:sectPr/></w:body></w:document>");

        foreach (var (ns, itemId) in parts)
        {
            var part = main.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            WriteUtf8(part.GetStream(FileMode.Create, FileAccess.Write), Dataset(ns));

            if (itemId is not null)
            {
                var props = part.AddNewPart<CustomXmlPropertiesPart>();
                props.DataStoreItem = new Ds.DataStoreItem { ItemId = itemId, SchemaReferences = new Ds.SchemaReferences() };
                props.DataStoreItem.Save();
            }
        }

        return path;
    }

    private static void WriteUtf8(Stream stream, string content)
    {
        using (stream)
        using (var writer = new StreamWriter(stream, new System.Text.UTF8Encoding(false)))
        {
            writer.Write(content);
        }
    }

    private static string Field(string storeItemId, string? ns = null) =>
        ns is null
            ? SyntheticLayout.BoundField(CompanyNameXPath, storeItemId)
            : SyntheticLayout.BoundFieldInNamespace(CompanyNameXPath, storeItemId, ns);

    private static (string Chosen, string? Reason, string ReportName) Choice(string path)
    {
        using var doc = LayoutPackage.Open(path, isEditable: false);
        var scan = SchemaProvider.Scan(doc.MainDocumentPart!);
        return (scan.Chosen!.Namespace, scan.ChoiceReason, SchemaProvider.FromLayout(doc).Report.ReportName);
    }

    [Fact]
    public void The_part_most_bindings_name_by_storeItemID_wins_even_listed_last()
    {
        var path = LayoutWithParts(Field(ItemB) + Field(ItemB) + Field(ItemA), (ReportA, ItemA), (ReportB, ItemB));
        try
        {
            var (chosen, reason, reportName) = Choice(path);

            Assert.Equal(ReportB, chosen);
            Assert.Equal("ReportB", reportName);
            Assert.Equal("2 of the layout's 3 dataset binding(s) name its itemID as their storeItemID", reason);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void With_no_itemProps_to_vote_on_the_part_the_bindings_namespace_names_wins()
    {
        var path = LayoutWithParts(Field(ItemA, ReportB), (ReportA, null), (ReportB, null));
        try
        {
            var (chosen, reason, _) = Choice(path);

            Assert.Equal(ReportB, chosen);
            Assert.Equal(
                "no part's itemID is named by more bindings than another's, and 1 of the layout's 1 dataset "
                + "binding(s) name its namespace",
                reason);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void With_no_evidence_the_lowest_package_path_wins_whatever_the_relationship_order()
    {
        var path = LayoutWithParts(string.Empty, (ReportA, ItemA), (ReportB, ItemB));
        try
        {
            var (inOrder, reason, _) = Choice(path);
            PackageSurgery.MoveRelationshipFirst(path, "item2.xml");
            var (reordered, _, _) = Choice(path);

            Assert.Equal(ReportA, inOrder);
            Assert.Equal(ReportA, reordered);
            Assert.Equal(
                "no binding singles out one part, by storeItemID or by namespace, so the part whose package path "
                + "sorts first is used",
                reason);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_template_whose_BC_part_cannot_be_parsed_has_it_replaced_not_kept_beside_the_new_one()
    {
        // create_layout removes every BC part a templatePath carries before attaching its own; an unreadable
        // one is a BC part too, and left behind it would be the second BC part of the new layout.
        var template = SyntheticLayout.CreateWordPackage(WordprocessingDocumentType.Document, ".docx");
        var output = Path.Combine(Path.GetTempPath(), $"bcwl-template-out-{Guid.NewGuid():N}.docx");
        try
        {
            using (var doc = WordprocessingDocument.Open(template, isEditable: true))
            {
                var part = doc.MainDocumentPart!.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
                var bytes = System.Text.Encoding.UTF8.GetBytes(Dataset(ReportA).Replace("UTF-8", "utf-16", StringComparison.Ordinal));
                stream.Write(bytes, 0, bytes.Length);
            }

            var response = LifecycleTools.CreateLayout(Corpus.Path(Corpus.JobQuote), output, template);

            Assert.True(response.Ok, response.Error?.Message);
            Assert.True(Assert.IsType<CreateResultDto>(response.Data).ReplacedExistingBcPart);
            using var built = LayoutPackage.Open(output, isEditable: false);
            var scan = SchemaProvider.Scan(built.MainDocumentPart!);
            Assert.Equal(1, scan.Count);
            Assert.Equal("urn:microsoft-dynamics-nav/reports/Job_Quote/1016/", Assert.Single(scan.Parts).Namespace);
        }
        finally
        {
            File.Delete(template);
            File.Delete(output);
        }
    }

    // ---- duplicate content-control ids ----

    /// <summary>A copy of <paramref name="corpusFile"/> in which the first control of <paramref name="partName"/> is given <paramref name="id"/>.</summary>
    private static string WithControlIdCopied(string corpusFile, string partName, int id)
    {
        var path = PackageSurgery.CopyOfCorpus(corpusFile, "dupid");
        using var doc = WordprocessingDocument.Open(path, isEditable: true);
        var main = doc.MainDocumentPart!;
        OpenXmlPartRootElement root = partName == "document.xml"
            ? main.Document!
            : main.HeaderParts.Single(h => Path.GetFileName(h.Uri.OriginalString) == partName).Header!;
        var control = root.Descendants<SdtElement>().First(s => s.SdtProperties?.GetFirstChild<SdtId>()?.Val?.Value != id);
        control.SdtProperties!.GetFirstChild<SdtId>()!.Val = id;
        root.Save();
        return path;
    }

    [Fact]
    public void Validate_warns_about_an_id_two_controls_share()
    {
        var path = WithControlIdCopied(Corpus.JobQuote, "header2.xml", BillToAddress3Id);
        try
        {
            var result = LayoutValidator.Quick(path);

            var finding = Assert.Single(result.Findings, f => f.Check == "duplicate-control-id");
            Assert.Equal(FindingSeverity.Warning, finding.Severity);
            Assert.Contains(
                "Content-control id 189959013 is used by 2 controls (document.xml: '#Nav: /Job/BillToAddress3', "
                + "header2.xml: '#Nav: /Job/TodayFormatted')",
                finding.Message);
            Assert.Equal("document.xml, header2.xml", finding.Location);
            Assert.Equal(0, result.ErrorCount);

            // Non-tautology guard: the untouched corpus layout draws no such finding.
            Assert.DoesNotContain(LayoutValidator.Quick(Corpus.Path(Corpus.JobQuote)).Findings, f => f.Check == "duplicate-control-id");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveControl_refuses_a_shared_id_instead_of_removing_the_first_match()
    {
        var path = WithControlIdCopied(Corpus.JobQuote, "header2.xml", BillToAddress3Id);
        try
        {
            var before = File.ReadAllBytes(path);

            var response = EditTools.RemoveControl(path, BillToAddress3Id);

            Assert.False(response.Ok);
            Assert.Equal("invalid_layout", response.Error!.Code);
            Assert.StartsWith("Control id 189959013 is carried by 2 content controls in this layout (", response.Error.Message);
            Assert.Contains("No tool here renumbers a control", response.Error.Hint);
            Assert.Equal(before, File.ReadAllBytes(path));

            // A unique id still removes normally.
            var info = Assert.IsType<LayoutInfoDto>(ReadTools.GetLayoutInfo(path).Data);
            var other = info.Controls.First(c => c.SdtId is { } id && id != BillToAddress3Id).SdtId!.Value;
            Assert.True(EditTools.RemoveControl(path, other).Ok);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("body", null)]
    [InlineData("header", "header2.xml")]
    public void AfterControl_refuses_a_shared_id_whichever_part_it_targets(string layoutPart, string? partName)
    {
        var path = WithControlIdCopied(Corpus.JobQuote, "header2.xml", BillToAddress3Id);
        try
        {
            var response = EditTools.InsertText(path, "x", "afterControl", controlId: BillToAddress3Id, layoutPart: layoutPart, partName: partName);

            Assert.Equal("invalid_layout", response.Error!.Code);
            Assert.Contains("is carried by 2 content controls", response.Error.Message);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_shared_parent_repeater_id_is_refused()
    {
        var info = Assert.IsType<LayoutInfoDto>(ReadTools.GetLayoutInfo(Corpus.Path(Corpus.SalesInvoice)).Data);
        var repeater = info.Controls.First(c => c.Kind == "Repeater" && c.SdtId.HasValue);
        var path = WithControlIdCopied(Corpus.SalesInvoice, "document.xml", repeater.SdtId!.Value);
        try
        {
            // Refused before the row is even laid out, so the cell specs never get looked at.
            var response = TableTools.InsertSubtotalRow(path, repeater.SdtId.Value, "-");

            Assert.Equal("invalid_layout", response.Error!.Code);
            Assert.Contains($"Control id {repeater.SdtId.Value} is carried by 2 content controls", response.Error.Message);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
