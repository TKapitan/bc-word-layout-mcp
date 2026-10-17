using System.Text;
using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using DocumentFormat.OpenXml.Packaging;

namespace BcWordLayout.Tests;

/// <summary>
/// GitHub issue #96: a BC dataset part that is there but does not parse used to be reported as "no BC part"
/// by every tool, with the parser's error swallowed, and refresh_xml_part - the one tool that replaces the
/// part - refused to touch it. Reproduced here exactly as the issue did, on copies of JobQuote.docx whose
/// <c>customXml/item1.xml</c> (UTF-16 LE with a BOM, declaring utf-16, in the corpus) is broken six ways.
/// </summary>
public class UnreadableDatasetPartTests
{
    private const string JobQuotePart = "customXml/item1.xml";
    private const string JobQuoteItemId = "{066C1F47-B699-4C82-A6FD-7B5BDFDF4CBA}";
    private const string JobQuoteNamespace = "urn:microsoft-dynamics-nav/reports/Job_Quote/1016/";

    // SalespersonCommission.docx's BC part: UTF-8 with a BOM, and NO itemProps - so a broken copy of it can
    // only be recognised from its raw bytes.
    private const string CommissionPart = "customXml/item2.xml";

    /// <summary>
    /// The six broken shapes from the issue, each with a fragment its error must carry: the parser's own
    /// line/position where the parser reports one, else the encoding contradiction this fix names (the
    /// parser reports "There is no Unicode byte order mark" with no position at all).
    /// </summary>
    public static TheoryData<string, string> BrokenShapes => new()
    {
        { "utf8-bytes-declaring-utf16", "says encoding=\"utf-16\" but its bytes are UTF-8 -" },
        { "utf8-bom-declaring-utf16", "says encoding=\"utf-16\" but its bytes are UTF-8 (with a UTF-8 byte order mark)" },
        { "utf16-bom-declaring-utf8", "Line 2, position 1." },
        { "unescaped-ampersand", "Line 3, position 29." },
        { "truncated", "Unexpected end of file" },
        { "doctype", "The part carries a DOCTYPE declaration" },
    };

    private static byte[] Break(string shape, byte[] corpusPart)
    {
        // The corpus part is UTF-16 LE with a BOM; work on its text, then re-encode per shape.
        var text = Encoding.Unicode.GetString(corpusPart, 2, corpusPart.Length - 2);
        byte[] Utf16WithBom(string s) => [0xFF, 0xFE, .. Encoding.Unicode.GetBytes(s)];

        return shape switch
        {
            "utf8-bytes-declaring-utf16" => new UTF8Encoding(false).GetBytes(text),
            "utf8-bom-declaring-utf16" => [0xEF, 0xBB, 0xBF, .. new UTF8Encoding(false).GetBytes(text)],
            "utf16-bom-declaring-utf8" => Utf16WithBom(text.Replace("encoding=\"utf-16\"", "encoding=\"utf-8\"", StringComparison.Ordinal)),
            "unescaped-ampersand" => Utf16WithBom(text.Replace("<BCReportInformation>", "<BCReportInformation>Tom & Jerry", StringComparison.Ordinal)),
            "truncated" => Utf16WithBom(text[..(text.Length / 2)]),
            "doctype" => Utf16WithBom(text.Replace("?>", "?>\r\n<!DOCTYPE NavWordReportXmlPart>", StringComparison.Ordinal)),
            _ => throw new ArgumentOutOfRangeException(nameof(shape)),
        };
    }

    private static string BrokenJobQuote(string shape)
    {
        var path = PackageSurgery.CopyOfCorpus(Corpus.JobQuote, "unreadable");
        PackageSurgery.ReplaceEntry(path, JobQuotePart, Break(shape, PackageSurgery.ReadEntry(path, JobQuotePart)));
        return path;
    }

    [Theory]
    [MemberData(nameof(BrokenShapes))]
    public void Validate_reports_the_part_and_its_parse_error_instead_of_no_BC_part(string shape, string errorFragment)
    {
        var path = BrokenJobQuote(shape);
        try
        {
            var result = LayoutValidator.Quick(path);

            // One finding naming the cause, not "No BC dataset custom XML part found" beside a schema-parse
            // finding that said the same thing.
            Assert.DoesNotContain(result.Findings, f => f.Check == "single-bc-part");
            var finding = Assert.Single(result.Findings, f => f.Check == "schema-parse");
            Assert.Equal(FindingSeverity.Error, finding.Severity);
            Assert.Equal(JobQuotePart, finding.Location);
            Assert.Contains($"'{JobQuotePart}' (recognised by its itemProps schemaRef, '{JobQuoteNamespace}')", finding.Message);
            Assert.Contains(errorFragment, finding.Message);
            Assert.Contains("run refresh_xml_part with this layout as layoutPath", finding.Message);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [MemberData(nameof(BrokenShapes))]
    public void GetLayoutInfo_fails_invalid_layout_naming_the_part_and_pointing_at_refresh_xml_part(string shape, string errorFragment)
    {
        var path = BrokenJobQuote(shape);
        try
        {
            var response = ReadTools.GetLayoutInfo(path);

            Assert.False(response.Ok);
            Assert.Equal("invalid_layout", response.Error!.Code);
            Assert.Contains($"The BC dataset custom XML part '{JobQuotePart}'", response.Error.Message);
            Assert.Contains(errorFragment, response.Error.Message);
            Assert.Contains("refresh_xml_part repairs it in place", response.Error.Hint);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [MemberData(nameof(BrokenShapes))]
    public void RefreshXmlPart_replaces_the_unreadable_part_keeping_its_storeItemId(string shape, string errorFragment)
    {
        var path = BrokenJobQuote(shape);
        try
        {
            var response = LifecycleTools.RefreshXmlPart(path, Corpus.Path(Corpus.JobQuote));

            Assert.True(response.Ok, response.Error?.Message);
            var data = Assert.IsType<RefreshResultDto>(response.Data);
            Assert.Contains($"'{JobQuotePart}'", data.RepairedUnreadablePart);
            Assert.Contains(errorFragment, data.RepairedUnreadablePart);
            Assert.Equal(JobQuoteItemId, data.StoreItemId);
            Assert.Equal(JobQuoteNamespace, data.OldNamespace); // what its schemaRef named
            Assert.False(data.NamespaceChanged);
            Assert.Empty(data.OrphanedBindings);
            Assert.Empty(data.NewUnboundFields); // no readable old schema to diff against
            Assert.Equal(0, data.QuickValidation.ErrorCount);

            // The part now holds the intact schema, byte for byte, in the SAME part with the SAME itemID.
            Assert.Equal(
                PackageSurgery.ReadEntry(Corpus.Path(Corpus.JobQuote), JobQuotePart),
                PackageSurgery.ReadEntry(path, JobQuotePart));
            using var doc = LayoutPackage.Open(path, isEditable: false);
            var part = Assert.Single(doc.MainDocumentPart!.CustomXmlParts);
            Assert.Equal(JobQuoteItemId, part.CustomXmlPropertiesPart!.DataStoreItem!.ItemId!.Value);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Insert_tools_and_create_layout_report_the_same_cause()
    {
        var path = BrokenJobQuote("utf8-bytes-declaring-utf16");
        var output = Path.Combine(Path.GetTempPath(), $"bcwl-unreadable-out-{Guid.NewGuid():N}.docx");
        try
        {
            var insert = EditTools.InsertField(path, "/Job/No_Job", "documentEnd");
            Assert.Equal("invalid_layout", insert.Error!.Code);
            Assert.Contains($"'{JobQuotePart}'", insert.Error.Message);

            // create_layout's schemaSource is read through the same lookup, so a broken source names its part
            // too instead of claiming it has none.
            var create = LifecycleTools.CreateLayout(path, output);
            Assert.Equal("invalid_layout", create.Error!.Code);
            Assert.Contains($"'{JobQuotePart}'", create.Error.Message);
            Assert.False(File.Exists(output));
        }
        finally
        {
            File.Delete(path);
            File.Delete(output);
        }
    }

    [Fact]
    public void A_part_with_no_itemProps_is_recognised_from_its_raw_bytes_and_repaired_without_minting_one()
    {
        // UTF-8 bytes (BOM and all) that now declare utf-16 - the text-editor re-save shape - on the corpus's
        // no-itemProps layout, so the BC namespace in the raw bytes is the only evidence there is.
        var path = PackageSurgery.CopyOfCorpus(Corpus.SalespersonCommission, "unreadable");
        try
        {
            var original = PackageSurgery.ReadEntry(path, CommissionPart);
            var text = Encoding.UTF8.GetString(original, 3, original.Length - 3);
            PackageSurgery.ReplaceEntry(
                path,
                CommissionPart,
                [0xEF, 0xBB, 0xBF, .. Encoding.UTF8.GetBytes(text.Replace("encoding=\"utf-8\"", "encoding=\"utf-16\"", StringComparison.Ordinal))]);

            var finding = Assert.Single(LayoutValidator.Quick(path).Findings, f => f.Check == "schema-parse");
            Assert.Contains(
                $"'{CommissionPart}' (recognised by the BC report namespace in its raw content, "
                + "'urn:microsoft-dynamics-nav/reports/Salesperson_Commission/115/')",
                finding.Message);

            var response = LifecycleTools.RefreshXmlPart(path, Corpus.Path(Corpus.SalespersonCommission));

            Assert.True(response.Ok, response.Error?.Message);
            var data = Assert.IsType<RefreshResultDto>(response.Data);
            Assert.NotNull(data.RepairedUnreadablePart);
            Assert.Null(data.StoreItemId);
            Assert.Equal(0, data.QuickValidation.ErrorCount);
            Assert.Equal(original, PackageSurgery.ReadEntry(path, CommissionPart));
            using var doc = LayoutPackage.Open(path, isEditable: false);
            Assert.Null(Assert.Single(doc.MainDocumentPart!.CustomXmlParts).CustomXmlPropertiesPart);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("bom-less-utf16")]
    [InlineData("declaration-less-utf8")]
    public void Shapes_whose_bytes_do_not_contradict_a_declaration_still_parse(string shape)
    {
        // The issue's control group: only a declaration that contradicts the bytes fails, so nothing here
        // may start treating these as unreadable.
        var path = PackageSurgery.CopyOfCorpus(Corpus.JobQuote, "unreadable");
        try
        {
            var corpusPart = PackageSurgery.ReadEntry(path, JobQuotePart);
            var text = Encoding.Unicode.GetString(corpusPart, 2, corpusPart.Length - 2);
            var bytes = shape == "bom-less-utf16"
                ? Encoding.Unicode.GetBytes(text)
                : new UTF8Encoding(false).GetBytes(text[(text.IndexOf("?>", StringComparison.Ordinal) + 2)..].TrimStart());
            PackageSurgery.ReplaceEntry(path, JobQuotePart, bytes);

            var response = ReadTools.GetLayoutInfo(path);

            Assert.True(response.Ok, response.Error?.Message);
            Assert.Equal("Job_Quote", Assert.IsType<LayoutInfoDto>(response.Data).Report.ReportName);
            Assert.True(LayoutValidator.Quick(path).Passed);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_broken_part_that_is_not_the_BC_part_is_still_just_not_the_BC_part()
    {
        // Non-tautology guard: recognition must not claim every broken part. One layout keeps its good BC
        // part beside a broken unrelated one (nothing changes); the other has ONLY the broken unrelated part
        // (still "no BC part", exactly as before).
        var withGoodPart = SyntheticLayout.Create(string.Empty);
        var withoutBcPart = SyntheticLayout.CreateWordPackage(DocumentFormat.OpenXml.WordprocessingDocumentType.Document, ".docx");
        try
        {
            foreach (var path in new[] { withGoodPart, withoutBcPart })
            {
                using var doc = WordprocessingDocument.Open(path, isEditable: true);
                var part = doc.MainDocumentPart!.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
                var broken = Encoding.UTF8.GetBytes("<b:Sources xmlns:b=\"http://schemas.openxmlformats.org/officeDocument/2006/bibliography\">&</b:Sources>");
                stream.Write(broken, 0, broken.Length);
            }

            Assert.True(ReadTools.GetLayoutInfo(withGoodPart).Ok);
            Assert.DoesNotContain(LayoutValidator.Quick(withGoodPart).Findings, f => f.Check is "schema-parse" or "single-bc-part");

            var findings = LayoutValidator.Quick(withoutBcPart).Findings;
            Assert.Contains(findings, f => f.Check == "single-bc-part" && f.Message.StartsWith("No BC dataset custom XML part found", StringComparison.Ordinal));
            Assert.Contains(findings, f => f.Check == "schema-parse" && f.Location is null);
        }
        finally
        {
            File.Delete(withGoodPart);
            File.Delete(withoutBcPart);
        }
    }
}
