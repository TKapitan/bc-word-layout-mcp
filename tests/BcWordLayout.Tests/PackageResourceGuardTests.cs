using System.IO.Compression;
using System.Text;
using BcWordLayout.Domain;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using BcWordLayout.Merge;

namespace BcWordLayout.Tests;

/// <summary>
/// Covers <see cref="ResourceLimits.EnsurePackageWithinLimits"/> behind <see cref="LayoutPackage.Open"/>: the
/// caps that keep the OpenXml SDK's OWN recursive part loads (and <c>OpenXmlValidator</c>'s package walk)
/// from overflowing the stack and killing the server. The deep-nesting fixtures here go well past the depth
/// the SDK survives (measured: about 2,250 nested elements), so a regression does not fail one assertion —
/// it crashes the test host, which is the loudest failure available for an uncatchable
/// <see cref="StackOverflowException"/>.
/// </summary>
public class PackageResourceGuardTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>Content-control nesting well past what the SDK's recursive load survives.</summary>
    private const int CrashingSdtDepth = 3_000;

    private static string NestedSdtChain(int depth)
    {
        var open = new StringBuilder();
        var close = new StringBuilder();
        for (var i = 0; i < depth; i++)
        {
            open.Append("<w:sdt><w:sdtContent>");
            close.Append("</w:sdtContent></w:sdt>");
        }

        return $"{open}<w:p><w:r><w:t>x</w:t></w:r></w:p>{close}";
    }

    private static void AssertDepthRejection(ToolResponse response, string partName)
    {
        Assert.False(response.Ok);
        Assert.NotNull(response.Error);
        Assert.Equal("invalid_layout", response.Error!.Code);
        Assert.Contains($"XML part '/word/{partName}' nesting exceeds the supported depth", response.Error.Message, StringComparison.Ordinal);
        Assert.Contains(ResourceLimits.MaxPackageXmlDepth.ToString(), response.Error.Message, StringComparison.Ordinal);
        Assert.Contains("supported size/nesting limits", response.Error.Hint, StringComparison.Ordinal);
    }

    [Fact]
    public void Read_validate_edit_and_preview_all_reject_a_document_too_deep_for_the_SDK_instead_of_crashing()
    {
        // The reviewer's repro shape: a tiny file (about 2 KB zipped) whose document.xml nests deeper than
        // the SDK's recursive load survives. Every entry point opens it through LayoutPackage.Open.
        var layoutPath = SyntheticLayout.Create(NestedSdtChain(CrashingSdtDepth));
        var originalBytes = File.ReadAllBytes(layoutPath);
        var outputDir = Directory.CreateTempSubdirectory("bcwl-deep-sdk-").FullName;

        try
        {
            Assert.True(originalBytes.Length < 16 * 1024, "the fixture should stay a small file");

            AssertDepthRejection(ReadTools.GetLayoutInfo(layoutPath), "document.xml");
            AssertDepthRejection(ReadTools.ValidateLayout(layoutPath), "document.xml");
            AssertDepthRejection(ReadTools.ValidateLayout(layoutPath, level: "full"), "document.xml");
            AssertDepthRejection(EditTools.InsertText(layoutPath, "x", "documentEnd"), "document.xml");

            // A writable open goes through the read-only probe first, and a rejected edit leaves the file alone.
            Assert.Equal(originalBytes, File.ReadAllBytes(layoutPath));

            var ex = Assert.Throws<ResourceLimitExceededException>(
                () => MergeEngine.Merge(layoutPath, Path.Combine(outputDir, "merged.docx"), new MergeOptions()));
            Assert.Contains("nesting exceeds the supported depth", ex.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(layoutPath);
            Directory.Delete(outputDir, recursive: true);
        }
    }

    [Fact]
    public void A_header_part_too_deep_for_the_SDK_is_rejected_too()
    {
        // Not only document.xml: every part the SDK might parse is scanned.
        var layoutPath = SyntheticLayout.CreateWithHeader(
            SyntheticLayout.PlainParagraph("body"), NestedSdtChain(CrashingSdtDepth));

        try
        {
            var response = ReadTools.GetLayoutInfo(layoutPath);

            Assert.False(response.Ok);
            Assert.Equal("invalid_layout", response.Error!.Code);
            Assert.Contains("nesting exceeds the supported depth", response.Error.Message, StringComparison.Ordinal);
            Assert.Contains("header", response.Error.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            File.Delete(layoutPath);
        }
    }

    [Fact]
    public void Nesting_up_to_the_cap_still_opens()
    {
        // The cap must not reject a part that is deep but within it: a body nested to just under
        // MaxPackageXmlDepth still opens (the element walkers' own cap of 128 content controls is what
        // fails it, with that walker's own message).
        var sdtPairs = (ResourceLimits.MaxPackageXmlDepth - 8) / 2;
        var layoutPath = SyntheticLayout.Create(NestedSdtChain(sdtPairs));

        try
        {
            var response = ReadTools.GetLayoutInfo(layoutPath);

            Assert.False(response.Ok);
            Assert.Contains(ResourceLimits.MaxElementNestingDepth.ToString(), response.Error!.Message, StringComparison.Ordinal);
            Assert.DoesNotContain("XML part", response.Error.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(layoutPath);
        }
    }

    [Fact]
    public void A_part_chain_longer_than_the_part_cap_is_rejected_before_the_validator_walks_it()
    {
        // OpenXmlValidator's package walk recurses once per chained part (measured: it overflows at about
        // 1,500). The part-count cap bounds every such walk, because none can nest deeper than the number
        // of parts.
        var layoutPath = RawDocx.WithPartChain(ResourceLimits.MaxPackageParts + 10);

        try
        {
            var response = ReadTools.ValidateLayout(layoutPath);

            Assert.False(response.Ok);
            Assert.Equal("invalid_layout", response.Error!.Code);
            Assert.Contains($"more than {ResourceLimits.MaxPackageParts} parts", response.Error.Message, StringComparison.Ordinal);
            Assert.Contains("supported size/nesting limits", response.Error.Hint, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(layoutPath);
        }
    }

    [Fact]
    public void An_oversized_document_part_is_rejected_before_the_SDK_loads_it()
    {
        // The per-part byte cap now covers the SDK-loaded parts too, not only the custom XML parts this repo
        // reads itself.
        var filler = new string('A', (int)ResourceLimits.MaxPackageXmlPartBytes + 1024);
        var layoutPath = SyntheticLayout.Create($"<w:p><w:r><w:t>{filler}</w:t></w:r></w:p>");

        try
        {
            var response = ReadTools.GetLayoutInfo(layoutPath);

            Assert.False(response.Ok);
            Assert.Equal("invalid_layout", response.Error!.Code);
            Assert.Contains("'/word/document.xml' exceeds the supported size limit", response.Error.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(layoutPath);
        }
    }

    [Fact]
    public void XML_parts_whose_total_exceeds_the_aggregate_cap_are_rejected()
    {
        // Each part is under the per-part cap; together they are over the package-wide total that bounds
        // what OpenXmlValidator (which parses every part) would hold.
        var perPart = (int)(ResourceLimits.MaxPackageXmlPartBytes - 1024);
        var partCount = (int)(ResourceLimits.MaxPackageXmlBytes / perPart) + 1;
        var layoutPath = RawDocx.WithExtraXmlParts(partCount, perPart);

        try
        {
            var response = ReadTools.GetLayoutInfo(layoutPath);

            Assert.False(response.Ok);
            Assert.Equal("invalid_layout", response.Error!.Code);
            Assert.Contains("Package XML parts total more than", response.Error.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(layoutPath);
        }
    }

    [Fact]
    public void A_DTD_bearing_deep_part_cannot_slip_past_the_scan()
    {
        // The scan ignores (rather than prohibits) DTDs, so a DOCTYPE cannot end the scan early and leave the
        // deep content behind it unchecked. (The SDK itself prohibits DTDs, so it would never parse this part
        // either; the point is that the scan does not rely on that.)
        var deep = NestedSdtChain(CrashingSdtDepth);
        var layoutPath = RawDocx.WithDocumentXml(
            $"<?xml version=\"1.0\"?><!DOCTYPE w:document><w:document xmlns:w=\"{W}\"><w:body>{deep}</w:body></w:document>");

        try
        {
            var response = ReadTools.GetLayoutInfo(layoutPath);

            Assert.False(response.Ok);
            Assert.Equal("invalid_layout", response.Error!.Code);
            Assert.Contains("nesting exceeds the supported depth", response.Error.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(layoutPath);
        }
    }

    [Fact]
    public void Every_corpus_layout_passes_the_package_gate()
    {
        var corpusFiles = Directory.GetFiles(Path.Combine(AppContext.BaseDirectory, "corpus"), "*.docx");
        Assert.NotEmpty(corpusFiles);

        foreach (var path in corpusFiles)
        {
            using var doc = LayoutPackage.Open(path, isEditable: false);
            Assert.NotNull(doc.MainDocumentPart?.Document);
        }
    }

    /// <summary>
    /// Hand-assembled packages for shapes the SDK's own builders cannot produce directly (a long part chain,
    /// a DOCTYPE-bearing main part). Minimal: content types, the package relationship, a main part.
    /// </summary>
    private static class RawDocx
    {
        private const string ContentTypes =
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
            + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
            + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
            + "<Override PartName=\"/word/document.xml\" "
            + "ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
            + "</Types>";

        private const string PackageRels =
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rId1\" "
            + "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
            + "Target=\"word/document.xml\"/></Relationships>";

        private static readonly string PlainDocument =
            $"<w:document xmlns:w=\"{W}\"><w:body><w:p/><w:sectPr/></w:body></w:document>";

        internal static string WithDocumentXml(string documentXml) =>
            Build(z => Write(z, "word/document.xml", documentXml));

        /// <summary>document.xml → c/0.xml → c/1.xml → … through <paramref name="length"/> chained parts.</summary>
        internal static string WithPartChain(int length) => Build(z =>
        {
            Write(z, "word/document.xml", PlainDocument);
            Write(z, "word/_rels/document.xml.rels", Rels(("urn:bcwl-test:chain", "/c/0.xml")));
            for (var i = 0; i < length; i++)
            {
                Write(z, $"c/{i}.xml", "<a/>");
                if (i + 1 < length)
                {
                    Write(z, $"c/_rels/{i}.xml.rels", Rels(("urn:bcwl-test:chain", $"/c/{i + 1}.xml")));
                }
            }
        });

        /// <summary><paramref name="count"/> parts of <paramref name="bytesEach"/> bytes, all off document.xml.</summary>
        internal static string WithExtraXmlParts(int count, int bytesEach) => Build(z =>
        {
            Write(z, "word/document.xml", PlainDocument);
            Write(z, "word/_rels/document.xml.rels",
                Rels(Enumerable.Range(0, count).Select(i => ("urn:bcwl-test:extra", $"/x/{i}.xml")).ToArray()));
            var filler = new string('A', bytesEach - "<a></a>".Length);
            for (var i = 0; i < count; i++)
            {
                Write(z, $"x/{i}.xml", $"<a>{filler}</a>");
            }
        });

        private static string Rels(params (string Type, string Target)[] rels) =>
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + string.Concat(rels.Select((r, i) => $"<Relationship Id=\"r{i}\" Type=\"{r.Type}\" Target=\"{r.Target}\"/>"))
            + "</Relationships>";

        private static string Build(Action<ZipArchive> writeParts)
        {
            var path = Path.Combine(Path.GetTempPath(), $"bcwl-rawdocx-{Guid.NewGuid():N}.docx");
            using (var file = File.Create(path))
            using (var zip = new ZipArchive(file, ZipArchiveMode.Create))
            {
                Write(zip, "[Content_Types].xml", ContentTypes);
                Write(zip, "_rels/.rels", PackageRels);
                writeParts(zip);
            }

            return path;
        }

        private static void Write(ZipArchive zip, string name, string content)
        {
            using var writer = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false));
            writer.Write(content);
        }
    }
}
