using System.Text;
using BcWordLayout.Domain;
using BcWordLayout.McpHost.Tools;
using BcWordLayout.Merge;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace BcWordLayout.Tests;

/// <summary>
/// Covers the two crash classes <see cref="ResourceLimits"/> exists to prevent: an oversized/zip-bomb
/// custom XML part loaded whole (→ OOM) and
/// (unbounded recursion over nested XML/document structure → uncatchable <see cref="StackOverflowException"/>),
/// plus two follow-ups: aggregate OOM from retaining EVERY BC-namespaced custom
/// XML part's parsed root simultaneously (and the missing hard part-count ceiling), and the dedicated
/// <see cref="ResourceLimitExceededException"/> type and its tailored <c>invalid_layout</c> hint). Proves a
/// too-large part/file, a package with too many custom XML parts, and a too-deeply-nested schema/document are
/// all rejected as a normal <c>invalid_layout</c> tool failure (or, for <see cref="MergeEngine"/>'s own
/// recursive walk, a plain <see cref="ResourceLimitExceededException"/>) rather than crashing the process, and
/// that a LEGITIMATE multi-part layout still validates exactly as before that refactor. Every fixture here
/// is generated at runtime (never committed) and kept just over the relevant <see cref="ResourceLimits"/>
/// cap/depth/count so these tests stay fast.
/// </summary>
public class ResourceLimitsTests
{
    private const string BcNamespace = "urn:microsoft-dynamics-nav/reports/TestReport/50000/";

    // ---- Oversized custom XML part / schema file (the "zip bomb" vector) ----

    [Fact]
    public void ListDatasetFields_with_oversized_standalone_schema_file_returns_invalid_layout()
    {
        // Just over ResourceLimits.MaxCustomXmlPartBytes (16 MB): one huge filler text node inside an
        // otherwise well-formed BC-namespaced schema document.
        var filler = new string('A', (int)ResourceLimits.MaxCustomXmlPartBytes + 1024);
        var xmlPath = Path.Combine(Path.GetTempPath(), $"bcwl-oversized-schema-{Guid.NewGuid():N}.xml");
        File.WriteAllText(
            xmlPath,
            $"<NavWordReportXmlPart xmlns=\"{BcNamespace}\"><Filler>{filler}</Filler></NavWordReportXmlPart>");

        try
        {
            var response = ReadTools.ListDatasetFields(xmlPath);

            Assert.False(response.Ok);
            Assert.NotNull(response.Error);
            Assert.Equal("invalid_layout", response.Error!.Code);
            Assert.Contains("supported size limit", response.Error.Message, StringComparison.Ordinal);
            Assert.Contains("16 MB", response.Error.Message, StringComparison.Ordinal);

            // The tailored ResourceLimitExceededException hint fires (not the generic "missing dataset
            // part/wrong namespace" one a plain InvalidDataException gets).
            Assert.Contains("supported size/nesting limits", response.Error.Hint, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(xmlPath);
        }
    }

    [Fact]
    public void GetLayoutInfo_with_oversized_custom_xml_part_returns_invalid_layout()
    {
        // The real zip-bomb-shaped vector: the OVERSIZED part lives INSIDE the .docx package (a custom XML
        // part), not as a standalone file - proving the per-part cap inside SchemaProvider.FindBcPart.
        var filler = new string('A', (int)ResourceLimits.MaxCustomXmlPartBytes + 1024);
        var hugeDatasetXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<NavWordReportXmlPart xmlns=\"{BcNamespace}\"><Filler>{filler}</Filler></NavWordReportXmlPart>";

        var layoutPath = SyntheticLayout.Create(
            SyntheticLayout.PlainParagraph("x"), datasetXml: hugeDatasetXml);

        try
        {
            var response = ReadTools.GetLayoutInfo(layoutPath);

            Assert.False(response.Ok);
            Assert.NotNull(response.Error);
            Assert.Equal("invalid_layout", response.Error!.Code);
            Assert.Contains("supported size limit", response.Error.Message, StringComparison.Ordinal);
            Assert.Contains("16 MB", response.Error.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(layoutPath);
        }
    }

    // ---- Unbounded recursion over nested XML/document structure ----

    /// <summary>Builds <paramref name="depth"/> singly-nested wrapper elements ending in one leaf column.</summary>
    private static string NestedSchemaBody(int depth)
    {
        var open = new StringBuilder();
        var close = new StringBuilder();
        for (var i = 0; i < depth; i++)
        {
            open.Append($"<Level{i}>");
            close.Insert(0, $"</Level{i}>");
        }

        return open + "<Leaf>x</Leaf>" + close;
    }

    [Fact]
    public void ListDatasetFields_with_schema_nesting_beyond_depth_limit_returns_invalid_layout()
    {
        // Comfortably past ResourceLimits.MaxSchemaDepth (64).
        var depth = ResourceLimits.MaxSchemaDepth + 20;
        var xmlPath = Path.Combine(Path.GetTempPath(), $"bcwl-deep-schema-{Guid.NewGuid():N}.xml");
        File.WriteAllText(
            xmlPath, $"<NavWordReportXmlPart xmlns=\"{BcNamespace}\">{NestedSchemaBody(depth)}</NavWordReportXmlPart>");

        try
        {
            var response = ReadTools.ListDatasetFields(xmlPath);

            Assert.False(response.Ok);
            Assert.NotNull(response.Error);
            Assert.Equal("invalid_layout", response.Error!.Code);
            Assert.Contains("nesting exceeds the supported depth", response.Error.Message, StringComparison.Ordinal);
            Assert.Contains(ResourceLimits.MaxSchemaDepth.ToString(), response.Error.Message, StringComparison.Ordinal);

            // same tailored hint fires for a DEPTH rejection as for a SIZE rejection above.
            Assert.Contains("supported size/nesting limits", response.Error.Hint, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(xmlPath);
        }
    }

    /// <summary>Builds <paramref name="depth"/> nested content-control wrappers (no <c>w:sdtPr</c> needed)
    /// around one leaf paragraph — a shape <see cref="BcWordLayout.Domain.LayoutReader.Walk"/> and
    /// <see cref="MergeEngine"/>'s own document walk both recurse through once per level.</summary>
    private static string NestedSdtChain(int depth)
    {
        var content = "<w:p><w:r><w:t>x</w:t></w:r></w:p>";
        for (var i = 0; i < depth; i++)
        {
            content = $"<w:sdt><w:sdtContent>{content}</w:sdtContent></w:sdt>";
        }

        return content;
    }

    [Fact]
    public void GetLayoutInfo_with_document_nesting_beyond_depth_limit_returns_invalid_layout()
    {
        // Comfortably past ResourceLimits.MaxElementNestingDepth (128), but shallow enough that the OpenXml
        // SDK's own (out-of-scope) document load never comes close to its own limits.
        var depth = ResourceLimits.MaxElementNestingDepth + 20;
        var layoutPath = SyntheticLayout.Create(NestedSdtChain(depth));

        try
        {
            var response = ReadTools.GetLayoutInfo(layoutPath);

            Assert.False(response.Ok);
            Assert.NotNull(response.Error);
            Assert.Equal("invalid_layout", response.Error!.Code);
            Assert.Contains("nesting exceeds the supported depth", response.Error.Message, StringComparison.Ordinal);
            Assert.Contains(ResourceLimits.MaxElementNestingDepth.ToString(), response.Error.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(layoutPath);
        }
    }

    [Fact]
    public void Merge_throws_ResourceLimitExceededException_on_document_nesting_beyond_depth_limit()
    {
        // Same nested shape as the LayoutReader test above, but exercises MergeEngine.WalkElement's OWN
        // depth guard directly - a distinct hand-rolled recursive walker over the same kind of untrusted
        // document structure (see the B10 inventory: LayoutReader.Walk and MergeEngine.WalkElement are two
        // separate methods, each needing its own counter).
        var depth = ResourceLimits.MaxElementNestingDepth + 20;
        var layoutPath = SyntheticLayout.Create(NestedSdtChain(depth));

        // An isolated directory (not the shared temp root) so the "no staging leftovers" check below can't
        // be confused by another test's concurrent merge into the same shared temp directory.
        var outputDir = Path.Combine(Path.GetTempPath(), $"bcwl-merge-deep-{Guid.NewGuid():N}");
        Directory.CreateDirectory(outputDir);
        var outputPath = Path.Combine(outputDir, "output.docx");

        try
        {
            var ex = Assert.Throws<ResourceLimitExceededException>(
                () => MergeEngine.Merge(layoutPath, outputPath, new MergeOptions()));

            Assert.Contains("nesting exceeds the supported depth", ex.Message, StringComparison.Ordinal);
            Assert.Contains(ResourceLimits.MaxElementNestingDepth.ToString(), ex.Message, StringComparison.Ordinal);

            // a mid-merge exception must leave NO output file behind - not even a half-merged one.
            // Previously the merge opened outputPath directly with AutoSave on, so a partially-walked tree
            // could be flushed to it on Dispose; the merge now stages into a throwaway copy and only ever
            // replaces outputPath on a fully successful merge.
            Assert.False(File.Exists(outputPath), "a failed merge must not leave any file at outputPath");

            // And no staging leftovers either (the same directory as outputPath - see MergeEngine.Merge).
            Assert.Empty(Directory.GetFiles(outputDir, ".bcwl-merge-stage-*"));
        }
        finally
        {
            File.Delete(layoutPath);
            Directory.Delete(outputDir, recursive: true);
        }
    }

    [Fact]
    public void Merge_throws_leaves_a_preexisting_output_file_byte_identical()
    {
        // Symmetric with the "no output file" case above: when outputPath ALREADY holds a (different)
        // file before a mid-merge exception, that file must survive completely untouched - the same
        // "file on disk left untouched on failure" guarantee the edit-tool pipeline gives, extended here
        // to the merge pipeline via staging (see MergeEngine.Merge's remarks).
        var depth = ResourceLimits.MaxElementNestingDepth + 20;
        var layoutPath = SyntheticLayout.Create(NestedSdtChain(depth));
        var outputPath = Path.Combine(Path.GetTempPath(), $"bcwl-merge-deep-preexisting-{Guid.NewGuid():N}.docx");
        var preexistingBytes = "not a real docx - just a marker file"u8.ToArray();
        File.WriteAllBytes(outputPath, preexistingBytes);

        try
        {
            Assert.Throws<ResourceLimitExceededException>(
                () => MergeEngine.Merge(layoutPath, outputPath, new MergeOptions()));

            Assert.True(File.Exists(outputPath), "the pre-existing output file must survive a failed merge");
            Assert.Equal(preexistingBytes, File.ReadAllBytes(outputPath));
        }
        finally
        {
            File.Delete(layoutPath);
            File.Delete(outputPath);
        }
    }

    // ---- Aggregate OOM from retaining every BC-namespaced part's parsed root simultaneously ----

    /// <summary>
    /// Builds a minimal .docx with <paramref name="bcNamespacedPartCount"/> small BC-namespaced custom XML
    /// parts (the first carries a <see cref="CustomXmlPropertiesPart"/>/<c>ds:itemID</c>, matching
    /// <see cref="SyntheticLayout.Create"/>'s shape, so it is discoverable exactly like a real BC part) plus
    /// <paramref name="otherPartCount"/> unrelated (non-BC-namespaced) custom XML parts. Every part is tiny -
    /// this exists purely to exercise PART COUNT, not per-part size.
    /// </summary>
    private static string BuildDocxWithCustomXmlParts(int bcNamespacedPartCount, int otherPartCount = 0)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-multipart-{Guid.NewGuid():N}.docx");
        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            WriteRaw(
                main.GetStream(FileMode.Create, FileAccess.Write),
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
                + "<w:body><w:p/><w:sectPr/></w:body></w:document>");

            for (var i = 0; i < bcNamespacedPartCount; i++)
            {
                var cxp = main.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                WriteRaw(
                    cxp.GetStream(FileMode.Create, FileAccess.Write),
                    $"<NavWordReportXmlPart xmlns=\"{BcNamespace}\"><Header><CompanyName>c{i}</CompanyName>"
                    + "</Header></NavWordReportXmlPart>");

                if (i == 0)
                {
                    var props = cxp.AddNewPart<CustomXmlPropertiesPart>();
                    WriteRaw(
                        props.GetStream(FileMode.Create, FileAccess.Write),
                        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?>"
                        + "<ds:datastoreItem ds:itemID=\"{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}\" "
                        + "xmlns:ds=\"http://schemas.openxmlformats.org/officeDocument/2006/customXml\">"
                        + "<ds:schemaRefs/></ds:datastoreItem>");
                }
            }

            for (var i = 0; i < otherPartCount; i++)
            {
                var cxp = main.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                WriteRaw(cxp.GetStream(FileMode.Create, FileAccess.Write), $"<Other xmlns=\"urn:not-bc\">{i}</Other>");
            }
        }

        return path;
    }

    private static void WriteRaw(Stream stream, string content)
    {
        using (stream)
        using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
        {
            writer.Write(content);
        }
    }

    [Fact]
    public void Quick_validation_reports_multiple_BC_parts_exactly_as_before_the_F1_refactor()
    {
        // Pins existing behavior across the aggregate-memory refactor: SchemaProvider.FindBcParts now
        // returns bare CustomXmlPart references rather than (part, parsed-root) pairs for every match (see
        // its own remarks), but LayoutValidator.CheckSingleBcPart's duplicate-part detection - which only
        // ever needed the COUNT plus the FIRST part's properties, never every root - must still count/report
        // identically. Two small parts (legitimate scale, not the ceiling test below).
        var layoutPath = BuildDocxWithCustomXmlParts(bcNamespacedPartCount: 2);

        try
        {
            var result = LayoutValidator.Quick(layoutPath);

            Assert.Contains(
                result.Findings,
                f => f.Check == "single-bc-part"
                    && f.Message.Contains("Found 2 BC dataset custom XML parts", StringComparison.Ordinal));
        }
        finally
        {
            File.Delete(layoutPath);
        }
    }

    [Fact]
    public void GetLayoutInfo_with_custom_xml_part_count_over_ceiling_returns_invalid_layout()
    {
        // Comfortably past ResourceLimits.MaxCustomXmlParts (1024). None need to be BC-namespaced or even
        // well-formed XML - EnsurePartCountWithinLimit rejects up front, before any part is opened/parsed, so
        // the per-part iteration/parse LOOP itself stays bounded regardless of how cheap each individual part
        // is (the addendum's second requirement, distinct from the per-part byte cap above).
        var partCount = ResourceLimits.MaxCustomXmlParts + 1;
        var layoutPath = BuildDocxWithCustomXmlParts(bcNamespacedPartCount: 0, otherPartCount: partCount);

        try
        {
            var response = ReadTools.GetLayoutInfo(layoutPath);

            Assert.False(response.Ok);
            Assert.NotNull(response.Error);
            Assert.Equal("invalid_layout", response.Error!.Code);
            Assert.Contains("custom XML parts", response.Error.Message, StringComparison.Ordinal);
            Assert.Contains(ResourceLimits.MaxCustomXmlParts.ToString(), response.Error.Message, StringComparison.Ordinal);
            Assert.Contains("supported size/nesting limits", response.Error.Hint, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(layoutPath);
        }
    }

    // ---- sanity: the real corpus (small, shallow, well within every cap) still loads fine ----
    // Covered implicitly by the rest of the suite (SchemaProviderTests, LayoutReaderTests, MergeEngineTests,
    // McpHostToolTests, etc. all exercise the corpus through these same code paths) - no separate test needed.
}
