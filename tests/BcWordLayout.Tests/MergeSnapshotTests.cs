using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using BcWordLayout.Merge;
using DocumentFormat.OpenXml.Packaging;

namespace BcWordLayout.Tests;

/// <summary>
/// Regression guard (design doc §7 "snapshot tests"): merges each corpus layout and compares the merged
/// main document part's XML against an approved snapshot, so any unintended change to the merge engine's
/// output shows up as a failing test with an actionable diff instead of silently shipping.
/// </summary>
/// <remarks>Joins the label-convention-seam collection: the <c>InventoryOrderDetails.docx</c> snapshot's
/// generated text differs for label vs. field columns (<c>SampleDataGenerator.GenerateLeafValue</c>), which
/// a concurrently-swapped <c>LabelConvention.Current</c> could disturb (see
/// <see cref="LabelConventionSeamCollection"/>).</remarks>
[Collection("label-convention-seam")]
public class MergeSnapshotTests
{
    private static string TempDocxPath() =>
        Path.Combine(Path.GetTempPath(), $"bcwl-snapshot-{Guid.NewGuid():N}.docx");

    /// <summary>
    /// Re-serializes an OOXML fragment (e.g. <c>MainDocumentPart.Document.OuterXml</c>, which the OpenXml
    /// SDK writes as a single unbroken line) through <see cref="XDocument"/> with indentation, so the
    /// approved snapshot is one meaningful element per line and a future mismatch produces a small,
    /// readable diff rather than one line-long needle-in-a-haystack.
    /// </summary>
    private static string PrettyPrint(string outerXml)
    {
        var xdoc = XDocument.Parse(outerXml);
        using var stringWriter = new StringWriter();
        using (var xmlWriter = XmlWriter.Create(stringWriter, new XmlWriterSettings
               {
                   Indent = true,
                   IndentChars = "  ",
                   NewLineChars = "\n",
                   OmitXmlDeclaration = true,
               }))
        {
            xdoc.Save(xmlWriter);
        }

        return stringWriter.ToString();
    }

    /// <summary>
    /// The OpenXml SDK's <c>AddNewPart&lt;ImagePart&gt;()</c> (used by <c>MergeEngine</c> to add the shared
    /// placeholder image part the first time a picture control in a given part needs one — see
    /// <c>MergeEngine.GetOrCreatePlaceholderImagePart</c>) mints a fresh, non-deterministic relationship id
    /// on every call, even for byte-identical input across separate runs — confirmed directly against the
    /// SDK (repeated <c>AddNewPart&lt;ImagePart&gt;</c> calls on fresh copies of the same corpus file each
    /// produced a different id, e.g. <c>R563e95a1abf8453f</c> / <c>R737d5c531db14cee</c>). This is real
    /// SDK behavior, not a MergeEngine defect — the merged document is correct either way (the blip still
    /// resolves to a valid embedded placeholder PNG) — but it means a picture control living in
    /// <c>document.xml</c> itself (rather than only in a header/footer, which this snapshot never looks at)
    /// makes the snapshot/determinism tests otherwise flaky. Normalize every <c>r:embed</c> value to a fixed
    /// placeholder before comparing so the snapshot still catches genuine content regressions without
    /// tripping on this known, harmless id churn.
    /// </summary>
    private static readonly Regex RelationshipEmbedId = new("r:embed=\"[^\"]*\"", RegexOptions.Compiled);

    private static string NormalizeVolatileRelationshipIds(string xml) =>
        RelationshipEmbedId.Replace(xml, "r:embed=\"R_NORMALIZED\"");

    private static string MergeAndPrettyPrintDocumentXml(string layoutPath, string outputPath, MergeOptions options)
    {
        MergeEngine.Merge(layoutPath, outputPath, options);

        using var doc = WordprocessingDocument.Open(outputPath, false);
        var outerXml = doc.MainDocumentPart!.Document!.OuterXml;
        return NormalizeVolatileRelationshipIds(PrettyPrint(outerXml));
    }

    [Theory]
    [InlineData(Corpus.SalesInvoice)]
    [InlineData(Corpus.InventoryOrderDetails)]
    [InlineData(Corpus.StandardStatement)]
    public void Merged_main_part_xml_matches_the_approved_snapshot(string fileName)
    {
        var outputPath = TempDocxPath();
        try
        {
            var prettyXml = MergeAndPrettyPrintDocumentXml(
                Corpus.Path(fileName), outputPath, new MergeOptions { Seed = 12345, Rows = 2 });

            SnapshotAssert.Match($"{fileName}.document.xml", prettyXml);
        }
        finally
        {
            File.Delete(outputPath);
        }
    }

    /// <summary>
    /// Guards the assumption snapshotting itself depends on: merging the same layout with the same
    /// options twice must produce byte-identical output. If this ever fails, some part of the merge
    /// (or this test's own pretty-printing) is order- or identity-sensitive in a way that would make every
    /// snapshot flaky, and must be fixed rather than worked around.
    /// </summary>
    [Fact]
    public void Merging_the_same_corpus_file_twice_produces_byte_identical_pretty_printed_xml()
    {
        var output1 = TempDocxPath();
        var output2 = TempDocxPath();
        try
        {
            var options = new MergeOptions { Seed = 12345, Rows = 2 };
            var layoutPath = Corpus.Path(Corpus.SalesInvoice);

            var xml1 = MergeAndPrettyPrintDocumentXml(layoutPath, output1, options);
            var xml2 = MergeAndPrettyPrintDocumentXml(layoutPath, output2, options);

            Assert.Equal(xml1, xml2);
        }
        finally
        {
            File.Delete(output1);
            File.Delete(output2);
        }
    }
}
