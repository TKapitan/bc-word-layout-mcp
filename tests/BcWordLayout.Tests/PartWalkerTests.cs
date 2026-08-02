using BcWordLayout.Domain;
using DocumentFormat.OpenXml.Packaging;

namespace BcWordLayout.Tests;

/// <summary>
/// Direct coverage of <see cref="PartWalker"/> — the single "document.xml, then every header, then every
/// footer, skipping any part with no content root" implementation that replaced roughly eight hand-written
/// copies of the same three-block pattern.
/// </summary>
public class PartWalkerTests
{
    [Fact]
    public void ContentParts_yields_document_xml_first_with_the_main_documents_own_root()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.PlainParagraph("hello"));
        try
        {
            using var doc = WordprocessingDocument.Open(path, false);
            var main = doc.MainDocumentPart!;

            var parts = PartWalker.ContentParts(main).ToList();

            var first = Assert.Single(parts);
            Assert.Equal("document.xml", first.PartName);
            Assert.Same(main.Document, first.Root);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void ContentParts_includes_a_header_part_after_document_xml_in_order()
    {
        var path = SyntheticLayout.CreateWithHeader(
            SyntheticLayout.PlainParagraph("body"), SyntheticLayout.PlainParagraph("header"));
        try
        {
            using var doc = WordprocessingDocument.Open(path, false);
            var main = doc.MainDocumentPart!;

            var parts = PartWalker.ContentParts(main).ToList();

            Assert.Equal(2, parts.Count);
            Assert.Equal("document.xml", parts[0].PartName);
            Assert.Same(main.Document, parts[0].Root);

            var headerPart = main.HeaderParts.Single();
            Assert.Equal(PartWalker.PartFileName(headerPart), parts[1].PartName);
            Assert.Same(headerPart.Header, parts[1].Root);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void ContentPartsWithHost_yields_the_owning_OpenXmlPart_alongside_the_root_and_name()
    {
        var path = SyntheticLayout.CreateWithHeader(
            SyntheticLayout.PlainParagraph("body"), SyntheticLayout.PlainParagraph("header"));
        try
        {
            using var doc = WordprocessingDocument.Open(path, false);
            var main = doc.MainDocumentPart!;

            var parts = PartWalker.ContentPartsWithHost(main).ToList();

            Assert.Equal(2, parts.Count);
            Assert.Same(main, parts[0].Part);
            Assert.Equal("document.xml", parts[0].PartName);

            var headerPart = main.HeaderParts.Single();
            Assert.Same(headerPart, parts[1].Part);
            Assert.Equal(PartWalker.PartFileName(headerPart), parts[1].PartName);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void PartFileName_returns_the_parts_own_uri_file_name()
    {
        var path = SyntheticLayout.CreateWithHeader(
            SyntheticLayout.PlainParagraph("body"), SyntheticLayout.PlainParagraph("header"));
        try
        {
            using var doc = WordprocessingDocument.Open(path, false);
            var headerPart = doc.MainDocumentPart!.HeaderParts.Single();

            Assert.Equal("header1.xml", PartWalker.PartFileName(headerPart));
        }
        finally
        {
            File.Delete(path);
        }
    }
}
