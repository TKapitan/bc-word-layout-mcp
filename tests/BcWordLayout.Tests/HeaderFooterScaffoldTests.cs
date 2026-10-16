using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using BcWordLayout.McpHost.Tools;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// GitHub issue #91: a header/footer scaffolded on demand was referenced only from the LAST section's
/// <c>w:sectPr</c>, so in a multi-section layout every earlier section — page 1 included — inherited no
/// header/footer at all (OOXML inherits a missing reference from the PREVIOUS section, never the next). It
/// must be wired into the FIRST <c>w:sectPr</c> in document order, the one every other reader
/// (<see cref="HeaderFooterPartRoles"/>, <see cref="LocationResolver"/>'s default-part rule) already treats
/// as governing. Also: the scaffolded part's own empty paragraph must be reused by the first
/// <c>documentEnd</c> insert instead of being left behind as a leading empty line.
/// </summary>
public class HeaderFooterScaffoldTests
{
    /// <summary>
    /// A copy of JobQuote — the only multi-section stock layout (one paragraph-level <c>w:sectPr</c> plus
    /// three continuous ones, references on the FIRST) — with every header/footer part and reference
    /// removed: the shape the scaffold has to handle.
    /// </summary>
    private static string JobQuoteWithoutHeadersOrFooters()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-hfscaffold-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(Corpus.JobQuote), path, overwrite: true);

        using var doc = WordprocessingDocument.Open(path, true);
        var main = doc.MainDocumentPart!;
        foreach (var reference in main.Document!.Body!.Descendants<SectionProperties>()
                     .SelectMany(s => s.ChildElements.Where(c => c is HeaderReference or FooterReference)).ToList())
        {
            reference.Remove();
        }

        foreach (var part in main.HeaderParts.Cast<OpenXmlPart>().Concat(main.FooterParts).ToList())
        {
            main.DeletePart(part);
        }

        main.Document.Save();
        return path;
    }

    [Theory]
    [InlineData("header")]
    [InlineData("footer")]
    public void Scaffolded_part_on_a_multi_section_layout_is_referenced_from_the_first_section(string layoutPart)
    {
        var path = JobQuoteWithoutHeadersOrFooters();
        try
        {
            var response = EditTools.InsertText(path, "SCAFFOLDED", "documentEnd", layoutPart: layoutPart);
            Assert.True(response.Ok, response.Error?.Message);

            using var doc = WordprocessingDocument.Open(path, false);
            var main = doc.MainDocumentPart!;
            var sections = main.Document!.Body!.Descendants<SectionProperties>().ToList();
            Assert.Equal(4, sections.Count);
            Assert.IsType<ParagraphProperties>(sections[0].Parent); // the first is the paragraph-level one

            OpenXmlPart part = layoutPart == "header" ? Assert.Single(main.HeaderParts) : Assert.Single(main.FooterParts);
            var id = main.GetIdOfPart(part);
            var referenceCounts = sections
                .Select(s => s.ChildElements.Count(c => (c is HeaderReference h && h.Id?.Value == id) || (c is FooterReference f && f.Id?.Value == id)))
                .ToList();
            Assert.Equal(new[] { 1, 0, 0, 0 }, referenceCounts);

            // Every reader of section roles agrees: it is the default part, and a later layoutPart insert
            // with no partName resolves to it.
            Assert.Equal(HeaderFooterRole.Default, HeaderFooterPartRoles.Read(main).RoleOf(PartWalker.PartFileName(part)));
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Scaffolded_part_on_a_synthetic_two_section_layout_is_referenced_from_the_paragraph_level_sectPr()
    {
        var path = SyntheticLayout.Create(
            "<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>section 1</w:t></w:r></w:p>"
            + SyntheticLayout.PlainParagraph("section 2"));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            Assert.True(HeaderFooterScaffold.EnsureExists(doc, LayoutPart.Footer));

            var body = doc.MainDocumentPart!.Document!.Body!;
            var first = body.Descendants<SectionProperties>().First();
            var last = body.Elements<SectionProperties>().Single();
            Assert.NotSame(first, last);
            Assert.Single(first.Elements<FooterReference>());
            Assert.Empty(last.Elements<FooterReference>());
            Assert.IsType<FooterReference>(first.FirstChild); // CT_SectPr's leading choice group
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("header")]
    [InlineData("footer")]
    public void First_documentEnd_insert_into_a_scaffolded_part_reuses_its_empty_paragraph(string layoutPart)
    {
        var path = SyntheticLayout.Create(SyntheticLayout.PlainParagraph("body"));
        try
        {
            Assert.True(EditTools.InsertText(path, "FIRST", "documentEnd", layoutPart: layoutPart).Ok);

            using (var doc = WordprocessingDocument.Open(path, false))
            {
                OpenXmlElement root = layoutPart == "header"
                    ? Assert.Single(doc.MainDocumentPart!.HeaderParts).Header!
                    : Assert.Single(doc.MainDocumentPart!.FooterParts).Footer!;

                // <w:ftr><w:p><w:r><w:t>FIRST</w:t></w:r></w:p></w:ftr> - no leading <w:p/>.
                var paragraph = Assert.IsType<Paragraph>(Assert.Single(root.ChildElements));
                Assert.Equal("FIRST", paragraph.InnerText);
            }

            // Once the part has real content, a further documentEnd insert opens its own paragraph, exactly
            // as it does in the body.
            Assert.True(EditTools.InsertText(path, "SECOND", "documentEnd", layoutPart: layoutPart).Ok);

            using (var doc = WordprocessingDocument.Open(path, false))
            {
                OpenXmlElement root = layoutPart == "header"
                    ? doc.MainDocumentPart!.HeaderParts.Single().Header!
                    : doc.MainDocumentPart!.FooterParts.Single().Footer!;
                Assert.Equal(new[] { "FIRST", "SECOND" }, root.Elements<Paragraph>().Select(p => p.InnerText));
            }
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void DocumentEnd_block_insert_into_a_lone_empty_paragraph_part_goes_before_it()
    {
        // A block (a table) cannot BE that paragraph, so it goes in front of it - and the paragraph stays as
        // the table's required trailing paragraph rather than a second empty one being added.
        var path = SyntheticLayout.Create(SyntheticLayout.PlainParagraph("body"));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            HeaderFooterScaffold.EnsureExists(doc, LayoutPart.Header);
            var header = doc.MainDocumentPart!.HeaderParts.Single().Header!;
            var emptyParagraph = header.GetFirstChild<Paragraph>()!;

            var anchor = LocationResolver.Resolve(
                new Location { Type = LocationKind.DocumentEnd, Part = LayoutPart.Header }, doc);
            var table = new Table(new TableProperties(), new TableGrid(new GridColumn { Width = "1000" }),
                new TableRow(new TableCell(new Paragraph())));
            anchor.InsertBlock(table);

            Assert.Equal(2, header.ChildElements.Count);
            Assert.Same(table, header.FirstChild);
            Assert.Same(emptyParagraph, header.LastChild);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void DocumentEnd_into_a_part_with_real_content_still_appends_a_new_paragraph()
    {
        var path = SyntheticLayout.CreateWithHeader(
            SyntheticLayout.PlainParagraph("body"), SyntheticLayout.PlainParagraph("existing header text"));
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var anchor = LocationResolver.Resolve(
                new Location { Type = LocationKind.DocumentEnd, Part = LayoutPart.Header }, doc);
            var marker = new Run(new Text("MARKER"));
            anchor.InsertInline(marker);

            var header = doc.MainDocumentPart!.HeaderParts.Single().Header!;
            Assert.Equal(2, header.Elements<Paragraph>().Count());
            Assert.Same(header.LastChild, marker.Parent);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
