using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// Covers GitHub issue #5: header/footer part ROLES (default/first/even) must be visible to a caller —
/// <see cref="HeaderFooterPartRoles"/> reading the section references, <see cref="LayoutReader"/>'s
/// <see cref="LayoutInventory.PartDetails"/>/<see cref="LayoutInventory.HasTitlePage"/>, and the
/// first-page-visibility note <see cref="LayoutEditor"/> appends when an insert lands in the default
/// header/footer of a <c>w:titlePg</c> layout. Expected corpus values were pinned by reading the raw
/// <c>w:headerReference</c>/<c>w:footerReference</c> elements on 2026-08-08: on
/// <c>StandardPurchaseOrder.docx</c>, <c>header1.xml</c> is EVEN, <c>header2.xml</c> is DEFAULT and
/// <c>header3.xml</c> is FIRST (footers identical), with <c>w:titlePg</c> set — the exact layout whose
/// invisible default-header insert cost a sandbox round to diagnose.
/// </summary>
public class HeaderFooterPartRolesTests
{
    private static string CopyOfCorpus(string corpusFile)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-partroles-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    // ---- HeaderFooterPartRoles (the raw section-reference read) ----

    [Fact]
    public void PurchaseOrder_roles_match_the_raw_section_references()
    {
        using var doc = WordprocessingDocument.Open(Corpus.Path(Corpus.StandardPurchaseOrder), false);
        var roles = HeaderFooterPartRoles.Read(doc.MainDocumentPart!);

        Assert.True(roles.HasTitlePage);
        Assert.Equal(HeaderFooterRole.Even, roles.RoleOf("header1.xml"));
        Assert.Equal(HeaderFooterRole.Default, roles.RoleOf("header2.xml"));
        Assert.Equal(HeaderFooterRole.First, roles.RoleOf("header3.xml"));
        Assert.Equal(HeaderFooterRole.Even, roles.RoleOf("footer1.xml"));
        Assert.Equal(HeaderFooterRole.Default, roles.RoleOf("footer2.xml"));
        Assert.Equal(HeaderFooterRole.First, roles.RoleOf("footer3.xml"));

        // The main document part plays no header/footer role, and lookups are case-insensitive.
        Assert.Null(roles.RoleOf("document.xml"));
        Assert.Equal(HeaderFooterRole.Default, roles.RoleOf("HEADER2.XML"));
    }

    [Fact]
    public void SalespersonCommission_has_no_title_page_and_a_default_only_header()
    {
        // The valid sibling that must NOT trip the first-page machinery: one default header reference,
        // no w:titlePg at all.
        using var doc = WordprocessingDocument.Open(Corpus.Path(Corpus.SalespersonCommission), false);
        var roles = HeaderFooterPartRoles.Read(doc.MainDocumentPart!);

        Assert.False(roles.HasTitlePage);
        Assert.Equal(HeaderFooterRole.Default, roles.RoleOf("header1.xml"));
    }

    // ---- LayoutReader.PartDetails (what get_layout_info surfaces) ----

    [Fact]
    public void PurchaseOrder_part_details_report_kind_role_and_default_target()
    {
        var inv = LayoutReader.Read(Corpus.Path(Corpus.StandardPurchaseOrder));

        Assert.True(inv.HasTitlePage);

        // PartDetails mirrors Parts one to one, same order.
        Assert.Equal(inv.Parts, inv.PartDetails.Select(p => p.Name).ToList());

        var document = inv.PartDetails.Single(p => p.Name == "document.xml");
        Assert.Equal(LayoutPartKind.Document, document.Kind);
        Assert.Null(document.Role);
        Assert.False(document.IsDefaultTarget);

        var header2 = inv.PartDetails.Single(p => p.Name == "header2.xml");
        Assert.Equal(LayoutPartKind.Header, header2.Kind);
        Assert.Equal(HeaderFooterRole.Default, header2.Role);
        Assert.True(header2.IsDefaultTarget);

        // Exactly ONE header and ONE footer is the partName-less target, and it is the DEFAULT-role one —
        // NOT header1.xml/footer1.xml, which are the even-page parts on this layout.
        var headerTargets = inv.PartDetails.Where(p => p.Kind == LayoutPartKind.Header && p.IsDefaultTarget).ToList();
        var footerTargets = inv.PartDetails.Where(p => p.Kind == LayoutPartKind.Footer && p.IsDefaultTarget).ToList();
        Assert.Equal("header2.xml", Assert.Single(headerTargets).Name);
        Assert.Equal("footer2.xml", Assert.Single(footerTargets).Name);
        Assert.False(inv.PartDetails.Single(p => p.Name == "header1.xml").IsDefaultTarget);

        var header3 = inv.PartDetails.Single(p => p.Name == "header3.xml");
        Assert.Equal(HeaderFooterRole.First, header3.Role);
    }

    [Fact]
    public void An_unreferenced_header_part_reports_a_null_role_but_remains_the_fallback_target()
    {
        // Strip the (single) header reference from SalespersonCommission's sectPr: the part is then in the
        // package but rendered on no page — role must honestly be null, while the partName-less target
        // falls back to package order (LocationResolver's own fallback), which this same part still wins.
        var path = CopyOfCorpus(Corpus.SalespersonCommission);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var body = doc.MainDocumentPart!.Document!.Body!;
                foreach (var reference in body.Descendants<SectionProperties>()
                             .SelectMany(s => s.Elements<HeaderReference>()).ToList())
                {
                    reference.Remove();
                }

                doc.MainDocumentPart.Document.Save();
            }

            var inv = LayoutReader.Read(path);
            var header = inv.PartDetails.Single(p => p.Name == "header1.xml");
            Assert.Null(header.Role);
            Assert.True(header.IsDefaultTarget);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_layout_with_no_header_or_footer_parts_reports_only_the_document_part()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.PlainParagraph("body only"));
        try
        {
            var inv = LayoutReader.Read(path);

            Assert.False(inv.HasTitlePage);
            var part = Assert.Single(inv.PartDetails);
            Assert.Equal("document.xml", part.Name);
            Assert.Equal(LayoutPartKind.Document, part.Kind);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- LayoutEditor's first-page-visibility note ----

    private const string PurchaseOrderField = "/Purchase_Header/BuyFromAddr1";

    private static EditResult InsertIntoHeader(string path, string? partName)
    {
        using var doc = WordprocessingDocument.Open(path, true);
        var result = LayoutEditor.InsertField(doc, PurchaseOrderField, new Location
        {
            Type = LocationKind.DocumentEnd,
            Part = LayoutPart.Header,
            PartName = partName,
        });
        doc.MainDocumentPart!.Document!.Save();
        foreach (var header in doc.MainDocumentPart.HeaderParts)
        {
            header.Header?.Save();
        }

        return result;
    }

    [Fact]
    public void Insert_into_default_header_of_a_titlePg_layout_warns_it_is_invisible_on_page_1()
    {
        var path = CopyOfCorpus(Corpus.StandardPurchaseOrder);
        try
        {
            var result = InsertIntoHeader(path, partName: null);

            Assert.Equal("header2.xml", result.Part);
            Assert.Contains("DIFFERENT FIRST PAGE", result.Summary, StringComparison.Ordinal);
            Assert.Contains("partDetails", result.Summary, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Insert_with_an_explicit_partName_gets_no_first_page_note_even_on_a_titlePg_layout()
    {
        // The caller who NAMES the default part chose it deliberately; the note exists for the
        // partName-less resolution trap only.
        var path = CopyOfCorpus(Corpus.StandardPurchaseOrder);
        try
        {
            var result = InsertIntoHeader(path, partName: "header2.xml");

            Assert.Equal("header2.xml", result.Part);
            Assert.DoesNotContain("DIFFERENT FIRST PAGE", result.Summary, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Insert_into_the_header_of_a_layout_without_titlePg_gets_no_first_page_note()
    {
        // The valid sibling: JobQuote has the same even/default/first header trio as StandardPurchaseOrder
        // but NO w:titlePg, so its default header does render on page 1 and the note must not fire.
        var path = CopyOfCorpus(Corpus.JobQuote);
        try
        {
            EditResult result;
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                result = LayoutEditor.InsertField(
                    doc,
                    "/Job/BillToAddress1",
                    new Location { Type = LocationKind.DocumentEnd, Part = LayoutPart.Header });
                doc.MainDocumentPart!.Document!.Save();
                foreach (var header in doc.MainDocumentPart.HeaderParts)
                {
                    header.Header?.Save();
                }
            }

            Assert.Equal("header2.xml", result.Part);
            Assert.DoesNotContain("DIFFERENT FIRST PAGE", result.Summary, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
