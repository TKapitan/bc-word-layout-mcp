using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Domain;

/// <summary>
/// Reads which ROLE each header/footer part plays — <c>default</c>, <c>first</c> (first page), or
/// <c>even</c> (even pages) — from the body's <c>w:headerReference</c>/<c>w:footerReference</c> elements,
/// plus whether the layout renders a different first page at all (<c>w:titlePg</c>). This is the fact
/// <c>get_layout_info</c> could not report before: a part LIST alone ("header1.xml, header2.xml,
/// header3.xml") says nothing about which part a one-page render actually shows, and in five of the six
/// corpus layouts that have header parts, the package-order FIRST part is NOT the everyday default one
/// (<c>header1.xml</c> is the EVEN-page header on <c>StandardSalesQuote</c>, <c>StandardPurchaseOrder</c>
/// and <c>JobQuote</c>) — the trap that cost a whole sandbox round to diagnose (GitHub issue #5).
/// </summary>
/// <remarks>
/// Role resolution walks every <c>w:sectPr</c> in document order and takes, per part, the FIRST reference
/// that names it (a paragraph-level <c>w:sectPr</c> ends the section it sits in, so document order is
/// section order). An ABSENT <c>w:type</c> is <c>default</c> per the spec, so a missing attribute counts
/// as default — the same rule <see cref="LocationResolver"/>'s default-part selection applies. A part no
/// section references at all gets a null role: it exists in the package but no page ever renders it.
/// <para>
/// <see cref="HasTitlePage"/> reads <c>w:titlePg</c> from the FIRST section only — the section whose
/// header/footer references govern page 1, which is where the "my insert is invisible on a one-page
/// document" trap lives. A <c>w:titlePg</c> with an explicit <c>w:val="false"</c> counts as OFF
/// (<c>w:titlePg</c> is an on/off element; presence alone is "on" only when the attribute is absent or
/// true).
/// </para>
/// </remarks>
public sealed class HeaderFooterPartRoles
{
    private readonly IReadOnlyDictionary<string, HeaderFooterRole> _rolesByPartName;

    private HeaderFooterPartRoles(IReadOnlyDictionary<string, HeaderFooterRole> rolesByPartName, bool hasTitlePage)
    {
        _rolesByPartName = rolesByPartName;
        HasTitlePage = hasTitlePage;
    }

    /// <summary>
    /// True when the FIRST section sets <c>w:titlePg</c> (renders a distinct first page): page 1 then shows
    /// the <c>first</c>-role header/footer — or nothing at all when no <c>first</c> reference exists — and
    /// NOT the <c>default</c> one.
    /// </summary>
    public bool HasTitlePage { get; }

    /// <summary>
    /// The role the first section-reference in document order assigns to the header/footer part named
    /// <paramref name="partFileName"/> (case-insensitive, e.g. <c>header2.xml</c>), or null when no section
    /// references it (or it is not a header/footer part at all — e.g. <c>document.xml</c>).
    /// </summary>
    public HeaderFooterRole? RoleOf(string partFileName) =>
        _rolesByPartName.TryGetValue(partFileName, out var role) ? role : null;

    /// <summary>Reads the reference roles and first-section <c>w:titlePg</c> of <paramref name="main"/>.</summary>
    public static HeaderFooterPartRoles Read(MainDocumentPart main)
    {
        ArgumentNullException.ThrowIfNull(main);

        var roles = new Dictionary<string, HeaderFooterRole>(StringComparer.OrdinalIgnoreCase);
        var body = main.Document?.Body;
        if (body is null)
        {
            return new HeaderFooterPartRoles(roles, hasTitlePage: false);
        }

        var sections = body.Descendants<SectionProperties>().ToList();
        foreach (var sectPr in sections)
        {
            foreach (var child in sectPr.ChildElements)
            {
                var (relationshipId, type) = child switch
                {
                    HeaderReference h => (h.Id?.Value, h.Type),
                    FooterReference f => (f.Id?.Value, f.Type),
                    _ => (null, null),
                };

                if (relationshipId is null)
                {
                    continue;
                }

                OpenXmlPart part;
                try
                {
                    part = main.GetPartById(relationshipId);
                }
                catch (ArgumentOutOfRangeException)
                {
                    // A reference naming a relationship the package does not contain is a (real, shipped)
                    // layout defect, not something to throw over while READING — skip it, exactly as
                    // LocationResolver's default-part selection does.
                    continue;
                }

                // First reference to a part wins: the first section's view is the one that governs the
                // opening page(s), which is what a caller reasons about.
                roles.TryAdd(PartWalker.PartFileName(part), ToRole(type));
            }
        }

        return new HeaderFooterPartRoles(roles, FirstSectionHasTitlePage(sections));
    }

    /// <summary>An absent <c>w:type</c> is <c>default</c> per the spec — same rule as <see cref="LocationResolver"/>.</summary>
    private static HeaderFooterRole ToRole(DocumentFormat.OpenXml.EnumValue<HeaderFooterValues>? type)
    {
        if (type is null)
        {
            return HeaderFooterRole.Default;
        }

        if (type.Value == HeaderFooterValues.First)
        {
            return HeaderFooterRole.First;
        }

        return type.Value == HeaderFooterValues.Even ? HeaderFooterRole.Even : HeaderFooterRole.Default;
    }

    private static bool FirstSectionHasTitlePage(List<SectionProperties> sections)
    {
        var titlePg = sections.Count > 0 ? sections[0].GetFirstChild<TitlePage>() : null;
        if (titlePg is null)
        {
            return false;
        }

        // w:titlePg is an on/off element: present with no w:val (or w:val true) = on; explicit false = off.
        return titlePg.Val is null || titlePg.Val.Value;
    }
}
