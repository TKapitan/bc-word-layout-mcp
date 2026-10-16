using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Domain;

/// <summary>
/// Creates the empty <see cref="HeaderPart"/>/<see cref="FooterPart"/> a layout must already have before
/// anything can be authored into one. A blank <see cref="LayoutBuilder.Create"/> output (no
/// <c>templatePath</c>) has neither, so <c>insert_field layoutPart='footer'</c> on a from-scratch layout
/// used to fail <c>not_found</c> with no way forward short of starting over from a template shell — an
/// easy trap, and the binding constraint on authoring a real BC document from nothing (a legal/contact
/// block belongs in a footer that repeats per page, exactly as every corpus layout does it, not in the
/// body). Two callers close it from both ends: <see cref="LayoutBuilder"/> scaffolds both parts up front
/// for a blank build (the corpus shape — every captured layout has at least one header and one footer),
/// and <see cref="LayoutEditor"/> scaffolds on demand for a layout that predates that or was built
/// elsewhere.
/// </summary>
/// <remarks>
/// A scaffolded part holds exactly one empty paragraph (a <c>w:hdr</c>/<c>w:ftr</c> must carry at least
/// one block-level child — the first <c>documentEnd</c> insert reuses it rather than leaving it behind as a
/// leading empty line, see <see cref="LocationResolver"/>) and is wired into the FIRST section's
/// <c>w:sectPr</c> (later sections inherit it — see <see cref="AddSectionReference"/>) via a
/// <c>w:headerReference</c>/<c>w:footerReference</c> of type <c>default</c> — without that reference the
/// part exists in the package but no page ever renders it. Both reference elements belong to
/// <c>CT_SectPr</c>'s LEADING choice group (before <c>w:pgSz</c>/<c>w:pgMar</c>), which is why they are
/// inserted at position 0 rather than appended. Deliberately NOT modelled on the corpus's three-part
/// first/even/default trio: a scaffold is the minimum that makes <c>layoutPart='header'/'footer'</c>
/// resolvable, and a caller wanting a distinct first-page header can still start from a template shell.
/// </remarks>
public static class HeaderFooterScaffold
{
    /// <summary>
    /// Ensures <paramref name="doc"/> has at least one part of the kind <paramref name="part"/> names,
    /// adding an empty one (wired into the first section's <c>sectPr</c>) when it has none. Returns
    /// <c>true</c> only when a part was actually added — <see cref="LayoutPart.Body"/>, and any layout that
    /// already has a part of that kind, are no-ops returning <c>false</c>.
    /// </summary>
    /// <exception cref="InvalidDataException"><paramref name="doc"/> has no main document part or no body.</exception>
    public static bool EnsureExists(WordprocessingDocument doc, LayoutPart part)
    {
        ArgumentNullException.ThrowIfNull(doc);

        var main = doc.MainDocumentPart
            ?? throw new InvalidDataException("Layout has no main document part.");

        return part switch
        {
            LayoutPart.Header => EnsureHeader(main),
            LayoutPart.Footer => EnsureFooter(main),
            _ => false,
        };
    }

    /// <summary>Adds an empty default header part when <paramref name="main"/> has none; see <see cref="EnsureExists"/>.</summary>
    public static bool EnsureHeader(MainDocumentPart main)
    {
        ArgumentNullException.ThrowIfNull(main);
        if (main.HeaderParts.Any())
        {
            return false;
        }

        var headerPart = main.AddNewPart<HeaderPart>();
        headerPart.Header = new Header(new Paragraph());
        headerPart.Header.Save();

        AddSectionReference(
            main, new HeaderReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(headerPart) });
        return true;
    }

    /// <summary>Adds an empty default footer part when <paramref name="main"/> has none; see <see cref="EnsureExists"/>.</summary>
    public static bool EnsureFooter(MainDocumentPart main)
    {
        ArgumentNullException.ThrowIfNull(main);
        if (main.FooterParts.Any())
        {
            return false;
        }

        var footerPart = main.AddNewPart<FooterPart>();
        footerPart.Footer = new Footer(new Paragraph());
        footerPart.Footer.Save();

        AddSectionReference(
            main, new FooterReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(footerPart) });
        return true;
    }

    /// <summary>
    /// Wires <paramref name="reference"/> into the FIRST <c>w:sectPr</c> in document order as its first child
    /// (see this type's remarks on <c>CT_SectPr</c>'s leading choice group), adding a bare body-level
    /// <c>w:sectPr</c> when the document has none at all — a layout built by <see cref="LayoutBuilder"/>
    /// always has one carrying the corpus page setup, so that fallback only ever fires for an externally
    /// authored document.
    /// </summary>
    /// <remarks>
    /// WHY THE FIRST, NOT THE BODY'S TRAILING ONE (GitHub issue #91): a section with no reference of a given
    /// type INHERITS it from the PREVIOUS section — never from a later one. A paragraph-level <c>w:sectPr</c>
    /// ends the section it sits in, so in a multi-section layout the body's trailing <c>w:sectPr</c> governs
    /// only the LAST section; a reference placed there left every earlier section, page 1 included, with no
    /// header/footer at all while the tool reported it "wired into the page setup". The first one is the
    /// section every later one inherits from, the one <see cref="HeaderFooterPartRoles"/> and
    /// <see cref="LocationResolver"/>'s default-part rule already treat as governing, and where JobQuote —
    /// the only multi-section stock layout — keeps its own references. In a single-section layout (every
    /// other corpus layout, and every <see cref="LayoutBuilder"/> output) the first IS the trailing one, so
    /// nothing changes there. <c>Descendants</c> — like those two readers — because a paragraph-level
    /// <c>w:sectPr</c> lives in a <c>w:pPr</c>, not directly under <c>w:body</c>.
    /// </remarks>
    private static void AddSectionReference(MainDocumentPart main, OpenXmlElement reference)
    {
        var body = main.Document?.Body
            ?? throw new InvalidDataException("Layout has no document body.");

        var sectPr = body.Descendants<SectionProperties>().FirstOrDefault();
        if (sectPr is null)
        {
            sectPr = new SectionProperties();
            body.AppendChild(sectPr);
        }

        sectPr.InsertAt(reference, 0);
    }
}
