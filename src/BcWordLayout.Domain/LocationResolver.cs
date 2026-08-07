using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Domain;

/// <summary>
/// An insertion point resolved from a <see cref="Location"/> against a live document. Exposes two
/// positioning primitives — a block container (an element whose direct children are paragraphs/tables/
/// block-level sdt, e.g. <see cref="Body"/> or a <see cref="TableCell"/>) plus an optional reference
/// sibling to insert before (append at the end when null) — and, only when the anchor sits inside an
/// existing paragraph, that paragraph plus an optional reference run/sdt to insert after. Every
/// <see cref="InsertionAnchor"/> that <see cref="LocationResolver.Resolve"/> actually returns is safe to
/// insert into via either method below without the caller reasoning about OOXML's paragraph/run vs
/// body/table/row/cell nesting rules — <see cref="LocationResolver"/> itself rejects (with a
/// <see cref="NotFoundException"/> carrying <see cref="NotFoundTarget.AfterControlPosition"/>) any
/// <see cref="Location"/> that has no such safe anchor, e.g. <see cref="LocationKind.AfterControl"/>
/// targeting a row-level control, rather than ever handing back an anchor that would let an insert produce
/// invalid OOXML.
/// </summary>
public sealed class InsertionAnchor
{
    private readonly OpenXmlElement _blockContainer;
    private readonly OpenXmlElement? _insertBeforeBlock;
    private readonly Paragraph? _inlineParagraph;
    private readonly OpenXmlElement? _insertAfterInline;

    internal InsertionAnchor(
        OpenXmlElement blockContainer,
        OpenXmlElement? insertBeforeBlock,
        Paragraph? inlineParagraph,
        OpenXmlElement? insertAfterInline,
        string partName)
    {
        _blockContainer = blockContainer;
        _insertBeforeBlock = insertBeforeBlock;
        _inlineParagraph = inlineParagraph;
        _insertAfterInline = insertAfterInline;
        PartName = partName;
    }

    /// <summary>
    /// The OOXML part file name this anchor's content lands in (e.g. <c>document.xml</c>,
    /// <c>header1.xml</c>) — what <see cref="Location.Part"/>/<see cref="Location.PartName"/> resolved to.
    /// Callers (e.g. <see cref="LayoutEditor"/>) use this to report exactly where an edit landed instead of
    /// assuming it is always <c>document.xml</c>.
    /// </summary>
    public string PartName { get; }

    /// <summary>
    /// The <see cref="OpenXmlPart"/> this anchor's content lands in — the package-level counterpart of
    /// <see cref="PartName"/>, for the one caller that needs the part OBJECT rather than its file name:
    /// <see cref="LayoutEditor.InsertPicture"/>, which must add its <see cref="ImagePart"/> to this exact
    /// part (a relationship id only resolves within the part that owns it). Derived from the anchor's own
    /// container by walking to its part root, so it can never disagree with where the content actually goes.
    /// </summary>
    /// <exception cref="InvalidOperationException">
    /// The anchor's container is not attached to a part — impossible for an anchor
    /// <see cref="LocationResolver.Resolve"/> produced (it always resolves within a real, open part), so
    /// this is an internal-invariant failure rather than a caller error.
    /// </exception>
    public OpenXmlPart HostPart(WordprocessingDocument doc)
    {
        ArgumentNullException.ThrowIfNull(doc);

        var root = _blockContainer as OpenXmlPartRootElement
            ?? _blockContainer.Ancestors<OpenXmlPartRootElement>().FirstOrDefault();

        return root?.OpenXmlPart
            ?? throw new InvalidOperationException(
                $"The insertion anchor in '{PartName}' is not attached to an open part (unexpected).");
    }

    /// <summary>
    /// Inserts a block-level element (e.g. a <see cref="Table"/>, or a repeater <c>SdtBlock</c>) at this
    /// anchor: a direct child of the resolved block container, positioned immediately before the
    /// reference sibling recorded when the anchor was resolved (appended at the end of the container
    /// when there was none — e.g. <see cref="LocationKind.DocumentEnd"/> with no trailing section
    /// properties).
    /// </summary>
    /// <remarks>
    /// When <paramref name="blockContent"/> is a <see cref="Table"/>, this also guarantees a <see cref="Paragraph"/>
    /// immediately follows it (inserting an empty one when none already does — e.g. a <see cref="TableCell"/>
    /// anchor already ends on its own trailing paragraph, so nothing is added there). Every real corpus body
    /// has a paragraph after each top-level table; without one, two adjacent <c>w:tbl</c> elements (e.g. two
    /// repeater tables both inserted at <see cref="LocationKind.DocumentEnd"/>) get silently MERGED into a
    /// single table by Word on load, corrupting both.
    /// </remarks>
    public void InsertBlock(OpenXmlElement blockContent)
    {
        ArgumentNullException.ThrowIfNull(blockContent);

        if (_insertBeforeBlock is not null)
        {
            _blockContainer.InsertBefore(blockContent, _insertBeforeBlock);
        }
        else
        {
            _blockContainer.AppendChild(blockContent);
        }

        if (blockContent is Table && blockContent.NextSibling() is not Paragraph)
        {
            blockContent.InsertAfterSelf(new Paragraph());
        }
    }

    /// <summary>
    /// Inserts an inline (run-level) element — e.g. a field/label <c>SdtRun</c> built by
    /// <see cref="SdtFactory"/> — at this anchor. When the anchor sits inside an existing paragraph
    /// (<see cref="LocationKind.AfterControl"/> on an inline control, <see cref="LocationKind.TableCell"/>,
    /// or <see cref="LocationKind.AtText"/>), the content is placed directly in that paragraph's run
    /// flow — immediately after the reference run/control when one is known, otherwise appended at the
    /// paragraph's end. Otherwise (the anchor is purely block-level, e.g.
    /// <see cref="LocationKind.DocumentEnd"/>, or <see cref="LocationKind.AfterControl"/> on a
    /// block-level control) a brand new paragraph is created to host the content and inserted per the
    /// same rule <see cref="InsertBlock"/> uses.
    /// </summary>
    public void InsertInline(OpenXmlElement inlineContent)
    {
        ArgumentNullException.ThrowIfNull(inlineContent);

        if (_inlineParagraph is not null)
        {
            if (_insertAfterInline is not null)
            {
                _insertAfterInline.InsertAfterSelf(inlineContent);
            }
            else
            {
                _inlineParagraph.AppendChild(inlineContent);
            }

            return;
        }

        InsertBlock(new Paragraph(inlineContent));
    }
}

/// <summary>
/// Resolves a <see cref="Location"/> against a chosen part of a <see cref="WordprocessingDocument"/> — the
/// main document body by default, or a specific (or the first) header/footer part when
/// <see cref="Location.Part"/>/<see cref="Location.PartName"/> says so — into an
/// <see cref="InsertionAnchor"/>.
/// </summary>
public static class LocationResolver
{
    /// <summary>
    /// Resolves <paramref name="location"/> (after calling its own <see cref="Location.Validate"/>)
    /// against the part of <paramref name="doc"/> named by <see cref="Location.Part"/>/
    /// <see cref="Location.PartName"/>. Throws <see cref="ArgumentException"/> for a structurally-invalid
    /// <paramref name="location"/> (see <see cref="Location.Validate"/>), or a <see cref="NotFoundException"/>
    /// when a well-formed location does not actually resolve against this document — unknown control id
    /// (<see cref="NotFoundTarget.Control"/>), out-of-range table cell (<see cref="NotFoundTarget.TableCoordinate"/>),
    /// text not found (<see cref="NotFoundTarget.SearchText"/>), WITHIN THE CHOSEN PART (a control that
    /// exists only in a different part is exactly as "not found" as one that does not exist at all); a
    /// layout with no header/footer parts at all (<see cref="NotFoundTarget.HeaderFooterParts"/>), or none
    /// matching <see cref="Location.PartName"/> (<see cref="NotFoundTarget.NamedHeaderFooterPart"/>), is
    /// reported the same way. Also includes <see cref="LocationKind.AfterControl"/> naming a control whose
    /// parent cannot safely host a sibling (<see cref="NotFoundTarget.AfterControlPosition"/> — a row-level
    /// control, whose parent is a <see cref="Table"/>, or any other unrecognized nesting); a cell-level
    /// control (parent <see cref="TableRow"/>) is supported by anchoring inside its own cell. A small number
    /// of "should be impossible" structural invariants (e.g. a paragraph with no parent) still throw a plain
    /// <see cref="InvalidOperationException"/> instead — see the throw sites' own comments.
    /// </summary>
    public static InsertionAnchor Resolve(Location location, WordprocessingDocument doc)
    {
        ArgumentNullException.ThrowIfNull(location);
        ArgumentNullException.ThrowIfNull(doc);
        location.Validate();

        var (root, partName) = ResolvePartRoot(location, doc);
        var partDescription = DescribePart(partName);

        return location.Type switch
        {
            LocationKind.DocumentEnd => ResolveDocumentEnd(root, partName),
            LocationKind.AfterControl => ResolveAfterControl(root, location.ControlId!.Value, partName, partDescription),
            LocationKind.TableCell => ResolveTableCell(
                root, location.TableIndex!.Value, location.Row!.Value, location.Col!.Value, partName, partDescription),
            LocationKind.AtText => ResolveAtText(root, location.SearchText!, partName, partDescription),
            _ => throw new ArgumentException($"Unsupported location kind '{location.Type}'.", nameof(location)),
        };
    }

    /// <summary>
    /// Resolves the block-content root of the part named by <paramref name="part"/>/<paramref name="partName"/>
    /// (the main body by default, else a specific/first header/footer part), plus that part's file name —
    /// for callers that address a whole table/column rather than an insertion point (e.g.
    /// <see cref="TableStructureEditor"/>). Uses the SAME part-selection rules and "not found" messages as
    /// <see cref="Resolve"/>, so error hints and behaviour stay identical across the tool surface.
    /// </summary>
    public static (OpenXmlElement Root, string PartName) ResolvePart(WordprocessingDocument doc, LayoutPart part, string? partName)
    {
        ArgumentNullException.ThrowIfNull(doc);
        return ResolvePartRoot(new Location { Type = LocationKind.DocumentEnd, Part = part, PartName = partName }, doc);
    }

    /// <summary>Human-readable phrase for "not found" messages (e.g. "the document body"/"part 'header1.xml'").</summary>
    public static string DescribePart(string partName) =>
        partName == DocumentPartName ? "the document body" : $"part '{partName}'";

    // ---- part selection ----

    private const string DocumentPartName = "document.xml";

    /// <summary>
    /// Picks the root block-content element <see cref="Location.Type"/> resolves within — the main body,
    /// or a specific/first header/footer part — plus that part's own file name (e.g. <c>header1.xml</c>,
    /// reused both for <see cref="InsertionAnchor.PartName"/> and every "not found" message below).
    /// </summary>
    private static (OpenXmlElement Root, string PartName) ResolvePartRoot(Location location, WordprocessingDocument doc)
    {
        var main = doc.MainDocumentPart
            ?? throw new InvalidDataException("Layout has no main document part.");

        return location.Part switch
        {
            LayoutPart.Body => (
                main.Document?.Body ?? throw new InvalidDataException("Layout has no main document body."),
                DocumentPartName),

            LayoutPart.Header => ResolveHeaderOrFooterRoot(
                main.HeaderParts.Select(h => ((OpenXmlPart)h, (OpenXmlElement?)h.Header)), location.PartName, "header",
                DefaultReferencedPart(main, header: true)),

            LayoutPart.Footer => ResolveHeaderOrFooterRoot(
                main.FooterParts.Select(f => ((OpenXmlPart)f, (OpenXmlElement?)f.Footer)), location.PartName, "footer",
                DefaultReferencedPart(main, header: false)),

            _ => throw new ArgumentException($"Unsupported layout part '{location.Part}'.", nameof(location)),
        };
    }

    /// <summary>
    /// The header/footer part the FIRST section references as its <c>default</c> (primary) one, or null when
    /// the layout declares no such reference.
    /// </summary>
    /// <remarks>
    /// This is what "the header" means to a person, and it is emphatically NOT the first part in the package.
    /// A Word document with distinct first-page/even-page headers carries three header parts, and their
    /// relationship order has nothing to do with which one is the everyday header: measured 2026-08-01 across
    /// the corpus, FIVE of the six layouts that have header/footer parts had a different first part than
    /// first-section default — on <c>StandardSalesQuote</c>, <c>StandardPurchaseOrder</c> and
    /// <c>JobQuote</c>, <c>header1.xml</c> is the EVEN-PAGE header, so a <c>layoutPart: 'header'</c> insert
    /// with no <c>partName</c> used to land on even pages only and was invisible on a one-page document.
    /// Footers were the same story via the first-page footer, which is why previews of page 1 looked correct.
    /// <para>
    /// The first <c>w:sectPr</c> in document order governs the first section (a paragraph-level
    /// <c>w:sectPr</c> ends the section it sits in), so the search takes the first section that declares any
    /// reference of the requested kind and prefers its <c>default</c>-typed one. A <c>w:type</c> that is
    /// absent IS "default" per the spec, so a missing attribute must count — and where a section declares
    /// only first/even references and no default at all, its first declared reference is the honest answer.
    /// </para>
    /// </remarks>
    private static OpenXmlPart? DefaultReferencedPart(MainDocumentPart main, bool header)
    {
        var body = main.Document?.Body;
        if (body is null)
        {
            return null;
        }

        foreach (var sectPr in body.Descendants<SectionProperties>())
        {
            // (relationship id, is this the default-typed reference?). An ABSENT w:type is "default" per the
            // spec, so a null Type must count as default rather than be skipped.
            var references = header
                ? sectPr.Elements<HeaderReference>()
                    .Select(r => (Id: r.Id?.Value, IsDefault: r.Type is null || r.Type.Value == HeaderFooterValues.Default))
                    .ToList()
                : sectPr.Elements<FooterReference>()
                    .Select(r => (Id: r.Id?.Value, IsDefault: r.Type is null || r.Type.Value == HeaderFooterValues.Default))
                    .ToList();

            if (references.Count == 0)
            {
                continue;
            }

            var chosen = references.FirstOrDefault(r => r.IsDefault);
            var relationshipId = chosen.Id ?? references[0].Id;
            if (relationshipId is null)
            {
                return null;
            }

            try
            {
                return main.GetPartById(relationshipId);
            }
            catch (ArgumentOutOfRangeException)
            {
                // A sectPr referencing a relationship the package does not contain is a broken layout, not
                // something to throw over here: fall back to the positional rule below.
                return null;
            }
        }

        return null;
    }

    /// <summary>
    /// Shared header/footer part-selection logic: the part named <paramref name="partName"/> (matched
    /// case-insensitively against its own file name); else <paramref name="preferred"/> — the first section's
    /// <c>default</c>-typed part, see <see cref="DefaultReferencedPart"/>; else the FIRST one in
    /// <paramref name="candidates"/>'s own order. <paramref name="kindLabel"/> ("header"/"footer") only feeds
    /// the exception messages below.
    /// </summary>
    private static (OpenXmlElement Root, string PartName) ResolveHeaderOrFooterRoot(
        IEnumerable<(OpenXmlPart Part, OpenXmlElement? Root)> candidates, string? partName, string kindLabel,
        OpenXmlPart? preferred = null)
    {
        var list = candidates.ToList();
        if (list.Count == 0)
        {
            throw new NotFoundException(
                $"Layout has no {kindLabel} parts; cannot resolve a location targeting {kindLabel} part "
                + $"{(partName is null ? "(unspecified, would use the default one)" : $"'{partName}'")}.",
                NotFoundTarget.HeaderFooterParts);
        }

        (OpenXmlPart Part, OpenXmlElement? Root) selected;
        if (partName is null)
        {
            selected = preferred is null
                ? list[0]
                : list.FirstOrDefault(c => ReferenceEquals(c.Part, preferred), list[0]);
        }
        else
        {
            var matches = list.Where(c => string.Equals(PartWalker.PartFileName(c.Part), partName, StringComparison.OrdinalIgnoreCase)).ToList();
            if (matches.Count == 0)
            {
                var available = string.Join(", ", list.Select(c => PartWalker.PartFileName(c.Part)));
                throw new NotFoundException(
                    $"No {kindLabel} part named '{partName}' was found; this layout's {kindLabel} parts are: {available}.",
                    NotFoundTarget.NamedHeaderFooterPart);
            }

            selected = matches[0];
        }

        var partFileName = PartWalker.PartFileName(selected.Part);
        var root = selected.Root
            ?? throw new InvalidDataException($"{kindLabel} part '{partFileName}' has no content.");

        return (root, partFileName);
    }

    // DocumentEnd: append at the end of the CHOSEN part's own content, before its trailing w:sectPr when
    // one is a direct child (Elements<T>() — unlike Descendants<T>() — only looks at direct children, so
    // this cannot be confused with a w:sectPr nested inside a paragraph's w:pPr, which marks a mid-document
    // section break instead). A header/footer never has one — HeaderFooterContent's own content model has
    // no place for a w:sectPr at all — so for those parts this always falls through to a plain append at
    // the end of the part's own content, exactly as a body with no trailing sectPr already does.
    private static InsertionAnchor ResolveDocumentEnd(OpenXmlElement root, string partName)
    {
        var sectPr = root.Elements<SectionProperties>().FirstOrDefault();
        return new InsertionAnchor(
            blockContainer: root, insertBeforeBlock: sectPr, inlineParagraph: null, insertAfterInline: null, partName: partName);
    }

    // AfterControl: locate the w:sdt (any kind — run/block/row/cell) whose w:sdtPr/w:id/@w:val matches,
    // WITHIN THE CHOSEN PART ONLY (a control living in a different part is exactly as "not found" as one
    // that doesn't exist anywhere). What "immediately after it, same parent" means safely depends on what
    // kind of parent the control actually has:
    //  - Paragraph (an inline SdtRun): keep that paragraph so InsertInline lands right after the control
    //    in the same run flow; a same-parent InsertBlock is impossible there (paragraphs cannot host
    //    block content), so block inserts are promoted to immediately after the enclosing paragraph.
    //  - TableRow (a cell-level SdtCell — e.g. the real corpus label YourReference_Lbl, whose sdtContent
    //    wraps a whole <w:tc>): a <w:tr> only ever accepts <w:tc>/an sdtCell/customXml as direct children,
    //    so "same parent" would try to insert a paragraph/block sdt straight into the row — invalid
    //    OOXML. Anchor INSIDE that control's own cell content instead (same placement rule TableCell
    //    addressing uses).
    //  - Table (a row-level SdtRow — e.g. a repeatingSection/repeatingSectionItem): a <w:tbl> only ever
    //    accepts <w:tr>/an sdtRow/customXml as direct children, and there is no well-defined "inside"
    //    fallback for a row the way there is for a cell, so this is rejected outright.
    //  - Body / Header / Footer / TableCell / SdtContentBlock (ordinary block-content hosts, all sharing
    //    the same block content model): "same parent" applies literally for both insert kinds.
    //  - anything else unrecognized: rejected outright rather than risk emitting invalid OOXML.
    private static InsertionAnchor ResolveAfterControl(OpenXmlElement root, int controlId, string partName, string partDescription)
    {
        var target = root.Descendants<SdtElement>().FirstOrDefault(sdt => SdtInspector.ReadControlId(sdt) == controlId);
        if (target is null)
        {
            throw new NotFoundException(
                $"No control with id {controlId} was found in {partDescription}.", NotFoundTarget.Control);
        }

        if (target.Parent is Paragraph paragraph)
        {
            // "Has no parent" is a structural invariant violation (every real paragraph in a saved OOXML
            // document sits inside SOME block container), not a lookup failure the caller can fix by
            // retrying with different arguments - left as a plain InvalidOperationException (→ internal_error)
            // rather than NotFoundException.
            var paragraphParent = paragraph.Parent
                ?? throw new InvalidOperationException($"Control {controlId}'s enclosing paragraph has no parent.");

            return new InsertionAnchor(
                blockContainer: paragraphParent,
                insertBeforeBlock: paragraph.NextSibling(),
                inlineParagraph: paragraph,
                insertAfterInline: target,
                partName: partName);
        }

        if (target.Parent is TableRow)
        {
            if (target is not SdtCell sdtCell)
            {
                throw new NotFoundException(
                    $"Control {controlId}'s parent is a table row, but the control is not a recognized "
                    + "cell-level control; AfterControl cannot insert a sibling there safely. Use TableCell "
                    + "addressing instead.",
                    NotFoundTarget.AfterControlPosition);
            }

            var cell = sdtCell.SdtContentCell?.GetFirstChild<TableCell>()
                ?? throw new NotFoundException(
                    $"Control {controlId} is a cell-level control but its content has no table cell to insert into.",
                    NotFoundTarget.AfterControlPosition);

            return AnchorInsideCell(cell, partName);
        }

        if (target.Parent is Table)
        {
            throw new NotFoundException(
                $"Control {controlId} is a row-level control (its parent is a table); AfterControl is not "
                + "supported for that control kind because a table cannot host a paragraph or block sdt as a "
                + "direct sibling. Use TableCell addressing instead.",
                NotFoundTarget.AfterControlPosition);
        }

        if (target.Parent is Body or Header or Footer or TableCell or SdtContentBlock)
        {
            return new InsertionAnchor(
                blockContainer: target.Parent,
                insertBeforeBlock: target.NextSibling(),
                inlineParagraph: null,
                insertAfterInline: null,
                partName: partName);
        }

        throw new NotFoundException(
            $"Control {controlId}'s parent ({target.Parent?.GetType().Name ?? "none"}) cannot safely host a "
            + "sibling control; AfterControl is not supported for this control kind. Use TableCell addressing "
            + "instead, or target a different control.",
            NotFoundTarget.AfterControlPosition);
    }

    // TableCell: the (0-based) tableIndex-th w:tbl in the CHOSEN part (document order, including nested
    // tables), its row-th row, col-th cell. Anchor lands INSIDE that cell — see AnchorInsideCell.
    //
    // Rows and cells are enumerated through TableGridNavigator — the SAME row/cell walker
    // TableStructureReader (get_layout_info) and table-structure editing use — so an index a caller reads
    // back from get_layout_info's tables[] always addresses the same physical spot here:
    //   - a "row" is a w:tr OR a row-level sdt (SdtRow, e.g. a BC repeater's repeatingSection) sitting where
    //     a w:tr would; a targeted SdtRow is descended to the inner w:tr it ultimately wraps.
    //   - a "cell" is a w:tc OR a cell-level sdt (SdtCell, e.g. a BC header address/label field, whose
    //     sdtContent wraps a whole w:tc) sitting where a w:tc would; a targeted SdtCell is descended to the
    //     inner w:tc it wraps.
    // Counting only bare w:tr/w:tc (root.Elements<TableRow>()/row.Elements<TableCell>()) would UNDER-count
    // any row/table that uses these control wrappers — e.g. a BC header row of three address fields is
    // [SdtCell, w:tc, SdtCell], which has one bare w:tc, so col 1/2 would wrongly read as out of range even
    // though get_layout_info reports three cells for that row.
    private static InsertionAnchor ResolveTableCell(
        OpenXmlElement root, int tableIndex, int row, int col, string partName, string partDescription)
    {
        var (targetCell, _) = FindTableCell(root, tableIndex, row, col, partName, partDescription);
        return AnchorInsideCell(targetCell, partName);
    }

    /// <summary>
    /// Resolves a <see cref="LocationKind.TableCell"/> <paramref name="location"/> to the physical
    /// <see cref="TableCell"/> it addresses (plus the part file name it lives in), for callers that need to
    /// mutate the cell's own content directly rather than insert INTO it (e.g. plain-text cell editing —
    /// <see cref="CellTextEditor"/>). Table/row/cell enumeration goes through <see cref="TableGridNavigator"/>
    /// — the same walker <see cref="ResolveTableCell"/>/<see cref="TableStructureReader"/> use — so an index
    /// a caller reads back from <c>get_layout_info</c>'s <c>tables[]</c> addresses the same physical cell
    /// here. Throws <see cref="ArgumentException"/> when <paramref name="location"/> is not a
    /// <see cref="LocationKind.TableCell"/> (or is otherwise structurally invalid — see
    /// <see cref="Location.Validate"/>), or a <see cref="NotFoundException"/> (<see cref="NotFoundTarget.TableCoordinate"/>)
    /// when the well-formed location does not resolve against <paramref name="doc"/> (out-of-range
    /// table/row/cell, or a cell-level control with no inner cell).
    /// </summary>
    public static (TableCell Cell, string PartName) ResolveCellElement(Location location, WordprocessingDocument doc)
    {
        ArgumentNullException.ThrowIfNull(location);
        ArgumentNullException.ThrowIfNull(doc);
        location.Validate();

        if (location.Type != LocationKind.TableCell)
        {
            throw new ArgumentException(
                $"A cell can only be resolved from a {nameof(LocationKind)}.{nameof(LocationKind.TableCell)} "
                + $"location (got {location.Type}).",
                nameof(location));
        }

        var (root, partName) = ResolvePartRoot(location, doc);
        var partDescription = DescribePart(partName);
        return FindTableCell(root, location.TableIndex!.Value, location.Row!.Value, location.Col!.Value, partName, partDescription);
    }

    private static (TableCell Cell, string PartName) FindTableCell(
        OpenXmlElement root, int tableIndex, int row, int col, string partName, string partDescription)
    {
        var table = TableGridNavigator.TableAt(root, tableIndex, partDescription);

        var rowSlots = TableGridNavigator.Rows(table);
        if (row < 0 || row >= rowSlots.Count)
        {
            throw new NotFoundException(
                $"Row index {row} is out of range for table {tableIndex}; it has {rowSlots.Count} row(s).",
                NotFoundTarget.TableCoordinate);
        }

        var innerRow = rowSlots[row].InnerRow;
        var cellSlots = innerRow is null
            ? Array.Empty<TableGridNavigator.CellSlot>()
            : TableGridNavigator.Cells(innerRow);
        if (col < 0 || col >= cellSlots.Count)
        {
            throw new NotFoundException(
                $"Column index {col} is out of range for table {tableIndex}, row {row}; it has {cellSlots.Count} cell(s).",
                NotFoundTarget.TableCoordinate);
        }

        var targetCell = cellSlots[col].InnerCell
            ?? throw new NotFoundException(
                $"Table {tableIndex}, row {row}, col {col} resolves to a cell-level control with no inner "
                + "table cell to operate on.",
                NotFoundTarget.TableCoordinate);

        return (targetCell, partName);
    }

    // Shared placement rule for "anchor inside this table cell" (used by both TableCell addressing and
    // AfterControl on a cell-level control): inline content is appended into the cell's own last
    // paragraph (every well-formed w:tc already ends with one); block content is inserted immediately
    // before that trailing paragraph so the cell still ends on a paragraph, which is what real
    // Word/BC-produced cells always do.
    private static InsertionAnchor AnchorInsideCell(TableCell cell, string partName)
    {
        var lastParagraph = cell.Elements<Paragraph>().LastOrDefault();
        var lastParagraphIsLastChild = lastParagraph is not null && ReferenceEquals(cell.LastChild, lastParagraph);

        return new InsertionAnchor(
            blockContainer: cell,
            insertBeforeBlock: lastParagraphIsLastChild ? lastParagraph : null,
            inlineParagraph: lastParagraph,
            insertAfterInline: null,
            partName: partName);
    }

    // AtText: the first w:t (document order) whose text contains searchText (ordinal), WITHIN THE CHOSEN
    // PART ONLY; anchor is the containing paragraph — inline content is appended into that paragraph,
    // block content is inserted immediately after it in its own parent.
    private static InsertionAnchor ResolveAtText(OpenXmlElement root, string searchText, string partName, string partDescription)
    {
        var match = root.Descendants<Text>()
            .FirstOrDefault(t => t.Text is not null && t.Text.Contains(searchText, StringComparison.Ordinal));
        if (match is null)
        {
            throw new NotFoundException(
                $"No text containing '{searchText}' was found in {partDescription}.", NotFoundTarget.SearchText);
        }

        // Both "has no enclosing paragraph"/"has no parent" below are structural invariant violations (a
        // w:t Word itself produced is always inside a run inside a paragraph with a real parent) rather than
        // lookup failures the caller could fix by retrying - left as plain InvalidOperationException
        // (→ internal_error), not NotFoundException.
        var paragraph = match.Ancestors<Paragraph>().FirstOrDefault()
            ?? throw new InvalidOperationException($"Text containing '{searchText}' was found but has no enclosing paragraph.");

        var paragraphParent = paragraph.Parent
            ?? throw new InvalidOperationException($"Paragraph containing '{searchText}' has no parent.");

        return new InsertionAnchor(
            blockContainer: paragraphParent,
            insertBeforeBlock: paragraph.NextSibling(),
            inlineParagraph: paragraph,
            insertAfterInline: null,
            partName: partName);
    }

}
