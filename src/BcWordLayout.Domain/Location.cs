namespace BcWordLayout.Domain;

/// <summary>Discriminates the four insertion-point kinds a <see cref="Location"/> can describe.</summary>
public enum LocationKind
{
    /// <summary>End of the main document body, immediately before the trailing section properties.</summary>
    DocumentEnd,

    /// <summary>Immediately after an existing control, identified by its <c>w:id</c>.</summary>
    AfterControl,

    /// <summary>Inside a specific table cell, addressed by 0-based table/row/column index.</summary>
    TableCell,

    /// <summary>At the paragraph containing the first run whose text contains a search string.</summary>
    AtText,
}

/// <summary>
/// Discriminates WHICH OOXML part of the layout a <see cref="Location"/> resolves its <see cref="LocationKind"/>
/// within — the main document body (default), or a header/footer part (see <see cref="Location.PartName"/>
/// to pick a specific one). Every <see cref="LocationKind"/> is supported against every <see cref="LayoutPart"/>
/// value: e.g. <see cref="LocationKind.DocumentEnd"/> against <see cref="LayoutPart.Header"/>/
/// <see cref="LayoutPart.Footer"/> appends at the end of that part's own <c>w:hdr</c>/<c>w:ftr</c> content
/// (a header/footer never has a trailing <c>w:sectPr</c> the way the body does, so there is nothing to
/// insert before).
/// </summary>
public enum LayoutPart
{
    /// <summary>The main document body (<c>document.xml</c>). Default.</summary>
    Body,

    /// <summary>A header part (e.g. <c>header1.xml</c>).</summary>
    Header,

    /// <summary>A footer part (e.g. <c>footer1.xml</c>).</summary>
    Footer,
}

/// <summary>
/// Describes WHERE in a layout a new control should be inserted: which <see cref="LocationKind"/> within
/// which <see cref="LayoutPart"/> (the main body by default). Exactly one <see cref="Type"/> applies at a
/// time; which of <see cref="ControlId"/>, <see cref="TableIndex"/>, <see cref="Row"/>, <see cref="Col"/>
/// and <see cref="SearchText"/> are required depends on it — see <see cref="Validate"/>. This type only
/// captures the addressing INTENT and does no I/O; resolving it against an actual open document is
/// <see cref="LocationResolver"/>'s job.
/// </summary>
public sealed class Location
{
    public required LocationKind Type { get; init; }

    /// <summary>
    /// Which OOXML part <see cref="Type"/> resolves within. Defaults to <see cref="LayoutPart.Body"/> (the
    /// main document), which keeps every pre-existing caller's behavior unchanged.
    /// </summary>
    public LayoutPart Part { get; init; } = LayoutPart.Body;

    /// <summary>
    /// Optional specific header/footer part file name (e.g. <c>"header2.xml"</c>), matched case-
    /// insensitively. Only meaningful when <see cref="Part"/> is <see cref="LayoutPart.Header"/> or
    /// <see cref="LayoutPart.Footer"/> (silently ignored otherwise, same as every other kind-specific field
    /// on this type — see <see cref="Validate"/>'s own remarks). When null, the FIRST header/footer part
    /// (in the document's own part-collection order) is used. A layout with no header/footer parts at all,
    /// or none matching this name, surfaces as a <see cref="NotFoundException"/> from
    /// <see cref="LocationResolver.Resolve"/> — the same "does this actually resolve against THIS document"
    /// category as an unknown <see cref="ControlId"/> or an out-of-range <see cref="TableIndex"/>, so there
    /// is nothing to check in <see cref="Validate"/> itself.
    /// </summary>
    public string? PartName { get; init; }

    /// <summary>Required for <see cref="LocationKind.AfterControl"/>: the target control's <c>w:id</c>.</summary>
    public int? ControlId { get; init; }

    /// <summary>Required for <see cref="LocationKind.TableCell"/>: 0-based index of the table in the body.</summary>
    public int? TableIndex { get; init; }

    /// <summary>Required for <see cref="LocationKind.TableCell"/>: 0-based row index within the table.</summary>
    public int? Row { get; init; }

    /// <summary>Required for <see cref="LocationKind.TableCell"/>: 0-based cell (column) index within the row.</summary>
    public int? Col { get; init; }

    /// <summary>Required for <see cref="LocationKind.AtText"/>: substring to search for in run text (ordinal match).</summary>
    public string? SearchText { get; init; }

    /// <summary>
    /// Checks that the fields required by <see cref="Type"/> are present and well-formed, throwing
    /// <see cref="ArgumentException"/> with a specific, actionable message otherwise. This only validates
    /// the shape of the request — whether it actually resolves against a given document (an unknown
    /// control id, an out-of-range table cell, text that isn't present) is surfaced separately by
    /// <see cref="LocationResolver"/> as a <see cref="NotFoundException"/>.
    /// </summary>
    public void Validate()
    {
        switch (Type)
        {
            case LocationKind.DocumentEnd:
                break;

            case LocationKind.AfterControl:
                if (ControlId is null)
                {
                    throw new ArgumentException(
                        $"Location.{nameof(ControlId)} is required when Type is {LocationKind.AfterControl}.",
                        nameof(ControlId));
                }

                break;

            case LocationKind.TableCell:
                if (TableIndex is null || Row is null || Col is null)
                {
                    throw new ArgumentException(
                        $"Location.{nameof(TableIndex)}, {nameof(Row)} and {nameof(Col)} are all required when Type "
                        + $"is {LocationKind.TableCell}.",
                        nameof(TableIndex));
                }

                if (TableIndex < 0 || Row < 0 || Col < 0)
                {
                    throw new ArgumentException(
                        $"Location.{nameof(TableIndex)}, {nameof(Row)} and {nameof(Col)} must be non-negative "
                        + $"(got {TableIndex}, {Row}, {Col}).",
                        nameof(TableIndex));
                }

                break;

            case LocationKind.AtText:
                if (string.IsNullOrEmpty(SearchText))
                {
                    throw new ArgumentException(
                        $"Location.{nameof(SearchText)} is required (non-empty) when Type is {LocationKind.AtText}.",
                        nameof(SearchText));
                }

                break;

            default:
                throw new ArgumentException($"Unknown location kind '{Type}'.", nameof(Type));
        }
    }
}
