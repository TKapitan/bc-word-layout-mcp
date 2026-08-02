using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Domain;

/// <summary>
/// Deterministic OOXML mutations for the PLAIN-TEXT content of a table cell — setting/replacing it
/// (<see cref="SetCellText"/>) or clearing it (<see cref="ClearCellText"/>). This is the counterpart to
/// <see cref="LayoutEditor"/>, which edits content CONTROLS (bound field/label sdt) identified by
/// <c>w:id</c>: a plain-text cell (e.g. a line-items table's column-HEADER label like "GST Amount", which
/// BC authors as ordinary cell text, not a content control) has no control to target, so it is addressed
/// by its (table, row, column) coordinates instead — the very coordinates <c>get_layout_info</c> reports.
/// Like <see cref="LayoutEditor"/>, every method works against an ALREADY-OPEN document and does no file
/// I/O of its own; opening, pre-save validation, and saving are the caller's job (see
/// <c>BcWordLayout.McpHost.Tools.ToolGuards</c>'s <c>GuardCellEdit</c> flow, used by
/// <c>BcWordLayout.McpHost.Tools.EditTools</c>'s <c>set_cell_text</c>/<c>clear_cell_text</c>).
/// </summary>
/// <remarks>
/// SCOPE: these tools operate ONLY on genuinely plain-text cells. A cell that contains a content control
/// (a bound field/label — anything with a <c>w:sdt</c> inside it) is rejected up front, so a set/clear can
/// never silently clobber a binding — that is <see cref="LayoutEditor.RemoveControl"/> /
/// <c>insert_field</c>/<c>insert_label</c> territory. The cell (and thus the table column) is always
/// preserved: a set/clear only ever rewrites the cell's own paragraph content, never removes the
/// <c>w:tc</c>. Deleting or collapsing a whole column (its cell in every row + the <c>w:tblGrid</c>) is a
/// TABLE-STRUCTURE change, deliberately out of scope here (see <see cref="TableStructureEditor"/>).
/// </remarks>
public static class CellTextEditor
{
    /// <summary>
    /// Replaces the plain text of the cell addressed by <paramref name="location"/> with
    /// <paramref name="text"/> (a single run), collapsing any existing paragraphs/runs into one paragraph.
    /// The cell's <c>w:tcPr</c> (width, span, borders, vertical alignment) is preserved, and the new run
    /// inherits the first existing paragraph's <c>w:pPr</c> (e.g. <c>pStyle</c>) and first run's <c>w:rPr</c>
    /// when there was one, so a re-labelled header keeps its styling.
    /// </summary>
    /// <exception cref="ArgumentException">
    /// <paramref name="location"/> is not a <see cref="LocationKind.TableCell"/> (or is otherwise
    /// structurally invalid); or <paramref name="text"/> is null; or the cell contains a content control
    /// (see remarks).
    /// </exception>
    /// <exception cref="NotFoundException">
    /// <paramref name="location"/> does not resolve against <paramref name="doc"/> (out-of-range
    /// table/row/cell) — propagated from <see cref="LocationResolver.ResolveCellElement"/>.
    /// </exception>
    public static CellEditResult SetCellText(
        WordprocessingDocument doc, Location location, string text, CellTextFormat? format = null)
    {
        ArgumentNullException.ThrowIfNull(doc);
        ArgumentNullException.ThrowIfNull(location);
        ArgumentNullException.ThrowIfNull(text);

        return Edit(doc, location, "set_cell_text", text, format);
    }

    /// <summary>
    /// Clears all plain text from the cell addressed by <paramref name="location"/>, leaving a valid empty
    /// cell (its <c>w:tcPr</c> kept, its content collapsed to a single empty paragraph that preserves the
    /// first paragraph's <c>w:pPr</c> when there was one). The cell — and its table column — is preserved.
    /// </summary>
    /// <exception cref="ArgumentException">
    /// <paramref name="location"/> is not a <see cref="LocationKind.TableCell"/> (or is otherwise
    /// structurally invalid); or the cell contains a content control (see remarks).
    /// </exception>
    /// <exception cref="NotFoundException">
    /// <paramref name="location"/> does not resolve against <paramref name="doc"/> — propagated from
    /// <see cref="LocationResolver.ResolveCellElement"/>.
    /// </exception>
    public static CellEditResult ClearCellText(WordprocessingDocument doc, Location location)
    {
        ArgumentNullException.ThrowIfNull(doc);
        ArgumentNullException.ThrowIfNull(location);

        return Edit(doc, location, "clear_cell_text", text: null);
    }

    /// <summary>
    /// Parses a caller-facing alignment name (<c>left</c>/<c>center</c>/<c>right</c>, case-insensitive)
    /// into its OOXML justification. Shared by <see cref="SetCellText"/>'s formatting options and
    /// <see cref="LayoutEditor.InsertPlainTable"/>'s per-column alignments.
    /// </summary>
    /// <exception cref="ArgumentException">Anything else.</exception>
    internal static JustificationValues ParseAlignment(string alignment, string paramName)
    {
        return alignment.ToLowerInvariant() switch
        {
            "left" => JustificationValues.Left,
            "center" => JustificationValues.Center,
            "right" => JustificationValues.Right,
            _ => throw new ArgumentException(
                $"Unknown alignment '{alignment}'; use 'left', 'center', or 'right'.", paramName),
        };
    }

    private static CellEditResult Edit(
        WordprocessingDocument doc, Location location, string operation, string? text, CellTextFormat? format = null)
    {
        var (cell, partName) = LocationResolver.ResolveCellElement(location, doc);

        // Never clobber a bound field/label: a cell is not "plain text" if it either CONTAINS a content
        // control (a block/inline sdt inside the w:tc) or IS the content of a cell-level control (an
        // SdtCell wrapping the whole w:tc — e.g. a BC header address field; ResolveCellElement descends
        // into that control's inner w:tc, so the control is this cell's PARENT, not a descendant).
        if (cell.Descendants<SdtElement>().Any() || cell.Parent is SdtContentCell)
        {
            throw new ArgumentException(
                $"Table {location.TableIndex}, row {location.Row}, col {location.Col} contains a content "
                + "control (a bound field/label), so it is not a plain-text cell. Use remove_control to remove "
                + "that control (the cell/column is preserved either way), or insert_field/insert_label to "
                + "change its binding; set_cell_text/clear_cell_text only operate on plain-text cells.",
                nameof(location));
        }

        var previousText = string.Concat(cell.Descendants<Text>().Select(t => t.Text));

        ReplaceCellText(cell, text, format);

        var newText = text ?? string.Empty;
        return new CellEditResult
        {
            Operation = operation,
            Part = partName,
            TableIndex = location.TableIndex!.Value,
            Row = location.Row!.Value,
            Col = location.Col!.Value,
            PreviousText = previousText,
            NewText = newText,
            Summary = BuildSummary(operation, location, partName, previousText, newText),
        };
    }

    /// <summary>
    /// Rewrites <paramref name="cell"/>'s content to a single paragraph: its <c>w:tcPr</c> is kept first,
    /// then one paragraph carrying the first old paragraph's <c>w:pPr</c> (when any) plus — for a SET
    /// (<paramref name="text"/> non-null) — one run carrying the first old run's <c>w:rPr</c> (when any) and
    /// the new text. A CLEAR (<paramref name="text"/> null) leaves the paragraph run-less. A <c>w:tc</c>
    /// must contain at least one block-level child, so the single paragraph is always present, keeping the
    /// cell (and its column) well-formed.
    /// </summary>
    private static void ReplaceCellText(TableCell cell, string? text, CellTextFormat? format = null)
    {
        var tcPr = cell.GetFirstChild<TableCellProperties>();
        var firstParagraph = cell.Elements<Paragraph>().FirstOrDefault();
        var pPr = firstParagraph?.GetFirstChild<ParagraphProperties>()?.CloneNode(true);
        var rPr = firstParagraph?.Elements<Run>().FirstOrDefault()?.GetFirstChild<RunProperties>()?.CloneNode(true);

        cell.RemoveAllChildren();

        if (tcPr is not null)
        {
            cell.AppendChild(tcPr);
        }

        var paragraph = new Paragraph();
        if (pPr is not null)
        {
            paragraph.AppendChild(pPr);
        }

        if (text is not null)
        {
            var run = new Run();
            if (rPr is not null)
            {
                run.AppendChild(rPr);
            }

            run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            paragraph.AppendChild(run);
        }

        cell.AppendChild(paragraph);

        ApplyFormat(paragraph, format);
    }

    /// <summary>
    /// Applies the OPTIONAL formatting knobs of a set: every null member means "leave as inherited/kept" —
    /// the pre-format behavior exactly. The strongly-typed OpenXml property setters place each element at
    /// its schema-ordered position inside <c>w:pPr</c>/<c>w:rPr</c>, so the pre-save validator gate stays
    /// clean regardless of what the preserved properties already contain.
    /// </summary>
    private static void ApplyFormat(Paragraph paragraph, CellTextFormat? format)
    {
        if (format is null)
        {
            return;
        }

        if (format.Alignment is { } alignment)
        {
            paragraph.ParagraphProperties ??= new ParagraphProperties();
            paragraph.ParagraphProperties.Justification =
                new Justification { Val = ParseAlignment(alignment, nameof(format)) };
        }

        var run = paragraph.Elements<Run>().FirstOrDefault();
        if (run is null || (format.Bold is null && format.FontSizePoints is null))
        {
            return;
        }

        run.RunProperties ??= new RunProperties();
        if (format.Bold is { } bold)
        {
            run.RunProperties.Bold = bold ? new Bold() : null;
        }

        if (format.FontSizePoints is { } points)
        {
            if (points is < 4 or > 96)
            {
                throw new ArgumentException(
                    $"fontSizePoints must be between 4 and 96 (got {points}).", nameof(format));
            }

            // w:sz/w:szCs are half-points; set both, as Word itself does.
            var halfPoints = ((int)Math.Round(points * 2)).ToString(System.Globalization.CultureInfo.InvariantCulture);
            run.RunProperties.FontSize = new FontSize { Val = halfPoints };
            run.RunProperties.FontSizeComplexScript = new FontSizeComplexScript { Val = halfPoints };
        }
    }

    private static string BuildSummary(
        string operation, Location location, string partName, string previousText, string newText)
    {
        var where = $"table {location.TableIndex} row {location.Row} col {location.Col}"
            + (partName == "document.xml" ? string.Empty : $" in {partName}");

        return operation == "clear_cell_text"
            ? $"Cleared the plain text of {where} (was '{previousText}'); the cell and its column are preserved."
            : $"Set the plain text of {where} to '{newText}' (was '{previousText}').";
    }
}
