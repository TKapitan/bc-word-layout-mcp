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
/// preserved: a set/clear only ever replaces the text runs of the cell's own paragraphs, never removes the
/// <c>w:tc</c> or any non-text content in it (a nested table, a picture, a field). Deleting or collapsing a whole column (its cell in every row + the <c>w:tblGrid</c>) is a
/// TABLE-STRUCTURE change, deliberately out of scope here (see <see cref="TableStructureEditor"/>).
/// </remarks>
public static class CellTextEditor
{
    /// <summary>
    /// Replaces the plain text of the cell addressed by <paramref name="location"/> with
    /// <paramref name="text"/> (a single run): the cell's text runs are replaced and the paragraphs they
    /// emptied collapse into the one holding the new run, while any non-text content (a nested table, a
    /// picture, a field) stays in place — see <see cref="ReplaceCellText"/>. The cell's <c>w:tcPr</c>
    /// (width, span, borders, vertical alignment) is preserved, and the new run sits in the paragraph the old
    /// text was in, keeping its <c>w:pPr</c> (e.g. <c>pStyle</c>), and inherits the first replaced run's
    /// <c>w:rPr</c>, so a re-labelled header keeps its styling. The optional <paramref name="format"/> goes on
    /// the new run over that (see <see cref="RunFormatter"/>), and <paramref name="alignment"/>
    /// (<c>left</c>/<c>center</c>/<c>right</c>) on the paragraph holding it; omitted, the cell's look is kept.
    /// </summary>
    /// <exception cref="ArgumentException">
    /// <paramref name="location"/> is not a <see cref="LocationKind.TableCell"/> (or is otherwise
    /// structurally invalid); or <paramref name="text"/> is null; or the cell contains a content control
    /// (see remarks); or <paramref name="format"/> or <paramref name="alignment"/> is malformed.
    /// </exception>
    /// <exception cref="NotFoundException">
    /// <paramref name="location"/> does not resolve against <paramref name="doc"/> (out-of-range
    /// table/row/cell) — propagated from <see cref="LocationResolver.ResolveCellElement"/>.
    /// </exception>
    public static CellEditResult SetCellText(
        WordprocessingDocument doc, Location location, string text, TextFormat? format = null, string? alignment = null)
    {
        ArgumentNullException.ThrowIfNull(doc);
        ArgumentNullException.ThrowIfNull(location);
        ArgumentNullException.ThrowIfNull(text);

        var prepared = RunFormatter.Prepare(format, doc);
        JustificationValues? justification = alignment is null ? null : ParseAlignment(alignment, nameof(alignment));
        return Edit(doc, location, "set_cell_text", text, prepared, justification);
    }

    /// <summary>
    /// Clears all plain text from the cell addressed by <paramref name="location"/>, leaving a valid cell
    /// (its <c>w:tcPr</c> kept, the paragraphs its text emptied collapsed to one that keeps its <c>w:pPr</c>).
    /// Only the text runs go: a nested table, picture or field in the cell stays in place (see
    /// <see cref="ReplaceCellText"/>). The cell — and its table column — is preserved.
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
    /// into its OOXML justification. Shared by <see cref="SetCellText"/>, <c>set_text_format</c> and
    /// <see cref="LayoutEditor.InsertPlainTable"/>'s per-column alignments.
    /// </summary>
    /// <exception cref="ArgumentException">Anything else.</exception>
    internal static JustificationValues ParseAlignment(string alignment, string paramName, string? what = null)
    {
        return alignment.Trim().ToLowerInvariant() switch
        {
            "left" => JustificationValues.Left,
            "center" => JustificationValues.Center,
            "right" => JustificationValues.Right,
            _ => throw new ArgumentException(
                $"Unknown alignment '{alignment}'{(what is null ? string.Empty : $" for {what}")}; use 'left', "
                + "'center', or 'right'.",
                paramName),
        };
    }

    /// <summary>
    /// Sets <paramref name="paragraph"/>'s alignment (<c>w:pPr/w:jc</c>), creating its <c>w:pPr</c> when absent;
    /// the typed setter places <c>w:jc</c> at its schema-ordered position.
    /// </summary>
    internal static void Align(Paragraph paragraph, JustificationValues justification)
    {
        paragraph.ParagraphProperties ??= new ParagraphProperties();
        paragraph.ParagraphProperties.Justification = new Justification { Val = justification };
    }

    private static CellEditResult Edit(
        WordprocessingDocument doc,
        Location location,
        string operation,
        string? text,
        RunFormatter.Prepared? format = null,
        JustificationValues? alignment = null)
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

        var (previousText, kept) = ReplaceCellText(cell, text, format, alignment);

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
            Summary = BuildSummary(operation, location, partName, previousText, newText, kept),
        };
    }

    /// <summary>
    /// Replaces ONLY the cell's text runs (see <see cref="FindTextRuns"/>) and leaves everything else in the
    /// cell where it is — a nested table, a picture, a field, a bookmark. For a SET
    /// (<paramref name="text"/> non-null) one new run, carrying the first replaced run's <c>w:rPr</c>, takes
    /// the first replaced run's place (appended to the cell's first paragraph when the cell had no text);
    /// a CLEAR (<paramref name="text"/> null) only removes. A paragraph left empty by the removal collapses
    /// into the one that holds the new text, as the multi-paragraph header labels the tools were built for
    /// always did, except where it is still needed: <c>w:tc</c> must end with a paragraph, so one following a
    /// nested table stays. Returns the text of the replaced runs only, plus a description of the non-text
    /// content kept in place for the summary.
    /// </summary>
    /// <remarks>
    /// This used to rebuild the cell from its <c>w:tcPr</c> and first paragraph's <c>w:pPr</c>, which
    /// silently deleted any other content: on <c>StandardPurchaseOrder.docx</c>'s default footer, whose
    /// outer cell holds a nested 1×1 table, a clear removed that table and shifted every later
    /// <c>tableIndex</c> in the part, while the tools promise that only the text is removed (GitHub issue #86).
    /// Replacing in place keeps that promise without refusing such cells: the corpus's only nested-table cell
    /// and its only picture cell are both legitimate targets for a caption.
    /// </remarks>
    private static (string PreviousText, List<string> Kept) ReplaceCellText(
        TableCell cell, string? text, RunFormatter.Prepared? format, JustificationValues? alignment)
    {
        var paragraphs = cell.Elements<Paragraph>().ToList();
        var textRuns = FindTextRuns(paragraphs);
        var previousText = string.Concat(textRuns.SelectMany(r => r.Elements<Text>()).Select(t => t.Text));
        var host = textRuns.FirstOrDefault()?.Parent as Paragraph ?? paragraphs.FirstOrDefault();

        Run? newRun = null;
        if (text is not null)
        {
            newRun = new Run();
            if (textRuns.FirstOrDefault()?.RunProperties?.CloneNode(true) is { } rPr)
            {
                newRun.AppendChild(rPr);
            }

            newRun.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });

            if (textRuns.Count > 0)
            {
                textRuns[0].InsertBeforeSelf(newRun);
            }
            else if (host is not null)
            {
                host.AppendChild(newRun);
            }
            else
            {
                // No paragraph at all (not schema-valid, but nothing here should make it worse).
                host = cell.AppendChild(new Paragraph(newRun));
            }
        }

        var emptied = textRuns.Select(r => (Paragraph)r.Parent!).Distinct().ToList();
        foreach (var run in textRuns)
        {
            run.Remove();
        }

        foreach (var paragraph in emptied)
        {
            // Spell-check marks bracket the text just removed; with it gone they mark nothing.
            foreach (var proofError in paragraph.Elements<ProofError>().ToList())
            {
                proofError.Remove();
            }

            if (ReferenceEquals(paragraph, host)
                || paragraph.ChildElements.Any(c => c is not ParagraphProperties)
                || IsRequiredTrailingParagraph(paragraph))
            {
                continue;
            }

            paragraph.Remove();
        }

        if (host is not null && alignment is { } justification)
        {
            Align(host, justification);
        }

        // On the new run only, never on a picture or field run kept beside it.
        if (newRun is not null && format is not null)
        {
            newRun.RunProperties ??= new RunProperties();
            RunFormatter.Apply(newRun.RunProperties, format);
        }

        return (previousText, DescribeKeptContent(cell));
    }

    /// <summary>
    /// The runs of the cell's own paragraphs that hold nothing but text (<c>w:t</c>, tabs, breaks and the
    /// like) — what "the cell's text" means here. Runs inside a nested table belong to that table and are
    /// never reached (only the cell's DIRECT paragraphs are walked). A run that carries a picture, a symbol
    /// or any other non-text child is kept, and so is every run of a complex field, from its
    /// <c>fldChar begin</c> to its <c>end</c>: the field's cached result is <c>w:t</c> too, and replacing it
    /// would leave a field whose result is the caller's caption. Field depth is tracked across paragraphs
    /// because a field may span them. Runs wrapped in another element (a hyperlink, a simple field, a
    /// tracked change) are not direct paragraph children and are kept the same way.
    /// </summary>
    private static List<Run> FindTextRuns(IEnumerable<Paragraph> paragraphs)
    {
        var textRuns = new List<Run>();
        var fieldDepth = 0;
        foreach (var run in paragraphs.SelectMany(p => p.Elements<Run>()))
        {
            var fieldChars = run.Elements<FieldChar>().ToList();
            if (fieldChars.Count > 0)
            {
                foreach (var fieldChar in fieldChars)
                {
                    if (fieldChar.FieldCharType?.Value == FieldCharValues.Begin)
                    {
                        fieldDepth++;
                    }
                    else if (fieldChar.FieldCharType?.Value == FieldCharValues.End && fieldDepth > 0)
                    {
                        fieldDepth--;
                    }
                }

                continue;
            }

            if (fieldDepth == 0 && run.ChildElements.All(IsTextContent))
            {
                textRuns.Add(run);
            }
        }

        return textRuns;
    }

    private static bool IsTextContent(OpenXmlElement runChild) =>
        runChild is RunProperties or Text or TabChar or Break or CarriageReturn or NoBreakHyphen or SoftHyphen
            or LastRenderedPageBreak;

    /// <summary>
    /// True when <paramref name="paragraph"/> is the cell's last child and the block before it is not a
    /// paragraph (a nested table): <c>w:tc</c> must end with a <c>w:p</c>, so that paragraph has to stay.
    /// </summary>
    private static bool IsRequiredTrailingParagraph(Paragraph paragraph) =>
        paragraph.NextSibling() is null
        && paragraph.ElementsBefore().LastOrDefault(e => e is Paragraph or Table or SdtBlock) is not Paragraph;

    /// <summary>
    /// Phrases for the non-text content still in the cell after the edit, for the summary to name, so the
    /// caller knows the cell was not reduced to the new text alone. Empty for an ordinary text cell.
    /// </summary>
    private static List<string> DescribeKeptContent(TableCell cell)
    {
        var kept = new List<string>();
        var ownParagraphs = cell.Elements<Paragraph>().ToList();

        var tables = cell.Elements<Table>().Count();
        if (tables > 0)
        {
            kept.Add(tables == 1 ? "a nested table" : $"{tables} nested tables");
        }

        var pictures = ownParagraphs.SelectMany(p => p.Descendants<Run>())
            .Count(r => r.Elements<Drawing>().Any() || r.Elements<Picture>().Any() || r.Elements<EmbeddedObject>().Any());
        if (pictures > 0)
        {
            kept.Add(pictures == 1 ? "a picture" : $"{pictures} pictures");
        }

        var fields = ownParagraphs.Sum(p =>
            p.Descendants<FieldChar>().Count(f => f.FieldCharType?.Value == FieldCharValues.Begin)
            + p.Descendants<SimpleField>().Count());
        if (fields > 0)
        {
            kept.Add(fields == 1 ? "a field" : $"{fields} fields");
        }

        if (ownParagraphs.SelectMany(p => p.ChildElements)
            .Any(c => c is not (ParagraphProperties or Run or ProofError or BookmarkStart or BookmarkEnd or SimpleField)))
        {
            kept.Add("other inline content (e.g. a hyperlink)");
        }

        return kept;
    }

    private static string BuildSummary(
        string operation, Location location, string partName, string previousText, string newText,
        List<string> kept)
    {
        var where = $"table {location.TableIndex} row {location.Row} col {location.Col}"
            + (partName == "document.xml" ? string.Empty : $" in {partName}");

        var summary = operation == "clear_cell_text"
            ? $"Cleared the plain text of {where} (was '{previousText}'); the cell and its column are preserved."
            : $"Set the plain text of {where} to '{newText}' (was '{previousText}').";

        return kept.Count == 0
            ? summary
            : $"{summary} Only the text was replaced; the cell's other content was kept in place: "
                + $"{string.Join(", ", kept)}.";
    }
}
