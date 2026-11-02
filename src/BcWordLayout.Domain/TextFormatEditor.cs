using System.Globalization;
using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Domain;

/// <summary>
/// <c>set_text_format</c>: restyles text that is ALREADY in the layout (GitHub issue #15). The insert tools only
/// format what they create, so before this a stock control could not be made blue, nor a header row bold,
/// without hand-editing the OOXML. Like <see cref="LayoutEditor"/>, it works on an already-open document and
/// leaves opening, validating and saving to the caller.
/// </summary>
/// <remarks>
/// <para>
/// WHAT EACH TARGET FORMATS. The run formatting goes through <see cref="RunFormatter"/>, the same code the
/// insert tools use, onto:
/// <list type="bullet">
/// <item><b>control</b> (by <c>w:id</c>, in any part): every text run inside it, the <c>w:sdtPr/w:rPr</c> of it
/// and of every field/label control inside it (what Business Central formats the bound value with, and where
/// stock controls carry their formatting), and the paragraph marks of any paragraphs inside a block-, row- or
/// cell-level control. A picture control has no text and is refused.</item>
/// <item><b>cell</b>, or every cell of a <b>row</b>: the same, for everything in the cell except a nested
/// table (which is a table of its own to address), plus the cell-level control a BC header cell is wrapped
/// in. The paragraph marks are formatted too, as Word does when a whole cell is selected: an empty cell then
/// already carries the formatting, and text inserted into it later takes it (GitHub issue #122).</item>
/// <item><b>text</b>: exactly the one occurrence of a substring, split out of its run when it is only part
/// of one. The split pieces keep the run's own properties, so nothing around the match changes. Text inside a
/// data-bound control is refused: it is the value Business Central replaces at render time, and the control
/// target formats it properly.</item>
/// </list>
/// Alignment and the paragraph style (GitHub issue #146), which belong to the paragraph, go on the paragraphs of the
/// target (for an inline control or a text match, the one paragraph holding it). Clearing direct formatting
/// (<see cref="RunFormatter.ClearDirectFormatting"/>) reaches the same runs, control properties and paragraph marks
/// as the format, and happens first, so one call can clear a line, apply <c>Heading1</c> and keep it bold.
/// </para>
/// <para>
/// Picture runs are never formatted: their <c>w:rPr</c> is not text formatting (see <see cref="InheritedRunFormat"/>).
/// </para>
/// </remarks>
public static class TextFormatEditor
{
    /// <summary>
    /// Applies <paramref name="format"/>, <paramref name="alignment"/> and <paramref name="paragraphStyle"/> to
    /// <paramref name="target"/>, after removing its direct run formatting when <paramref name="clearDirectFormatting"/>.
    /// </summary>
    /// <exception cref="ArgumentException">
    /// The target is missing an argument its kind needs, or carries one it does not use (ParamName
    /// <c>target</c>); nothing to change is given, or a value is malformed (<c>format</c>, <c>alignment</c>,
    /// <c>paragraphStyle</c> - which must name a paragraph style the layout defines); the search text occurs more
    /// than once or inside a data-bound control (<c>searchText</c>); or the control is a picture (<c>target</c>).
    /// </exception>
    /// <exception cref="NotFoundException">The control, table coordinate or text does not exist.</exception>
    /// <exception cref="LayoutDefectException">More than one control carries the id.</exception>
    public static EditResult SetTextFormat(
        WordprocessingDocument doc,
        TextFormatTarget target,
        TextFormat? format,
        string? alignment,
        string? paragraphStyle = null,
        bool clearDirectFormatting = false)
    {
        ArgumentNullException.ThrowIfNull(doc);
        ArgumentNullException.ThrowIfNull(target);

        CheckTargetShape(target);
        var prepared = RunFormatter.Prepare(format, doc);
        JustificationValues? justification = alignment is null
            ? null
            : CellTextEditor.ParseAlignment(alignment, nameof(alignment));
        var style = paragraphStyle is null
            ? null
            : CellTextEditor.CheckParagraphStyle(doc, paragraphStyle, nameof(paragraphStyle));
        if (prepared is null && justification is null && style is null && !clearDirectFormatting)
        {
            throw new ArgumentException(
                "set_text_format has nothing to change: pass format with at least one member, alignment, "
                + "paragraphStyle, clearDirectFormatting true, or a combination.",
                nameof(format));
        }

        var change = new Change(prepared, justification, style, clearDirectFormatting, RunFormatter.FontNote(prepared, doc));
        return target.Kind switch
        {
            TextFormatTargetKind.Control => FormatControl(doc, target, change),
            TextFormatTargetKind.Cell or TextFormatTargetKind.Row => FormatTableCells(doc, target, change),
            _ => FormatText(doc, target, target.SearchText!, change),
        };
    }

    // ---- targets ----

    private static EditResult FormatControl(WordprocessingDocument doc, TextFormatTarget target, Change change)
    {
        var controlId = target.ControlId!.Value;
        var main = doc.MainDocumentPart
            ?? throw new InvalidDataException("Layout has no main document part.");
        var carriers = ControlIds.FindAll(main, controlId);
        var (control, part) = carriers.Count switch
        {
            0 => throw new NotFoundException(
                $"No control with id {Id(controlId)} was found in the document (main body, headers, or footers).",
                NotFoundTarget.Control),
            1 => carriers[0],
            _ => throw ControlIds.DuplicateDefect(controlId, carriers),
        };

        if (SdtInspector.Classify(control.SdtProperties) == SdtInspector.Classification.Picture)
        {
            throw new ArgumentException(
                $"Control {Id(controlId)} is a picture control: it has no text to format.", nameof(target));
        }

        change.FormatContent(control, belongsToScope: _ => true);
        if (control is SdtRun && control.Ancestors<Paragraph>().FirstOrDefault() is { } host)
        {
            // An inline control has no paragraph of its own; alignment and style go on the line it sits in.
            change.FormatParagraph(host);
        }

        var alias = SdtInspector.ReadAlias(control);
        return new EditResult
        {
            Operation = "set_text_format",
            ControlId = controlId,
            Alias = alias,
            XPath = SdtInspector.ReadXPath(control),
            Kind = SdtInspector.ClassifyControlKind(control).ToString(),
            Part = part,
            Summary = $"Formatted control {(alias is null ? string.Empty : $"'{alias}' ")}(id {Id(controlId)})"
                + PartSuffix(part) + change.Describe() + change.StyleNote(),
        };
    }

    private static EditResult FormatTableCells(WordprocessingDocument doc, TextFormatTarget target, Change change)
    {
        var tableIndex = target.TableIndex!.Value;
        var row = target.Row!.Value;
        List<TableCell> cells;
        string part;
        if (target.Kind == TextFormatTargetKind.Cell)
        {
            var (cell, partName) = LocationResolver.ResolveCellElement(
                new Location
                {
                    Type = LocationKind.TableCell,
                    TableIndex = tableIndex,
                    Row = row,
                    Col = target.Col,
                    Part = target.Part,
                    PartName = target.PartName,
                },
                doc);
            cells = [cell];
            part = partName;
        }
        else
        {
            (cells, part) = RowCells(doc, target, tableIndex, row);
        }

        foreach (var cell in cells)
        {
            change.FormatContent(cell, belongsToScope: e => ReferenceEquals(e.Ancestors<TableCell>().FirstOrDefault(), cell));
            foreach (var wrapper in CellLevelWrappers(cell))
            {
                change.FormatControlProperties(wrapper);
            }
        }

        var where = target.Kind == TextFormatTargetKind.Cell
            ? $"table {tableIndex} row {row} col {target.Col}"
            : $"all {cells.Count} cell(s) of table {tableIndex} row {row}";
        return new EditResult
        {
            Operation = "set_text_format",
            ControlId = 0,
            TableIndex = tableIndex,
            Part = part,
            Summary = $"Formatted {where}{PartSuffix(part)}" + change.Describe()
                + (change.Runs == 0 && change.Format is not null
                    ? " The target holds no text yet: its paragraph marks carry the formatting, so text inserted "
                      + "there later takes it."
                    : string.Empty)
                + change.StyleNote(),
        };
    }

    private static EditResult FormatText(WordprocessingDocument doc, TextFormatTarget target, string searchText, Change change)
    {
        var (root, part) = LocationResolver.ResolvePart(doc, target.Part, target.PartName);
        var description = LocationResolver.DescribePart(part);

        var holders = root.Descendants<Text>()
            .Where(t => t.Text is not null && t.Text.Contains(searchText, StringComparison.Ordinal))
            .ToList();
        var occurrences = holders.Sum(t => Occurrences(t.Text, searchText));
        if (occurrences == 0)
        {
            throw new NotFoundException(
                $"No text containing '{searchText}' was found in {description}.", NotFoundTarget.SearchText);
        }

        if (occurrences > 1)
        {
            throw new ArgumentException(
                $"'{searchText}' occurs {occurrences} times in {description}, so it does not identify one piece of "
                + "text; nothing was formatted. Pass a longer substring that occurs once, or target the cell or "
                + "control that holds the text.",
                nameof(searchText));
        }

        var holder = holders[0];
        if (holder.Ancestors<SdtElement>().FirstOrDefault(IsBoundContentControl) is { } bound)
        {
            var boundId = SdtInspector.ReadControlId(bound);
            var alias = SdtInspector.ReadAlias(bound);
            throw new ArgumentException(
                $"'{searchText}' is inside data-bound control {(alias is null ? string.Empty : $"'{alias}' ")}"
                + $"(id {boundId?.ToString(CultureInfo.InvariantCulture) ?? "?"}): its text is the value Business "
                + "Central replaces at render time. Format the control instead: target 'control', controlId "
                + $"{boundId?.ToString(CultureInfo.InvariantCulture) ?? "(see get_layout_info)"}.",
                nameof(searchText));
        }

        var run = holder.Parent as Run
            ?? throw new InvalidOperationException($"Text containing '{searchText}' is not inside a run.");
        var paragraph = run.Ancestors<Paragraph>().FirstOrDefault();
        var matched = SplitAround(run, holder, searchText, out var split);

        change.FormatRun(matched);
        if (paragraph is not null)
        {
            change.FormatParagraph(paragraph);
        }

        return new EditResult
        {
            Operation = "set_text_format",
            ControlId = 0,
            Part = part,
            Summary = $"Formatted the text '{searchText}' in {description}"
                + (split ? " (split out of the run it was part of; the rest keeps its formatting)" : string.Empty)
                + change.Describe()
                + change.StyleNote(),
        };
    }

    // ---- helpers ----

    /// <summary>
    /// Refuses a target missing an argument its kind needs, or carrying one it does not use: a stray argument
    /// means the caller misread the tool, and acting on the rest would be a guess (ADR-0003).
    /// </summary>
    private static void CheckTargetShape(TextFormatTarget target)
    {
        var kind = target.Kind.ToString().ToLowerInvariant();
        var missing = new List<string>();
        var stray = new List<string>();

        void Need(bool present, string name)
        {
            if (!present)
            {
                missing.Add(name);
            }
        }

        void Refuse(bool present, string name)
        {
            if (present)
            {
                stray.Add(name);
            }
        }

        switch (target.Kind)
        {
            case TextFormatTargetKind.Control:
                Need(target.ControlId is not null, "controlId");
                Refuse(target.TableIndex is not null, "tableIndex");
                Refuse(target.Row is not null, "row");
                Refuse(target.Col is not null, "col");
                Refuse(target.SearchText is not null, "searchText");
                Refuse(target.Part != LayoutPart.Body, "layoutPart");
                Refuse(target.PartName is not null, "partName");
                break;

            case TextFormatTargetKind.Cell:
            case TextFormatTargetKind.Row:
                Need(target.TableIndex is >= 0, "tableIndex");
                Need(target.Row is >= 0, "row");
                if (target.Kind == TextFormatTargetKind.Cell)
                {
                    Need(target.Col is >= 0, "col");
                }
                else
                {
                    Refuse(target.Col is not null, "col");
                }

                Refuse(target.ControlId is not null, "controlId");
                Refuse(target.SearchText is not null, "searchText");
                break;

            default:
                Need(!string.IsNullOrEmpty(target.SearchText), "searchText");
                Refuse(target.ControlId is not null, "controlId");
                Refuse(target.TableIndex is not null, "tableIndex");
                Refuse(target.Row is not null, "row");
                Refuse(target.Col is not null, "col");
                break;
        }

        if (missing.Count == 0 && stray.Count == 0)
        {
            return;
        }

        var problems = new List<string>();
        if (missing.Count > 0)
        {
            problems.Add($"needs {string.Join(", ", missing)}"
                + (target.Kind is TextFormatTargetKind.Cell or TextFormatTargetKind.Row ? " (0-based, non-negative)" : string.Empty));
        }

        if (stray.Count > 0)
        {
            problems.Add($"does not use {string.Join(", ", stray)}"
                + (target.Kind == TextFormatTargetKind.Control && stray.Any(s => s.StartsWith("layoutPart", StringComparison.Ordinal) || s == "partName")
                    ? " (a control is found by its id in any part)"
                    : target.Kind == TextFormatTargetKind.Row && stray.Contains("col")
                        ? " (a row target formats every cell; for one cell use target 'cell')"
                        : string.Empty));
        }

        throw new ArgumentException(
            $"target '{kind}' {string.Join(" and ", problems)}; nothing was formatted.", nameof(target));
    }

    /// <summary>
    /// Every physical cell of one row, the same walk <c>set_cell_format</c> makes (a cell-level control wrapping
    /// nothing has no cell to format and is skipped).
    /// </summary>
    private static (List<TableCell> Cells, string Part) RowCells(
        WordprocessingDocument doc, TextFormatTarget target, int tableIndex, int row)
    {
        var (root, part) = LocationResolver.ResolvePart(doc, target.Part, target.PartName);
        var table = TableGridNavigator.TableAt(root, tableIndex, LocationResolver.DescribePart(part));
        var rows = TableGridNavigator.Rows(table);
        if (row >= rows.Count)
        {
            throw new NotFoundException(
                $"Row index {Id(row)} is out of range for table {Id(tableIndex)}; it has {Id(rows.Count)} row(s).",
                NotFoundTarget.TableCoordinate);
        }

        var innerRow = rows[row].InnerRow
            ?? throw new NotFoundException(
                $"Table {Id(tableIndex)}, row {Id(row)} resolves to a row-level control with no inner table row.",
                NotFoundTarget.TableCoordinate);

        var cells = TableGridNavigator.Cells(innerRow)
            .Select(slot => slot.InnerCell)
            .OfType<TableCell>()
            .ToList();
        return (cells, part);
    }

    /// <summary>
    /// The cell-level controls wrapping <paramref name="cell"/> directly — the chain a BC header address field
    /// is (<c>w:sdt</c> › <c>w:sdtContent</c> › <c>w:tc</c>), never a control further out around the table.
    /// </summary>
    private static IEnumerable<SdtCell> CellLevelWrappers(TableCell cell)
    {
        var parent = cell.Parent;
        while (parent is SdtContentCell content && content.Parent is SdtCell wrapper)
        {
            yield return wrapper;
            parent = wrapper.Parent;
        }
    }

    /// <summary>
    /// Splits <paramref name="run"/> so <paramref name="searchText"/> (inside <paramref name="holder"/>) is a run
    /// of its own, and returns that run. Each piece is a copy of the run (its attributes and <c>w:rPr</c>) with
    /// its share of the content, so the text around the match renders exactly as before. A run that holds
    /// nothing but the match is returned as it is.
    /// </summary>
    private static Run SplitAround(Run run, Text holder, string searchText, out bool split)
    {
        var value = holder.Text;
        var start = value.IndexOf(searchText, StringComparison.Ordinal);
        var before = value[..start];
        var after = value[(start + searchText.Length)..];

        var content = run.ChildElements.Where(c => c is not RunProperties).ToList();
        var at = content.IndexOf(holder);
        var leading = content.Take(at).ToList();
        var trailing = content.Skip(at + 1).ToList();

        split = before.Length > 0 || after.Length > 0 || leading.Count > 0 || trailing.Count > 0;
        if (!split)
        {
            return run;
        }

        Run Piece(IEnumerable<OpenXmlElement> children)
        {
            var piece = (Run)run.CloneNode(false);
            if (run.RunProperties?.CloneNode(true) is RunProperties rPr)
            {
                piece.AppendChild(rPr);
            }

            foreach (var child in children)
            {
                piece.AppendChild(child.Parent is null ? child : child.CloneNode(true));
            }

            return piece;
        }

        static Text Literal(string text) => new(text) { Space = SpaceProcessingModeValues.Preserve };

        var pieces = new List<Run>();
        if (leading.Count > 0 || before.Length > 0)
        {
            pieces.Add(Piece(before.Length > 0 ? [.. leading, Literal(before)] : leading));
        }

        var match = Piece([Literal(searchText)]);
        pieces.Add(match);

        if (trailing.Count > 0 || after.Length > 0)
        {
            pieces.Add(Piece(after.Length > 0 ? [Literal(after), .. trailing] : trailing));
        }

        foreach (var piece in pieces)
        {
            run.InsertBeforeSelf(piece);
        }

        run.Remove();
        return match;
    }

    private static int Occurrences(string text, string searchText)
    {
        var count = 0;
        for (var at = text.IndexOf(searchText, StringComparison.Ordinal);
             at >= 0;
             at = text.IndexOf(searchText, at + searchText.Length, StringComparison.Ordinal))
        {
            count++;
        }

        return count;
    }

    /// <summary>A field/label or picture control: its content is its bound value, not layout.</summary>
    private static bool IsBoundContentControl(SdtElement sdt) =>
        SdtInspector.Classify(sdt.SdtProperties)
            is SdtInspector.Classification.Bound or SdtInspector.Classification.Picture;

    private static bool IsTextControl(SdtElement sdt) =>
        SdtInspector.Classify(sdt.SdtProperties) == SdtInspector.Classification.Bound;

    private static string PartSuffix(string part) => part == "document.xml" ? string.Empty : $" in {part}";

    private static string Id(int value) => value.ToString(CultureInfo.InvariantCulture);

    /// <summary>One call's change, and a tally of what it touched for the summary.</summary>
    private sealed class Change(
        RunFormatter.Prepared? format,
        JustificationValues? alignment,
        string? paragraphStyle,
        bool clearDirect,
        string fontNote)
    {
        public RunFormatter.Prepared? Format { get; } = format;

        public int Runs { get; private set; }

        private int controls;
        private int marks;
        private int paragraphs;
        private int cleared;
        private int themeColorsDropped;
        private int styleOverridingRuns;

        /// <summary>True when the call changes run formatting: it clears it, sets it, or both.</summary>
        private bool TouchesRuns => Format is not null || clearDirect;

        /// <summary>
        /// Formats everything under <paramref name="container"/> (and the container itself, when it is a
        /// control) that <paramref name="belongsToScope"/> accepts: text runs, field/label controls' own
        /// formatting, paragraph marks, and the paragraphs' alignment and style.
        /// </summary>
        public void FormatContent(OpenXmlElement container, Func<OpenXmlElement, bool> belongsToScope)
        {
            foreach (var run in container.Descendants<Run>().Where(r => belongsToScope(r)).ToList())
            {
                FormatRun(run);
            }

            var sdts = container.Descendants<SdtElement>().Where(s => belongsToScope(s)).ToList();
            if (container is SdtElement self)
            {
                sdts.Insert(0, self);
            }

            foreach (var sdt in sdts)
            {
                FormatControlProperties(sdt);
            }

            foreach (var paragraph in container.Descendants<Paragraph>().Where(p => belongsToScope(p)).ToList())
            {
                FormatMark(paragraph);
                FormatParagraph(paragraph);
            }
        }

        public void FormatRun(Run run)
        {
            if (!TouchesRuns || InheritedRunFormat.IsGraphic(run))
            {
                return;
            }

            Clear(run.RunProperties);
            if (Format is not null)
            {
                run.RunProperties ??= new RunProperties();
                Count(RunFormatter.Apply(run.RunProperties, Format));
            }

            Runs++;
        }

        /// <summary>The <c>w:sdtPr/w:rPr</c> of a field/label control, first in its <c>w:sdtPr</c> as stock controls carry it.</summary>
        public void FormatControlProperties(SdtElement sdt)
        {
            if (!TouchesRuns || !IsTextControl(sdt) || sdt.SdtProperties is not { } sdtPr)
            {
                return;
            }

            var rPr = sdtPr.GetFirstChild<RunProperties>();
            Clear(rPr);
            if (Format is not null)
            {
                rPr = sdtPr.GetFirstChild<RunProperties>();
                if (rPr is null)
                {
                    rPr = new RunProperties();
                    sdtPr.PrependChild(rPr);
                }

                Count(RunFormatter.Apply(rPr, Format));
            }

            controls++;
        }

        /// <summary>
        /// Sets the paragraph's alignment and style, when the call gives them. A style is what the paragraph's text
        /// shows only where no direct run formatting overrides it, so without <c>clearDirectFormatting</c> the runs
        /// that still carry some are counted for <see cref="StyleNote"/>; with it, the mark of a paragraph given a
        /// style is cleared too, even the line an inline control or a text match sits in, since the mark's size still
        /// sets the height of the paragraph's last line.
        /// </summary>
        public void FormatParagraph(Paragraph paragraph)
        {
            if (alignment is null && paragraphStyle is null)
            {
                return;
            }

            if (alignment is { } justification)
            {
                CellTextEditor.Align(paragraph, justification);
            }

            if (paragraphStyle is not null)
            {
                CellTextEditor.SetParagraphStyle(paragraph, paragraphStyle);
                Clear(paragraph.ParagraphProperties!.ParagraphMarkRunProperties);
                if (!clearDirect)
                {
                    styleOverridingRuns += paragraph.Descendants<Run>()
                        .Count(r => !InheritedRunFormat.IsGraphic(r) && RunFormatter.HasDirectFormatting(r.RunProperties));
                }
            }

            paragraphs++;
        }

        /// <summary>
        /// The sentence that warns when the paragraph style is hidden by direct run formatting the call left in place;
        /// empty otherwise.
        /// </summary>
        public string StyleNote() =>
            styleOverridingRuns == 0
                ? string.Empty
                : $" {Id(styleOverridingRuns)} run(s) in those paragraphs carry direct formatting (a font, size, colour "
                  + $"and the like), which wins over the paragraph style '{paragraphStyle}'. Pass clearDirectFormatting "
                  + "true to let the style decide.";

        private void FormatMark(Paragraph paragraph)
        {
            if (!TouchesRuns)
            {
                return;
            }

            Clear(paragraph.ParagraphProperties?.ParagraphMarkRunProperties);
            if (Format is not null)
            {
                paragraph.ParagraphProperties ??= new ParagraphProperties();
                paragraph.ParagraphProperties.ParagraphMarkRunProperties ??= new ParagraphMarkRunProperties();
                Count(RunFormatter.Apply(paragraph.ParagraphProperties.ParagraphMarkRunProperties, Format));
            }

            marks++;
        }

        /// <summary>Clears <paramref name="properties"/>' direct formatting when asked, dropping it once it is empty.</summary>
        private void Clear(OpenXmlCompositeElement? properties)
        {
            if (!clearDirect || properties is null)
            {
                return;
            }

            if (RunFormatter.ClearDirectFormatting(properties))
            {
                cleared++;
            }

            if (!properties.HasChildren)
            {
                properties.Remove();
            }
        }

        private void Count(bool droppedThemeColor)
        {
            if (droppedThemeColor)
            {
                themeColorsDropped++;
            }
        }

        /// <summary>
        /// ": bold, 8 pt (3 run(s), 1 control, 2 paragraph mark(s)); aligned right, paragraph style 'Heading1'
        /// (2 paragraph(s))."
        /// </summary>
        public string Describe()
        {
            var parts = new List<string>();
            if (TouchesRuns)
            {
                var touched = new List<string> { $"{Id(Runs)} run(s)" };
                if (controls > 0)
                {
                    touched.Add($"{Id(controls)} control(s)' own formatting");
                }

                if (marks > 0)
                {
                    touched.Add($"{Id(marks)} paragraph mark(s)");
                }

                var what = new List<string>(2);
                if (clearDirect)
                {
                    what.Add($"cleared the direct formatting ({Id(cleared)} had some)");
                }

                if (Format is not null)
                {
                    what.Add(RunFormatter.Describe(Format));
                }

                parts.Add($"{string.Join(", then ", what)} ({string.Join(", ", touched)})");
            }

            var paragraphChanges = new List<string>(2);
            if (alignment is { } justification)
            {
                paragraphChanges.Add($"aligned {justification.ToString().ToLowerInvariant()}");
            }

            if (paragraphStyle is not null)
            {
                paragraphChanges.Add($"paragraph style '{paragraphStyle}'");
            }

            if (paragraphChanges.Count > 0)
            {
                parts.Add($"{string.Join(", ", paragraphChanges)} ({Id(paragraphs)} paragraph(s))");
            }

            var themeNote = themeColorsDropped == 0
                ? string.Empty
                : $" {Id(themeColorsDropped)} of them had a theme colour, which the literal colour replaces: that text "
                  + "no longer follows the layout's theme. Pass a theme colour such as 'accent1' to keep it themable.";
            return ": " + string.Join("; ", parts) + "." + themeNote + fontNote;
        }
    }
}
