using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Office2013Word = DocumentFormat.OpenXml.Office2013.Word;

namespace BcWordLayout.Domain;

/// <summary>
/// Deterministic OOXML mutations for a BC Word layout: inserting field/label content controls and
/// removing controls. Every operation here works against an ALREADY-OPEN <see cref="WordprocessingDocument"/>
/// and does no file I/O of its own — opening, pre-save validation, and saving are the caller's job (see
/// <c>BcWordLayout.McpHost.Tools.EditTools</c>'s <c>insert_field</c>/<c>insert_label</c>/<c>remove_control</c>
/// and <c>BcWordLayout.McpHost.Tools.TableTools</c>'s <c>insert_repeater_table</c>, which wrap these calls
/// with an open/validate/save-or-reject flow so a bad edit can never reach disk). Building the control itself is
/// <see cref="SdtFactory"/>'s job; finding WHERE to put it is <see cref="LocationResolver"/>'s job — this
/// type is the thin layer that ties both together against a real document and reports what changed.
/// </summary>
public static class LayoutEditor
{
    /// <summary>
    /// Inserts a plain-text FIELD content control bound to <paramref name="datasetPath"/> at
    /// <paramref name="location"/>, using a freshly generated <c>w:id</c> guaranteed unique within
    /// <paramref name="doc"/> (see <see cref="GenerateUniqueId"/>). See <see cref="SdtFactory.BuildField"/>
    /// for dataset-path validation rules and <see cref="LocationResolver.Resolve"/> for location resolution
    /// rules; both propagate unchanged (<see cref="ArgumentException"/> / <see cref="NotFoundException"/>).
    /// The optional <paramref name="format"/> carries RUN-level knobs only (<see cref="CellTextFormat.Bold"/>,
    /// <see cref="CellTextFormat.FontSizePoints"/>) applied to the control's own runs and <c>w:sdtPr/w:rPr</c>
    /// — needed because a control inserted into a freshly authored plain-table cell has no styling to
    /// inherit; <see cref="CellTextFormat.Alignment"/> is a paragraph/cell concern and is rejected here
    /// (seed it via <c>insert_table</c>'s columnAlignments or <c>set_cell_text</c>). A header/footer-targeted
    /// insert at the end of a layout that has no such part yet scaffolds an empty one first rather than
    /// failing — see <see cref="EnsureTargetPartExists"/>.
    /// </summary>
    public static EditResult InsertField(
        WordprocessingDocument doc, string datasetPath, Location location, CellTextFormat? format = null) =>
        Insert(doc, datasetPath, location, isLabel: false, format);

    /// <summary>Inserts a LABEL content control bound to <paramref name="datasetPath"/>. See <see cref="InsertField"/>.</summary>
    public static EditResult InsertLabel(
        WordprocessingDocument doc, string datasetPath, Location location, CellTextFormat? format = null) =>
        Insert(doc, datasetPath, location, isLabel: true, format);

    /// <summary>
    /// Inserts a plain STATIC text run (<c>w:r</c>/<c>w:t</c>) at <paramref name="location"/> — no content
    /// control, no binding: the literal glue every real corpus header and footer uses between its controls (a
    /// separator space, a colon, a <c>" / "</c> between a date and a page number).
    /// </summary>
    /// <remarks>
    /// <para>
    /// Fills the gap two <c>afterControl</c> inserts leave: chaining inline controls concatenates their
    /// rendered text with nothing in between ("Document NoDOCU-0150"), and no other tool emits a bare run, so
    /// the only previous workaround was stacking each fragment in its own paragraph. A gap found by the
    /// 2026-07-31 from-scratch authoring exercise.
    /// </para>
    /// <para>
    /// UNLIKE EVERY OTHER INSERT, THIS CREATES NO CONTROL. There is no <c>w:sdt</c> and therefore no
    /// <c>w:id</c>, so the returned <see cref="EditResult.ControlId"/> is 0 and the run cannot afterwards be
    /// addressed by <c>remove_control</c> or used as an <c>afterControl</c> anchor. It can still be found and
    /// edited by text: <c>set_cell_text</c> for a run in a table cell, or an <c>atText</c> location. That
    /// asymmetry is inherent to static text rather than a shortcut — wrapping the run in an unbound content
    /// control purely to give it an id would put a control in the layout that BC has no reason to see, and
    /// would show up in every <c>get_layout_info</c> inventory as an unbound control needing explanation.
    /// </para>
    /// <para>
    /// Whitespace is preserved verbatim (<c>xml:space="preserve"</c>), because the single most likely use is a
    /// run that is nothing BUT whitespace — a separator space between two controls, which Word would otherwise
    /// collapse away and silently undo the whole point of the call.
    /// </para>
    /// </remarks>
    /// <exception cref="ArgumentException">
    /// <paramref name="text"/> is null/empty (nothing to insert — an empty run would be invisible clutter);
    /// or <paramref name="format"/> sets an alignment (a run is inline — see
    /// <see cref="TryPrepareRunFormat"/>) or an out-of-range font size. No length cap is imposed, matching
    /// <c>set_cell_text</c>, which has none either.
    /// </exception>
    /// <exception cref="NotFoundException">Propagated from <see cref="LocationResolver.Resolve"/>.</exception>
    public static EditResult InsertText(
        WordprocessingDocument doc, string text, Location location, CellTextFormat? format = null)
    {
        ArgumentNullException.ThrowIfNull(doc);
        ArgumentNullException.ThrowIfNull(location);

        if (string.IsNullOrEmpty(text))
        {
            throw new ArgumentException(
                "insert_text needs a non-empty text; there is nothing to insert. A run of spaces IS valid "
                + "(that is the separator case this tool exists for) - pass \" \" rather than \"\".",
                nameof(text));
        }

        var run = new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        ApplyRunFormat(run, format, "insert_text");

        var scaffolded = EnsureTargetPartExists(doc, location);
        var anchor = LocationResolver.Resolve(location, doc);
        anchor.InsertInline(run);

        return new EditResult
        {
            Operation = "insert_text",
            ControlId = 0, // no content control exists to carry an id - see this method's remarks
            Kind = "StaticText",
            Part = anchor.PartName,
            Summary = $"Inserted the static text \"{text}\" at {DescribeLocation(location, anchor.PartName)}. "
                + "It is a plain run, not a content control, so it has no controlId and cannot be targeted by "
                + "remove_control or used as an afterControl anchor."
                + (scaffolded
                    ? $" The layout had no {location.Part.ToString().ToLowerInvariant()} part, so an empty "
                      + $"{anchor.PartName} was created and wired into the page setup first."
                    : string.Empty),
        };
    }

    /// <summary>
    /// Inserts a PICTURE content control bound to <paramref name="datasetPath"/> at
    /// <paramref name="location"/> — the placeholder BC fills with a real image at render time, e.g. the
    /// company logo (<c>/Header/CompanyPicture</c>) a from-scratch layout could not carry at all before this
    /// (merge/preview already FILLED existing picture controls; only authoring one was missing). See
    /// <see cref="SdtFactory.BuildPicture"/> for the exact corpus-mirrored shape.
    /// </summary>
    /// <remarks>
    /// Two things a picture needs that a text control does not, both handled here rather than in
    /// <see cref="SdtFactory"/> (which never touches a package): a real <see cref="ImagePart"/> in the
    /// HOSTING part for the blip to reference — a dangling <c>r:embed</c> is a corrupt document to Word — and
    /// a document-wide-unique <c>wp:docPr/@id</c> (see <see cref="NextDrawingId"/>). The image part is added
    /// to the same part the control lands in (body/header/footer), since a relationship id only resolves
    /// within its own part, and holds <see cref="PlaceholderImage"/>'s gray PNG so the frame renders as a
    /// visible placeholder rather than a broken-image box.
    /// </remarks>
    /// <exception cref="ArgumentException">
    /// Propagated from <see cref="SdtFactory.BuildPicture"/> (bad dataset path, non-positive size), or
    /// <paramref name="widthMm"/>/<paramref name="heightMm"/> outside 1–500.
    /// </exception>
    /// <exception cref="NotFoundException">Propagated from <see cref="LocationResolver.Resolve"/>.</exception>
    public static EditResult InsertPicture(
        WordprocessingDocument doc, string datasetPath, Location location, double widthMm = 30, double heightMm = 30)
    {
        ArgumentNullException.ThrowIfNull(doc);
        ArgumentNullException.ThrowIfNull(location);

        if (widthMm is < 1 or > 500)
        {
            throw new ArgumentException($"widthMm must be between 1 and 500 (got {widthMm}).", nameof(widthMm));
        }

        if (heightMm is < 1 or > 500)
        {
            throw new ArgumentException($"heightMm must be between 1 and 500 (got {heightMm}).", nameof(heightMm));
        }

        var schema = SchemaProvider.FromLayout(doc);
        var id = GenerateUniqueId(doc);
        var drawingId = NextDrawingId(doc);

        var scaffolded = EnsureTargetPartExists(doc, location);
        var anchor = LocationResolver.Resolve(location, doc);

        // The image part must live in the SAME part as the control — a relationship id is part-scoped.
        var hostPart = anchor.HostPart(doc);
        var imagePart = hostPart.AddNewPart<ImagePart>("image/png");
        using (var stream = new MemoryStream(PlaceholderImage.PngBytes))
        {
            imagePart.FeedData(stream);
        }

        var sdt = SdtFactory.BuildPicture(
            schema,
            datasetPath,
            hostPart.GetIdOfPart(imagePart),
            drawingId,
            (long)Math.Round(widthMm * SdtFactory.EmuPerMillimetre),
            (long)Math.Round(heightMm * SdtFactory.EmuPerMillimetre),
            id);

        anchor.InsertInline(sdt);

        return new EditResult
        {
            Operation = "insert_picture",
            ControlId = id,
            Alias = sdt.SdtProperties!.GetFirstChild<SdtAlias>()?.Val?.Value,
            XPath = sdt.SdtProperties!.GetFirstChild<DataBinding>()?.XPath?.Value,
            Kind = ControlKind.Picture.ToString(),
            Part = anchor.PartName,
            Summary = $"Inserted picture placeholder '{datasetPath}' (id {id}, {widthMm}x{heightMm} mm) at "
                + $"{DescribeLocation(location, anchor.PartName)}; it renders as a gray placeholder until BC "
                + "fills it."
                + (scaffolded
                    ? $" The layout had no {location.Part.ToString().ToLowerInvariant()} part, so an empty "
                      + $"{anchor.PartName} was created and wired into the page setup first."
                    : string.Empty),
        };
    }

    /// <summary>
    /// The lowest <c>wp:docPr/@id</c> not already used by any drawing anywhere in <paramref name="doc"/>
    /// (main body, headers, footers). That id must be document-wide unique — a duplicate trips
    /// <see cref="OpenXmlValidator"/>'s <c>Sem_UniqueAttributeValue</c> and Word's own repair prompt, the
    /// same rule that forced <c>MergeEngine</c> to regenerate these ids when it clones a repeater row that
    /// contains a picture.
    /// </summary>
    private static uint NextDrawingId(WordprocessingDocument doc)
    {
        var main = doc.MainDocumentPart
            ?? throw new InvalidDataException("Layout has no main document part.");

        uint highest = 0;
        foreach (var (root, _) in PartWalker.ContentParts(main))
        {
            foreach (var docPr in root.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>())
            {
                if (docPr.Id?.Value is { } value && value > highest)
                {
                    highest = value;
                }
            }
        }

        return highest + 1;
    }

    /// <summary>
    /// Builds a complete repeater TABLE for <paramref name="dataItemPath"/> (see
    /// <see cref="SdtFactory.BuildRepeaterTable"/> for validation rules and exact shape) and inserts it as a
    /// BLOCK at <paramref name="location"/>, generating every sdt <c>w:id</c> it needs (header labels, data
    /// fields, the repeatingSectionItem wrapper, the repeatingSection wrapper itself) via a closure that
    /// guarantees uniqueness both against <paramref name="doc"/>'s pre-existing ids (<see cref="CollectAllControlIds"/>,
    /// the same check <see cref="GenerateUniqueId"/> uses) AND against every id already issued earlier in
    /// THIS SAME call — unlike <see cref="GenerateUniqueId"/>, which is only ever called once per single-
    /// control insert, this table needs many ids in one go, and each one must avoid every other one already
    /// handed out for this table before that id is ever attached to the in-memory tree (so a later
    /// <see cref="CollectAllControlIds"/>-style rescan would not yet see it).
    /// </summary>
    /// <remarks>
    /// SCOPE (this is the single enforcement chokepoint): <paramref name="location"/>.Part must be
    /// <see cref="LayoutPart.Body"/>. Unlike <see cref="InsertField"/>/<see cref="InsertLabel"/> (which
    /// fully support <see cref="LayoutPart.Header"/>/<see cref="LayoutPart.Footer"/>), a repeater TABLE in a
    /// header/footer is explicitly out of scope — tracked as GitHub issue #10 ("Repeater tables in
    /// headers/footers"). This is checked first, before any
    /// schema/id work, so a rejected call never touches <paramref name="doc"/>.
    /// </remarks>
    /// <exception cref="ArgumentException">
    /// <paramref name="location"/>.Part is not <see cref="LayoutPart.Body"/> (see remarks); or propagated
    /// unchanged from <see cref="SdtFactory.BuildRepeaterTable"/> (bad data item path, bad column, bad
    /// column-width count) or <see cref="LocationResolver.Resolve"/> (structurally invalid
    /// <paramref name="location"/>).
    /// </exception>
    /// <exception cref="NotFoundException">
    /// Propagated unchanged from <see cref="LocationResolver.Resolve"/> (a well-formed
    /// <paramref name="location"/> that does not actually resolve against <paramref name="doc"/>).
    /// </exception>
    public static EditResult InsertRepeaterTable(
        WordprocessingDocument doc,
        string dataItemPath,
        IReadOnlyList<string> columns,
        Location location,
        RepeaterTableOptions options)
    {
        ArgumentNullException.ThrowIfNull(doc);
        ArgumentNullException.ThrowIfNull(columns);
        ArgumentNullException.ThrowIfNull(location);
        ArgumentNullException.ThrowIfNull(options);

        if (location.Part != LayoutPart.Body)
        {
            throw new ArgumentException(
                $"insert_repeater_table only supports {nameof(Location)}.{nameof(Location.Part)} = "
                + $"{nameof(LayoutPart)}.{nameof(LayoutPart.Body)} in v1: a repeater TABLE in a header/footer "
                + "is explicitly deferred (repeaters in headers/footers - tracked as GitHub "
                + "issue #10). insert_field/insert_label "
                + "still fully support layoutPart='header'/'footer' - only a repeater TABLE cannot target one "
                + "yet; omit layoutPart (or pass 'body') to insert this table into the main document instead.",
                nameof(location));
        }

        var schema = SchemaProvider.FromLayout(doc);
        var nextId = MakeIdGenerator(doc);

        var table = SdtFactory.BuildRepeaterTable(schema, dataItemPath, columns, options, nextId);

        var anchor = LocationResolver.Resolve(location, doc);
        anchor.InsertBlock(table);
        SeparateFromAdjacentTables(table);

        var repeaterRow = table.Descendants<SdtRow>().First(SdtInspector.IsRepeater);
        var alias = SdtInspector.ReadAlias(repeaterRow);
        var xpath = SdtInspector.ReadXPath(repeaterRow);

        // "Unexpected" is the operative word: SdtFactory.BuildRepeaterTable always assigns the repeater row
        // an id from the caller's own generator - a missing one here would mean SdtFactory itself is broken,
        // not that the caller passed something invalid. Left as a plain InvalidOperationException
        // (→ internal_error), not NotFoundException - retrying will not fix a tool bug.
        var controlId = SdtInspector.ReadControlId(repeaterRow)
            ?? throw new InvalidOperationException("The newly built repeater row has no w:id (unexpected).");

        var (tableIndex, dataRowIndex) = LocateNewTable(doc, location, table, repeaterRow);

        return new EditResult
        {
            Operation = "insert_repeater_table",
            ControlId = controlId,
            Alias = alias,
            XPath = xpath,
            Kind = ControlKind.Repeater.ToString(),
            ColumnCount = columns.Count,
            TableIndex = tableIndex,
            DataRowIndex = dataRowIndex,
            Part = anchor.PartName,
            Summary = $"Inserted repeater table for '{dataItemPath}' with {columns.Count} column(s) "
                + $"(id {controlId}) at {DescribeLocation(location, anchor.PartName)}; it is table "
                + $"{tableIndex}, and its repeating DATA row is row {dataRowIndex} — the coordinate to pass "
                + "as tableIndex/row when adding a NESTED repeater (or any other per-row edit) inside it.",
        };
    }

    /// <summary>
    /// The freshly inserted table's own (tableIndex, dataRowIndex) coordinate — the same per-part,
    /// document-order numbering <c>get_layout_info</c> and every <c>tableCell</c> location use. Reported so
    /// authoring a NESTED repeater is one call away instead of requiring a re-read of the
    /// layout plus reasoning about which of the new table's rows is the repeating template: nesting is
    /// exactly "insert a repeater table into a cell of an existing repeater's data row", and until now
    /// nothing in the result said where that row was.
    /// </summary>
    private static (int TableIndex, int DataRowIndex) LocateNewTable(
        WordprocessingDocument doc, Location location, Table table, SdtRow repeaterRow)
    {
        var (root, _) = LocationResolver.ResolvePart(doc, location.Part, location.PartName);
        var tables = TableGridNavigator.Tables(root);

        var tableIndex = -1;
        for (var i = 0; i < tables.Count; i++)
        {
            if (ReferenceEquals(tables[i], table))
            {
                tableIndex = i;
                break;
            }
        }

        if (tableIndex < 0)
        {
            throw new InvalidOperationException("The newly inserted repeater table was not found in its own part (unexpected).");
        }

        var rows = TableGridNavigator.Rows(table);
        var dataRowIndex = -1;
        for (var i = 0; i < rows.Count; i++)
        {
            if (ReferenceEquals(rows[i].RowChild, repeaterRow))
            {
                dataRowIndex = i;
                break;
            }
        }

        if (dataRowIndex < 0)
        {
            throw new InvalidOperationException("The newly inserted repeater row was not found in its own table (unexpected).");
        }

        return (tableIndex, dataRowIndex);
    }

    /// <summary>
    /// Inserts a nested DETAIL ROW repeater into an EXISTING repeater's item — the standard BC shape for
    /// per-line detail ("another line with the required details" under the line row: serial/lot nos,
    /// assembly components), corpus-verified as sibling rows INSIDE the outer
    /// <c>repeatingSectionItem</c>, aligned to the same table grid via <c>gridSpan</c>s. The new repeater
    /// row is appended at the END of the outer item's content — after the line row and any detail rows
    /// already there — so repeated calls stack detail lines in authoring order. Cell construction and
    /// validation are <see cref="SdtFactory.BuildDetailRepeaterRow"/>'s (label-shaped names become label
    /// controls, chained inline); the spans must cover the parent table's grid exactly, which the
    /// grid-consistency guard also enforces mechanically before anything reaches disk.
    /// </summary>
    /// <exception cref="ArgumentException">
    /// <paramref name="parentRepeaterId"/> is not a row-level repeater; <paramref name="dataItemPath"/> is
    /// not a DIRECT child data item of the parent repeater's bound item; the spans don't sum to the parent
    /// grid; or propagated from <see cref="SdtFactory.BuildDetailRepeaterRow"/>.
    /// </exception>
    /// <exception cref="NotFoundException">No control with <paramref name="parentRepeaterId"/> exists in the body.</exception>
    public static EditResult InsertRepeaterRow(
        WordprocessingDocument doc, int parentRepeaterId, string dataItemPath, IReadOnlyList<RepeaterRowCell> cells)
    {
        ArgumentNullException.ThrowIfNull(doc);
        ArgumentNullException.ThrowIfNull(cells);

        var context = ResolveRepeaterItem(doc, parentRepeaterId, cells, rowNoun: "a detail row");

        // The parent's bound data item, from its own alias ('#Nav: /Header/Line').
        var parentAlias = SdtInspector.ReadAlias(context.Parent);
        var parentPath = parentAlias is not null && parentAlias.StartsWith("#Nav:", StringComparison.Ordinal)
            ? parentAlias["#Nav:".Length..].Trim()
            : null;
        var normalizedChild = "/" + dataItemPath.Trim().Trim('/');
        if (parentPath is null
            || !normalizedChild.StartsWith(parentPath + "/", StringComparison.Ordinal)
            || normalizedChild[(parentPath.Length + 1)..].Contains('/'))
        {
            throw new ArgumentException(
                $"dataItem '{dataItemPath}' is not a DIRECT child data item of the parent repeater's own "
                + $"item ('{parentPath ?? "?"}'). A detail row repeats once per parent row, so it must bind "
                + "to a repeating data item nested directly under the parent's (see list_dataset_fields); "
                + "for a grandchild, add a detail row to the CHILD's repeater instead.",
                nameof(dataItemPath));
        }

        var widths = ComputeCellWidthsFromGrid(context.GridColumns, cells);

        var schema = SchemaProvider.FromLayout(doc);
        var nextId = MakeIdGenerator(doc);
        var detailRow = SdtFactory.BuildDetailRepeaterRow(schema, normalizedChild, cells, widths, nextId);
        context.ItemContent.AppendChild(detailRow);

        var controlId = SdtInspector.ReadControlId(detailRow)
            ?? throw new InvalidOperationException("The newly built detail repeater row has no w:id (unexpected).");

        return new EditResult
        {
            Operation = "insert_repeater_row",
            ControlId = controlId,
            Alias = SdtInspector.ReadAlias(detailRow),
            XPath = SdtInspector.ReadXPath(detailRow),
            Kind = ControlKind.Repeater.ToString(),
            ColumnCount = cells.Count,
            TableIndex = context.TableIndex,
            Part = "document.xml",
            Summary = $"Added a nested '{normalizedChild}' detail row (id {controlId}) inside repeater "
                + $"{parentRepeaterId}'s item - it expands once per parent row, below the line row.",
        };
    }

    /// <summary>
    /// Appends one STATIC row at the END of an existing repeater's item — the stock BC shape for per-group
    /// SUBTOTAL (and spacer) rows (GitHub issue #30): <c>SalespersonCommission.docx</c> (115) ends each
    /// salesperson's <c>repeatingSectionItem</c> with an empty spacer <c>w:tr</c> and a bold subtotal
    /// <c>w:tr</c>, AFTER the nested <c>Cust_Ledger_Entry</c> detail repeater — so the row renders once per
    /// PARENT row (per group), never per detail row. Cell construction is
    /// <see cref="SdtFactory.BuildStaticRow"/>'s (each <see cref="RepeaterRowCell.Columns"/> entry a FULL
    /// dataset path — the corpus subtotal binds a sibling non-repeating <c>Subtotals</c> item's columns);
    /// the spans must cover the parent table's grid exactly. Repeated calls stack rows in authoring order,
    /// exactly like <see cref="InsertRepeaterRow"/> — the corpus order is spacer first, subtotal second.
    /// The GROUP-HEADER half of the stock group shape needs no tool: the repeater's own line row (built by
    /// <see cref="InsertRepeaterTable"/>) is the per-group header row.
    /// </summary>
    /// <remarks>
    /// LIKE <see cref="InsertText"/>, THIS CREATES NO SINGLE CONTROL: the row is a plain <c>w:tr</c> whose
    /// bound cells are ordinary inline field/label controls, so the result's
    /// <see cref="EditResult.ControlId"/> is 0 (the bound cells' own ids are in the summary). At merge
    /// time the whole item — this row included — is cloned once per data row, which is exactly what makes
    /// the subtotal per-group; the merge engine's per-level XPath re-anchoring covers the row's bindings
    /// the same way it covers the corpus original's.
    /// </remarks>
    /// <exception cref="ArgumentException">
    /// <paramref name="parentRepeaterId"/> is not a row-level repeater in a table, or its item shape is
    /// unrecognized; <paramref name="cells"/> is empty or its spans don't sum to the parent grid;
    /// <paramref name="format"/> sets an alignment (per-cell alignments belong on the cell specs); or
    /// propagated from <see cref="SdtFactory.BuildStaticRow"/> (a path that does not resolve).
    /// </exception>
    /// <exception cref="NotFoundException">No control with <paramref name="parentRepeaterId"/> exists in the body.</exception>
    public static EditResult InsertSubtotalRow(
        WordprocessingDocument doc, int parentRepeaterId, IReadOnlyList<RepeaterRowCell> cells, CellTextFormat? format = null)
    {
        ArgumentNullException.ThrowIfNull(doc);
        ArgumentNullException.ThrowIfNull(cells);

        if (format?.Alignment is not null)
        {
            // Same rule (and reason) as insert_table_row: ApplyControlRunFormat below only runs for BOUND
            // cells, so an all-spacer row would silently swallow the knob instead of refusing it.
            throw new ArgumentException(
                "alignment is not supported as a row-level format on insert_subtotal_row; pass per-cell "
                + "alignments alongside the cell specs instead (one 'left'/'center'/'right' or '-' per cell).",
                nameof(format));
        }

        var context = ResolveRepeaterItem(doc, parentRepeaterId, cells, rowNoun: "a subtotal row");
        var widths = ComputeCellWidthsFromGrid(context.GridColumns, cells);

        var schema = SchemaProvider.FromLayout(doc);
        var nextId = MakeIdGenerator(doc);
        var row = SdtFactory.BuildStaticRow(schema, cells, widths, nextId);

        var boundControlIds = new List<int>();
        foreach (var sdt in row.Descendants<SdtRun>())
        {
            ApplyControlRunFormat(sdt, format, "insert_subtotal_row");
            if (SdtInspector.ReadControlId(sdt) is { } id)
            {
                boundControlIds.Add(id);
            }
        }

        context.ItemContent.AppendChild(row);

        var boundSummary = boundControlIds.Count == 0
            ? "no bound cells (a spacer row)"
            : $"{boundControlIds.Count} bound control(s) (id(s) {string.Join(", ", boundControlIds)})";

        return new EditResult
        {
            Operation = "insert_subtotal_row",
            ControlId = 0, // a plain w:tr; its bound cells carry their own ids (see boundSummary)
            Kind = "StaticRow",
            ColumnCount = cells.Count,
            TableIndex = context.TableIndex,
            Part = "document.xml",
            Summary = $"Appended a static row with {cells.Count} cell(s) and {boundSummary} at the end of "
                + $"repeater {parentRepeaterId}'s item - it renders once per group (per parent row), after "
                + "the line row and any nested detail rows. Repeated calls stack further rows below it "
                + "(the stock shape is a spacer row, then the bold subtotal row).",
        };
    }

    /// <summary>The resolved insertion context both repeater-item row inserts share — see <see cref="ResolveRepeaterItem"/>.</summary>
    private readonly record struct RepeaterItemContext(
        SdtRow Parent, SdtContentRow ItemContent, Table Table, IReadOnlyList<GridColumn> GridColumns, int? TableIndex);

    /// <summary>
    /// The shared plumbing of <see cref="InsertRepeaterRow"/> and <see cref="InsertSubtotalRow"/>: find the
    /// row-level repeater with <paramref name="parentRepeaterId"/> in the body, require it to actually BE a
    /// repeater sitting inside a table, require <paramref name="cells"/>' spans to cover that table's grid
    /// exactly, and descend to its <c>repeatingSectionItem</c>'s content container — the element both
    /// operations append their row to. <paramref name="rowNoun"/> ("a detail row"/"a subtotal row") only
    /// feeds the error messages.
    /// </summary>
    private static RepeaterItemContext ResolveRepeaterItem(
        WordprocessingDocument doc, int parentRepeaterId, IReadOnlyList<RepeaterRowCell> cells, string rowNoun)
    {
        var body = doc.MainDocumentPart?.Document?.Body
            ?? throw new InvalidDataException("Layout has no document body.");

        var parent = body.Descendants<SdtRow>()
                .FirstOrDefault(s => SdtInspector.ReadControlId(s) == parentRepeaterId)
            ?? throw new NotFoundException(
                $"No row-level repeater control with id {parentRepeaterId} exists in the document body. "
                + "Pass the controlId insert_repeater_table returned (or the repeater's id from "
                + "get_layout_info's control inventory).",
                NotFoundTarget.Control);

        if (!SdtInspector.IsRepeater(parent))
        {
            throw new ArgumentException(
                $"Control {parentRepeaterId} is not a repeater (repeating section); {rowNoun} can only "
                + "be added inside a repeater's item.",
                nameof(parentRepeaterId));
        }

        // The parent table's grid is the row's coordinate system: spans must cover it exactly.
        var table = parent.Ancestors<Table>().FirstOrDefault()
            ?? throw new ArgumentException(
                $"Repeater {parentRepeaterId} does not sit inside a table; {rowNoun} only applies to "
                + "repeater TABLES.",
                nameof(parentRepeaterId));
        var gridColumns = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().ToList() ?? [];
        var spanSum = cells.Sum(c => c.Span);
        if (spanSum != gridColumns.Count)
        {
            throw new ArgumentException(
                $"The cells' spans sum to {spanSum} but the parent table has {gridColumns.Count} grid "
                + $"column(s); {rowNoun} must cover the grid exactly (use a spacer cell '-' with a span "
                + "for the unused width).",
                nameof(cells));
        }

        var itemSdt = parent.GetFirstChild<SdtContentRow>()?.Elements<SdtRow>()
                .FirstOrDefault(s => s.SdtProperties?.GetFirstChild<Office2013Word.SdtRepeatedSectionItem>() is not null)
            ?? throw new ArgumentException(
                $"Repeater {parentRepeaterId} has no repeatingSectionItem row to add {rowNoun} to "
                + "(unexpected shape).",
                nameof(parentRepeaterId));
        var itemContent = itemSdt.GetFirstChild<SdtContentRow>()
            ?? throw new ArgumentException(
                $"Repeater {parentRepeaterId}'s item has no content container (unexpected shape).",
                nameof(parentRepeaterId));

        var tables = TableGridNavigator.Tables(body);
        int? tableIndex = null;
        for (var i = 0; i < tables.Count; i++)
        {
            if (ReferenceEquals(tables[i], table))
            {
                tableIndex = i;
                break;
            }
        }

        return new RepeaterItemContext(parent, itemContent, table, gridColumns, tableIndex);
    }

    /// <summary>
    /// Each cell's <c>w:tcW</c> for a row laid out on an existing table's grid: the sum of the grid-column
    /// widths the cell covers, walking <paramref name="cells"/>' spans left to right (an unparsable/absent
    /// <c>w:gridCol</c> width falls back to 2000 twips per column). Shared by <see cref="InsertRepeaterRow"/>
    /// and <see cref="TableStructureEditor.InsertStaticRow"/> — the same arithmetic, kept in one place. The
    /// caller has already verified the spans sum to the grid count, so indexing cannot overrun.
    /// </summary>
    internal static int[] ComputeCellWidthsFromGrid(IReadOnlyList<GridColumn> gridColumns, IReadOnlyList<RepeaterRowCell> cells)
    {
        var widths = new int[cells.Count];
        var cursor = 0;
        for (var i = 0; i < cells.Count; i++)
        {
            var width = 0;
            for (var g = cursor; g < cursor + cells[i].Span; g++)
            {
                width += int.TryParse(gridColumns[g].Width?.Value, out var w) && w > 0 ? w : 2000;
            }

            widths[i] = width;
            cursor += cells[i].Span;
        }

        return widths;
    }

    /// <summary>
    /// Inserts a NEW plain (unbound) table of <paramref name="rows"/> × <paramref name="columns"/> empty
    /// single-column cells as a block at <paramref name="location"/> — the missing counterpart to
    /// <see cref="InsertRepeaterTable"/> for the NON-repeating blocks every real BC layout is built from
    /// (side-by-side address columns, label/value header-info grids, right-anchored totals blocks).
    /// Borderless by default — the corpus shape for those blocks — or with the same explicit single-line
    /// border set <see cref="SdtFactory.BuildRepeaterTable"/> emits when <paramref name="withBorders"/> is
    /// true. Each cell carries an explicit <c>w:tcW</c> matching its grid column, and an optional per-column
    /// paragraph justification seeds each cell's alignment (which <see cref="CellTextEditor.SetCellText"/>
    /// and a <c>tableCell</c>-located <see cref="InsertField"/> both preserve afterwards). The result's
    /// <see cref="TableEditResult.TableIndex"/> is the new table's 0-based document-order index — exactly
    /// the coordinate <c>set_cell_text</c> / <c>insert_field(tableCell …)</c> / <c>set_column_widths</c>
    /// address next.
    /// </summary>
    /// <exception cref="ArgumentException">
    /// <paramref name="rows"/>/<paramref name="columns"/> out of range; a width/alignment list whose count
    /// doesn't equal <paramref name="columns"/>; an unrecognized alignment; or
    /// <paramref name="location"/>.Part is not <see cref="LayoutPart.Body"/> (plain tables in
    /// headers/footers are deferred with the same v1 scope rule as <see cref="InsertRepeaterTable"/>).
    /// </exception>
    /// <exception cref="NotFoundException">Propagated from <see cref="LocationResolver.Resolve"/>.</exception>
    public static TableEditResult InsertPlainTable(
        WordprocessingDocument doc,
        int rows,
        int columns,
        Location location,
        IReadOnlyList<int>? columnWidths = null,
        IReadOnlyList<string>? columnAlignments = null,
        bool withBorders = false)
    {
        ArgumentNullException.ThrowIfNull(doc);
        ArgumentNullException.ThrowIfNull(location);

        if (rows is < 1 or > 100)
        {
            throw new ArgumentException($"rows must be between 1 and 100 (got {rows}).", nameof(rows));
        }

        if (columns is < 1 or > 30)
        {
            throw new ArgumentException($"columns must be between 1 and 30 (got {columns}).", nameof(columns));
        }

        if (location.Part != LayoutPart.Body)
        {
            throw new ArgumentException(
                "insert_table only supports the main document body in v1 (a plain table in a header/footer "
                + "is deferred with the same scope rule as insert_repeater_table); omit layoutPart (or pass "
                + "'body').",
                nameof(location));
        }

        if (columnWidths is not null && columnWidths.Count != columns)
        {
            throw new ArgumentException(
                $"columnWidths has {columnWidths.Count} entr{(columnWidths.Count == 1 ? "y" : "ies")} but the "
                + $"table has {columns} column(s); supply exactly one width per column, or omit columnWidths "
                + "for an even split of the default content width.",
                nameof(columnWidths));
        }

        if (columnAlignments is not null && columnAlignments.Count != columns)
        {
            throw new ArgumentException(
                $"columnAlignments has {columnAlignments.Count} entr{(columnAlignments.Count == 1 ? "y" : "ies")} "
                + $"but the table has {columns} column(s); supply exactly one of left/center/right per column, "
                + "or omit columnAlignments entirely.",
                nameof(columnAlignments));
        }

        // Even split of the document's own content width (page width minus margins, read from the body
        // sectPr; the corpus-standard 10206 twips when unset) when no explicit widths are given — a plain
        // top-section block spans the full content width in every corpus layout.
        var contentWidth = ResolveContentWidthTwips(doc);
        var widths = columnWidths?.ToArray()
            ?? Enumerable.Repeat(contentWidth / columns, columns).ToArray();
        var justifications = columnAlignments?
            .Select(a => (JustificationValues?)CellTextEditor.ParseAlignment(a, nameof(columnAlignments)))
            .ToArray();

        var tblPr = new TableProperties();
        if (withBorders)
        {
            tblPr.Append(new TableBorders
            {
                TopBorder = new TopBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                LeftBorder = new LeftBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                BottomBorder = new BottomBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                RightBorder = new RightBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                InsideHorizontalBorder =
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                InsideVerticalBorder =
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
            });
        }

        var grid = new TableGrid(widths.Select(w =>
            (OpenXmlElement)new GridColumn { Width = w.ToString(System.Globalization.CultureInfo.InvariantCulture) }));

        var table = new Table(tblPr, grid);
        for (var r = 0; r < rows; r++)
        {
            var row = new TableRow();
            for (var c = 0; c < columns; c++)
            {
                // Compact, single-spaced cells — the shape real BC layout cells render with (their
                // paragraphs override the document's airy defaults explicitly). Without this a blank
                // document's Word defaults (~8pt after each paragraph, 1.08 line) inflate every row.
                var paragraph = new Paragraph
                {
                    ParagraphProperties = new ParagraphProperties
                    {
                        SpacingBetweenLines = new SpacingBetweenLines { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto },
                    },
                };
                if (justifications?[c] is { } jc)
                {
                    paragraph.ParagraphProperties.Justification = new Justification { Val = jc };
                }

                row.AppendChild(new TableCell(
                    new TableCellProperties(new TableCellWidth
                    {
                        Type = TableWidthUnitValues.Dxa,
                        Width = widths[c].ToString(System.Globalization.CultureInfo.InvariantCulture),
                    }),
                    paragraph));
            }

            table.AppendChild(row);
        }

        var anchor = LocationResolver.Resolve(location, doc);
        anchor.InsertBlock(table);
        SeparateFromAdjacentTables(table);

        // The index every follow-up cell edit will address: the same flattened document-order numbering
        // get_layout_info / TableGridNavigator use for this part.
        var (root, _) = LocationResolver.ResolvePart(doc, location.Part, location.PartName);
        var tables = TableGridNavigator.Tables(root);
        var tableIndex = -1;
        for (var i = 0; i < tables.Count; i++)
        {
            if (ReferenceEquals(tables[i], table))
            {
                tableIndex = i;
                break;
            }
        }

        if (tableIndex < 0)
        {
            throw new InvalidOperationException("The newly inserted table was not found in its own part (unexpected).");
        }

        return new TableEditResult
        {
            Operation = "insert_table",
            Part = anchor.PartName,
            TableIndex = tableIndex,
            ColumnIndex = null,
            RowsAffected = rows,
            ColumnCountBefore = 0,
            ColumnCountAfter = columns,
            Summary = $"Inserted a plain {rows}x{columns} table at {DescribeLocation(location, anchor.PartName)}; "
                + $"it is table {tableIndex} - address its cells with set_cell_text/insert_field/insert_label "
                + "(locationType 'tableCell') using that index.",
        };
    }

    /// <summary>Corpus A4 content width in twips — the full-width fallback for a new plain table's grid.</summary>
    private const int DefaultPlainTableWidthTwips = 10206;

    /// <summary>
    /// The body's usable content width (page width minus left/right margins) from its trailing sectPr;
    /// the corpus-standard <see cref="DefaultPlainTableWidthTwips"/> when page size/margins are absent
    /// or nonsensical.
    /// </summary>
    private static int ResolveContentWidthTwips(WordprocessingDocument doc)
    {
        var sectPr = doc.MainDocumentPart?.Document?.Body?.Elements<SectionProperties>().FirstOrDefault();
        var pageWidth = (int?)sectPr?.GetFirstChild<PageSize>()?.Width?.Value;
        var margins = sectPr?.GetFirstChild<PageMargin>();
        if (pageWidth is null || margins is null)
        {
            return DefaultPlainTableWidthTwips;
        }

        var width = pageWidth.Value - (int)(margins.Left?.Value ?? 0) - (int)(margins.Right?.Value ?? 0);
        return width > 1000 ? width : DefaultPlainTableWidthTwips;
    }

    /// <summary>
    /// Applies the optional RUN-level formatting knobs (bold, font size) to a freshly built control: onto
    /// every run inside its content (what actually renders) AND its <c>w:sdtPr/w:rPr</c> (what Word styles
    /// the placeholder/future content with; the corpus-verified child order puts rPr first). Alignment is
    /// a paragraph/cell concern and rejected here — seed it via <c>insert_table</c>'s columnAlignments or
    /// <c>set_cell_text</c>'s alignment instead, which the control then simply sits inside.
    /// </summary>
    /// <summary>
    /// Shared validation for the inline run-format knobs: rejects an alignment (a paragraph/cell concern, not
    /// a run one) and an out-of-range font size, and converts the size to the half-points OOXML wants.
    /// Returns <c>false</c> when there is nothing to apply, so callers can skip the walk entirely.
    /// </summary>
    private static bool TryPrepareRunFormat(CellTextFormat? format, string opName, out string? halfPoints)
    {
        halfPoints = null;

        if (format is null)
        {
            return false;
        }

        if (format.Alignment is not null)
        {
            throw new ArgumentException(
                $"alignment is not supported on {opName} (the inserted content is inline; alignment belongs "
                + "to the paragraph/cell around it). Seed the cell's alignment via insert_table's "
                + "columnAlignments or set_cell_text's alignment instead.",
                nameof(format));
        }

        if (format.Bold is null && format.FontSizePoints is null)
        {
            return false;
        }

        if (format.FontSizePoints is { } points)
        {
            if (points is < 4 or > 96)
            {
                throw new ArgumentException(
                    $"fontSizePoints must be between 4 and 96 (got {points}).", nameof(format));
            }

            halfPoints = ((int)Math.Round(points * 2)).ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        return true;
    }

    /// <summary>Applies the inline run-format knobs to a single plain run (see <see cref="InsertText"/>).</summary>
    private static void ApplyRunFormat(Run run, CellTextFormat? format, string opName)
    {
        if (!TryPrepareRunFormat(format, opName, out var halfPoints))
        {
            return;
        }

        run.RunProperties ??= new RunProperties();
        ApplyRunProperties(run.RunProperties, format!.Bold, halfPoints);
    }

    /// <summary>
    /// Applies the inline run-format knobs to a freshly built control: onto every run inside its content
    /// AND its <c>w:sdtPr/w:rPr</c>. Internal (with a caller-supplied <paramref name="opName"/> for the
    /// rejection message) so <see cref="TableStructureEditor.InsertStaticRow"/> can style its bound cells
    /// through the exact same path <c>insert_field</c>/<c>insert_label</c> use.
    /// </summary>
    internal static void ApplyControlRunFormat(SdtRun sdt, CellTextFormat? format, string opName = "insert_field/insert_label")
    {
        if (!TryPrepareRunFormat(format, opName, out var halfPoints))
        {
            return;
        }

        var targets = sdt.Descendants<Run>().Select(r => (OpenXmlElement)r).ToList();
        foreach (var target in targets)
        {
            var run = (Run)target;
            run.RunProperties ??= new RunProperties();
            ApplyRunProperties(run.RunProperties, format!.Bold, halfPoints);
        }

        if (sdt.SdtProperties is { } sdtPr)
        {
            var rPr = sdtPr.GetFirstChild<RunProperties>();
            if (rPr is null)
            {
                rPr = new RunProperties();
                sdtPr.PrependChild(rPr);
            }

            ApplyRunProperties(rPr, format!.Bold, halfPoints);
        }
    }

    private static void ApplyRunProperties(RunProperties rPr, bool? bold, string? halfPoints)
    {
        if (bold is { } b)
        {
            rPr.Bold = b ? new Bold() : null;
        }

        if (halfPoints is not null)
        {
            rPr.FontSize = new FontSize { Val = halfPoints };
            rPr.FontSizeComplexScript = new FontSizeComplexScript { Val = halfPoints };
        }
    }

    /// <summary>
    /// Word renders two adjacent <c>w:tbl</c> siblings as ONE merged table — authoring in Word itself
    /// always keeps a paragraph between them. A freshly inserted block table therefore gets an empty
    /// separator paragraph on each side that directly touches another table, so stacking blocks
    /// (address grid, info grid, lines, totals) keeps them visually distinct.
    /// </summary>
    private static void SeparateFromAdjacentTables(Table table)
    {
        if (table.PreviousSibling() is Table)
        {
            table.InsertBeforeSelf(new Paragraph());
        }

        if (table.NextSibling() is Table)
        {
            table.InsertAfterSelf(new Paragraph());
        }
    }

    private static EditResult Insert(
        WordprocessingDocument doc, string datasetPath, Location location, bool isLabel, CellTextFormat? format = null)
    {
        ArgumentNullException.ThrowIfNull(doc);
        ArgumentNullException.ThrowIfNull(location);

        var schema = SchemaProvider.FromLayout(doc);
        var id = GenerateUniqueId(doc);

        var sdt = isLabel
            ? SdtFactory.BuildLabel(schema, datasetPath, id: id)
            : SdtFactory.BuildField(schema, datasetPath, id: id);
        ApplyControlRunFormat(sdt, format);

        var scaffolded = EnsureTargetPartExists(doc, location);

        // LocationResolver.Resolve picks whichever part location.Part/PartName names (the main body by
        // default) and reports it back via anchor.PartName, so this method's own Part/Summary are always
        // accurate — not assumed to be document.xml.
        var anchor = LocationResolver.Resolve(location, doc);
        anchor.InsertInline(sdt);

        var alias = sdt.SdtProperties!.GetFirstChild<SdtAlias>()?.Val?.Value;
        var xpath = sdt.SdtProperties!.GetFirstChild<DataBinding>()?.XPath?.Value;
        var kind = isLabel ? ControlKind.Label : ControlKind.Field;

        return new EditResult
        {
            Operation = isLabel ? "insert_label" : "insert_field",
            ControlId = id,
            Alias = alias,
            XPath = xpath,
            Kind = kind.ToString(),
            Part = anchor.PartName,
            Summary = $"Inserted {kind.ToString().ToLowerInvariant()} '{datasetPath}' (id {id}) at "
                + $"{DescribeLocation(location, anchor.PartName)}."
                + (scaffolded
                    ? $" The layout had no {location.Part.ToString().ToLowerInvariant()} part, so an empty "
                      + $"{anchor.PartName} was created and wired into the page setup first."
                    : string.Empty),
        };
    }

    /// <summary>
    /// Creates the empty header/footer part this insert targets when the layout has none at all, returning
    /// <c>true</c> when it had to (see <see cref="HeaderFooterScaffold"/> for the shape, and
    /// why the alternative — failing <c>not_found</c> on a layout authored from scratch — was a dead end for
    /// the caller). Deliberately narrow: only <see cref="LocationKind.DocumentEnd"/> qualifies, since it is
    /// the only location that can resolve inside a part that was empty a moment ago. An
    /// <c>afterControl</c>/<c>tableCell</c>/<c>atText</c> location naming a part that does not exist could
    /// not resolve against a freshly scaffolded one either, so those keep their existing (accurate)
    /// <see cref="NotFoundException"/> rather than silently gaining an empty part as a side effect. A
    /// location that names a SPECIFIC <see cref="Location.PartName"/> is likewise left alone — inventing a
    /// part under a caller-chosen file name would be guessing at what they meant.
    /// </summary>
    private static bool EnsureTargetPartExists(WordprocessingDocument doc, Location location) =>
        location.Type == LocationKind.DocumentEnd
        && location.PartName is null
        && HeaderFooterScaffold.EnsureExists(doc, location.Part);

    /// <summary>
    /// Removes the control (any kind — run/block/row/cell/run-ruby) whose <c>w:sdtPr/w:id/@w:val</c> equals
    /// <paramref name="controlId"/>, searching the main document body plus every header and footer part (in
    /// that order). When <paramref name="keepText"/> is true, the control is unwrapped rather than deleted
    /// outright: the children of its own content element (e.g. the runs of a field/label, the table row of
    /// a row-level control) are spliced into its former position and only the <c>w:sdt</c> wrapper itself
    /// is discarded — see <see cref="UnwrapSdt"/>. When false, the control and all of its content is removed.
    /// </summary>
    /// <remarks>
    /// CELL-LEVEL CONTROLS ARE NEVER DELETED WHOLESALE. A cell-level control (<see cref="SdtCell"/> — the
    /// shape BC uses for header address fields like <c>CustomerAddress1..6</c>) wraps a whole <c>w:tc</c>
    /// table cell, which is one column of the table grid. A naive <c>target.Remove()</c> would take that
    /// <c>w:tc</c> with it, leaving the row with fewer cells than the grid declares — a visually broken table
    /// with a missing column (this passes <see cref="OpenXmlValidator"/> because a row MAY have any number of
    /// cells, so the corruption is silent). Instead the cell is ALWAYS preserved: the <c>w:sdt</c> wrapper is
    /// unwrapped so the <c>w:tc</c> survives in place, and <paramref name="keepText"/> then decides only
    /// whether the cell keeps its former text (<c>true</c>) or is emptied to a blank cell (<c>false</c>).
    /// Either way the column stays. To truly drop a column you must remove its cell from every row and adjust
    /// the grid — deliberately out of scope for a single-control edit.
    /// </remarks>
    /// <exception cref="NotFoundException">No control with <paramref name="controlId"/> was found.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="keepText"/> is true and the target is a repeating-section control (see remarks).
    /// </exception>
    public static EditResult RemoveControl(WordprocessingDocument doc, int controlId, bool keepText)
    {
        ArgumentNullException.ThrowIfNull(doc);

        var found = FindControlById(doc, controlId)
            ?? throw new NotFoundException(
                $"No control with id {controlId} was found in the document (main body, headers, or footers).",
                NotFoundTarget.Control);

        var (target, part) = found;

        // Capture everything the result needs to describe BEFORE mutating. This is not strictly required
        // for the alias/xpath reads (w:sdtPr itself is untouched by UnwrapSdt, which only ever touches the
        // sdt's content element), but reading before removal/unwrap is the simplest way to make that true
        // regardless of how the unwrap is implemented.
        var alias = SdtInspector.ReadAlias(target);
        var xpath = SdtInspector.ReadXPath(target);
        var kind = SdtInspector.ClassifyControlKind(target).ToString();

        if (keepText && SdtInspector.IsRepeater(target))
        {
            // Unwrapping a repeatingSection would splice its repeatingSectionItem row template directly
            // into the repeater's former position, orphaned from any enclosing repeatingSection - that is
            // structurally valid OOXML (so it would still save and pass OpenXmlValidator), but it breaks
            // the BC repeater shape LayoutValidator.Quick checks for (repeater-shape: every
            // repeatingSectionItem must be enclosed by a repeatingSection), i.e. a "successful" edit that
            // immediately fails its own post-edit quick validation. Reject it up front instead.
            throw new ArgumentException(
                $"Control {controlId} is a repeating-section control; removing it with keepText=true would "
                + "orphan its repeatingSectionItem row template from any enclosing repeatingSection, breaking "
                + "the BC repeater shape. Use keepText=false to remove the whole repeater instead.",
                nameof(keepText));
        }

        bool keptColumn = false;
        if (target is SdtCell sdtCell)
        {
            // Cell-level control: preserve the w:tc column no matter what (see remarks). Capture the cell
            // BEFORE unwrapping, then empty it when the caller did not ask to keep its text.
            var cell = sdtCell.SdtContentCell?.GetFirstChild<TableCell>();
            UnwrapSdt(target);
            if (cell is not null)
            {
                keptColumn = true;
                if (!keepText)
                {
                    ClearCellContent(cell);
                }
            }
        }
        else if (keepText)
        {
            UnwrapSdt(target);
        }
        else
        {
            // A block-level (or inline) control can be the SOLE content of a table cell — the shape BC uses
            // for the amount/quantity fields in a line-items row, where each field is an SdtBlock sitting
            // directly in its own w:tc. Removing it wholesale would leave an empty w:tc (just its w:tcPr, no
            // paragraph): invalid OOXML that Word treats as a corrupt document, even though OpenXmlValidator
            // silently accepts it (the same class of silent corruption as PlainTextNestingGuard's, so the
            // pre-save OpenXmlValidator gate cannot catch it either). Preserve the cell as a valid empty
            // column instead — the same guarantee the SdtCell branch above makes.
            //
            // A row-level control (SdtRow, e.g. a BC repeater's repeatingSection/repeatingSectionItem, or a
            // plain row-level sdt) is a direct child of w:tbl in place of a w:tr — removing it wholesale
            // when it is the table's ONLY row would leave a rowless <w:tbl>: schema-legal (OpenXmlValidator
            // accepts it) but Word-hostile (Word silently drops the whole table on open). Remove the table
            // itself in that case rather than leave dead markup nothing can address.
            //
            // That table can itself be the SOLE content of an OUTER table's cell (a nested data table — a
            // real BC shape). Removing it wholesale would then repeat the exact same hazard one level up:
            // an empty outer <w:tc> (just its w:tcPr, no paragraph), schema-legal to OpenXmlValidator but
            // Word-corrupting. Capture the table's own parent cell BEFORE removing the table (Parent is
            // cleared once removed) and apply the SAME EnsureCellNotEmpty guarantee to it, mirroring the
            // parentCell branch above.
            var parentCell = target.Parent as TableCell;
            var parentTable = target.Parent as Table;
            var grandparentCell = parentTable?.Parent as TableCell;
            target.Remove();
            if (parentCell is not null && EnsureCellNotEmpty(parentCell))
            {
                keptColumn = true;
            }
            else if (parentTable is not null && TableGridNavigator.Rows(parentTable).Count == 0)
            {
                parentTable.Remove();
                if (grandparentCell is not null && EnsureCellNotEmpty(grandparentCell))
                {
                    keptColumn = true;
                }
            }
        }

        return new EditResult
        {
            Operation = "remove_control",
            ControlId = controlId,
            Alias = alias,
            XPath = xpath,
            Kind = kind,
            Part = part,
            Summary = BuildRemoveSummary(controlId, keepText, keptColumn),
        };
    }

    private static string BuildRemoveSummary(int controlId, bool keepText, bool keptColumn)
    {
        if (keptColumn)
        {
            return keepText
                ? $"Removed control id {controlId} (kept its table cell and its text in place — the column is preserved)."
                : $"Removed control id {controlId} and emptied its table cell (the column is preserved).";
        }

        return keepText
            ? $"Removed control id {controlId} (kept its content in place)."
            : $"Removed control id {controlId}.";
    }

    /// <summary>
    /// Empties a table cell's visible content while leaving it a valid, correctly-sized column: the cell's
    /// own <c>w:tcPr</c> (which carries its width/formatting) is kept, and its block content is replaced with
    /// a single empty paragraph that preserves the first paragraph's <c>w:pPr</c> (alignment/spacing) when
    /// there was one. A <c>w:tc</c> must contain at least one block-level child, so the empty paragraph is
    /// required for the cell to stay well-formed.
    /// </summary>
    private static void ClearCellContent(TableCell cell)
    {
        var tcPr = cell.GetFirstChild<TableCellProperties>();
        var pPr = cell.Elements<Paragraph>().FirstOrDefault()?.GetFirstChild<ParagraphProperties>()?.CloneNode(true);

        cell.RemoveAllChildren();

        if (tcPr is not null)
        {
            cell.AppendChild(tcPr);
        }

        var emptyParagraph = new Paragraph();
        if (pPr is not null)
        {
            emptyParagraph.AppendChild(pPr);
        }

        cell.AppendChild(emptyParagraph);
    }

    /// <summary>
    /// Guarantees a <c>w:tc</c> still contains at least one paragraph after a control that was its sole
    /// content has been removed. A cell with no <c>w:p</c> (e.g. only its <c>w:tcPr</c> left) is invalid
    /// OOXML that Word treats as a corrupt document, yet <see cref="OpenXmlValidator"/> accepts it silently
    /// — so this is enforced here rather than relying on the pre-save structural gate to catch it. Appends
    /// an empty paragraph only when the cell has no paragraph left, preserving the column; returns
    /// <c>true</c> when it had to add one (i.e. the removal emptied the cell).
    /// </summary>
    private static bool EnsureCellNotEmpty(TableCell cell)
    {
        if (cell.Elements<Paragraph>().Any())
        {
            return false;
        }

        cell.AppendChild(new Paragraph());
        return true;
    }

    /// <summary>
    /// Describes <paramref name="location"/> for a human-readable edit summary. When
    /// <paramref name="resolvedPartName"/> is the main document (<c>document.xml</c>) this is EXACTLY the
    /// pre-Phase-4.3 wording (body-targeted summaries are unchanged, word for word); a header/footer part
    /// gets a short " in {part}" suffix instead, since the whole point of targeting one is to make that
    /// visible in the result.
    /// </summary>
    private static string DescribeLocation(Location location, string resolvedPartName)
    {
        var where = location.Type switch
        {
            LocationKind.DocumentEnd => "the end of the document",
            LocationKind.AfterControl => $"after control {location.ControlId}",
            LocationKind.TableCell => $"table {location.TableIndex} row {location.Row} col {location.Col}",
            LocationKind.AtText => $"the text '{location.SearchText}'",
            _ => location.Type.ToString(),
        };

        return resolvedPartName == "document.xml" ? where : $"{where} in {resolvedPartName}";
    }

    // ---- unique id generation ----

    /// <summary>
    /// Generates a <c>w:id</c> guaranteed not to collide with any <c>w:sdt/w:sdtPr/w:id</c> already present
    /// anywhere in <paramref name="doc"/> (main body, headers, footers). <see cref="SdtFactory"/>'s own
    /// auto-generated ids only avoid colliding with ids it has issued itself earlier in the current process
    /// — not with a specific target document's own pre-existing ids — so that check belongs here instead.
    /// </summary>
    private static int GenerateUniqueId(WordprocessingDocument doc)
    {
        var existing = CollectAllControlIds(doc);
        var random = new Random();
        int candidate;
        do
        {
            candidate = random.Next(int.MinValue, int.MaxValue);
        }
        while (existing.Contains(candidate));

        return candidate;
    }

    /// <summary>
    /// Returns a <c>Func&lt;int&gt;</c> that generates ids guaranteed unique both against <paramref name="doc"/>'s
    /// pre-existing ids (snapshotted once via <see cref="CollectAllControlIds"/>) AND against every id the
    /// SAME closure has already handed out — needed anywhere more than one fresh id is required before any
    /// of them are attached to the document tree (e.g. <see cref="InsertRepeaterTable"/>'s many sdts),
    /// unlike <see cref="GenerateUniqueId"/> which only ever hands out exactly one id per call.
    /// </summary>
    internal static Func<int> MakeIdGenerator(WordprocessingDocument doc)
    {
        var existing = CollectAllControlIds(doc);
        var issued = new HashSet<int>();
        var random = new Random();

        return () =>
        {
            int candidate;
            do
            {
                candidate = random.Next(int.MinValue, int.MaxValue);
            }
            while (existing.Contains(candidate) || !issued.Add(candidate));

            return candidate;
        };
    }

    private static HashSet<int> CollectAllControlIds(WordprocessingDocument doc)
    {
        var main = doc.MainDocumentPart
            ?? throw new InvalidDataException("Layout has no main document part.");

        var ids = new HashSet<int>();

        foreach (var (root, _) in PartWalker.ContentParts(main))
        {
            foreach (var sdt in root.Descendants<SdtElement>())
            {
                var id = SdtInspector.ReadControlId(sdt);
                if (id.HasValue)
                {
                    ids.Add(id.Value);
                }
            }
        }

        return ids;
    }

    // ---- control lookup across document.xml + headers + footers ----

    private static (SdtElement Target, string Part)? FindControlById(WordprocessingDocument doc, int controlId)
    {
        var main = doc.MainDocumentPart
            ?? throw new InvalidDataException("Layout has no main document part.");

        foreach (var (root, partName) in PartWalker.ContentParts(main))
        {
            var match = root.Descendants<SdtElement>().FirstOrDefault(s => SdtInspector.ReadControlId(s) == controlId);
            if (match is not null)
            {
                return (match, partName);
            }
        }

        return null;
    }

    // ---- keepText unwrap: generic across all five concrete Sdt*/SdtContent* kinds ----

    /// <summary>
    /// Replaces <paramref name="sdt"/> with the children of its own content element — whichever of
    /// <see cref="SdtContentRun"/> (on <see cref="SdtRun"/>), <see cref="SdtContentBlock"/> (on
    /// <see cref="SdtBlock"/>), <see cref="SdtContentCell"/> (on <see cref="SdtCell"/>),
    /// <see cref="SdtContentRow"/> (on <see cref="SdtRow"/>), or <see cref="SdtContentRunRuby"/> (on
    /// <see cref="SdtRunRuby"/>, the East Asian "ruby"/phonetic-guide run control) applies to its concrete
    /// kind — preserving their order and position, then discards the (now empty) sdt wrapper. These five
    /// are the complete set of concrete <see cref="SdtElement"/> subclasses in the OOXML wordprocessing
    /// schema, so this switch is exhaustive — there is no sixth kind to fall through to the "no recognized
    /// container" branch below for. Content-model compatibility is guaranteed by construction: each
    /// concrete sdt kind only ever appears where its own content kind is itself legal (an
    /// <see cref="SdtRun"/> only where run content is legal, an <see cref="SdtCell"/> only where a
    /// <see cref="TableCell"/> is legal, etc.), so that content's children are always legal in the sdt's
    /// own position too — real corpus content controls also carry non-run children alongside their visible
    /// text (e.g. <c>w:proofErr</c> spell-check markers), which is exactly why every child is moved, not
    /// just the ones that look like the "real" content.
    /// </summary>
    private static void UnwrapSdt(SdtElement sdt)
    {
        OpenXmlElement? container = sdt switch
        {
            SdtRun r => r.SdtContentRun,
            SdtBlock b => b.SdtContentBlock,
            SdtCell c => c.SdtContentCell,
            SdtRow w => w.SdtContentRow,
            SdtRunRuby ry => ry.SdtContentRunRuby,
            _ => null,
        };

        if (container is null)
        {
            // No recognized content container (or the sdt has no content at all): nothing to keep.
            sdt.Remove();
            return;
        }

        OpenXmlElement anchor = sdt;
        foreach (var child in container.ChildElements.ToList())
        {
            child.Remove();
            anchor.InsertAfterSelf(child);
            anchor = child;
        }

        sdt.Remove();
    }

}
