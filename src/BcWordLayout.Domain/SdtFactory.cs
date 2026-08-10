using System.Globalization;
using System.Text.RegularExpressions;
using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Dw = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Office2013Word = DocumentFormat.OpenXml.Office2013.Word;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;

namespace BcWordLayout.Domain;

/// <summary>
/// Builds Business Central-compatible field/label content controls (<c>w:sdt</c>) as free-standing
/// OpenXml elements. This type does no file I/O and does not insert anything into a document — see
/// <see cref="LocationResolver"/> / <see cref="InsertionAnchor"/> for placement.
/// </summary>
/// <remarks>
/// The exact shape mirrored here (element order, attributes, namespaces) was extracted from a real
/// control in <c>tests/corpus/SalesInvoiceForSubscriptionBilling.docx</c> (the <c>CompanyLegalOffice</c>
/// field, an inline <c>w:sdt</c> whose content is a bare run):
/// <code>
/// &lt;w:sdt&gt;&lt;w:sdtPr&gt;
///   &lt;w:alias w:val="#Nav: /Header/CompanyLegalOffice" /&gt;
///   &lt;w:tag w:val="#Nav: Standard_Sales_Invoice/1306" /&gt;
///   &lt;w:id w:val="1332101128" /&gt;
///   &lt;w:placeholder&gt;&lt;w:docPart w:val="FAF6EB61142E4D80BD1FFD7E376AA00A" /&gt;&lt;/w:placeholder&gt;
///   &lt;w:dataBinding w:prefixMappings="xmlns:ns0='urn:microsoft-dynamics-nav/reports/Standard_Sales_Invoice/1306/'"
///                  w:xpath="/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:CompanyLegalOffice[1]"
///                  w:storeItemID="{AF7A6226-6056-400F-ADDA-E1ADA7C08250}" /&gt;
///   &lt;w:text /&gt;
/// &lt;/w:sdtPr&gt;&lt;w:sdtContent&gt;&lt;w:r&gt;&lt;w:t&gt;CompanyLegalOffice&lt;/w:t&gt;&lt;/w:r&gt;&lt;/w:sdtContent&gt;&lt;/w:sdt&gt;
/// </code>
/// Its sibling label control (<c>CompanyLegalOffice_Lbl</c>, immediately preceding it in the same
/// paragraph) is the same shape — only the naming convention differs — confirming label/field are
/// structurally identical. <c>w:placeholder</c> references either a custom GUID-named glossary entry (as
/// both controls above do) or Word's own built-in <c>DefaultPlaceholder_-1854013440</c> entry (seen on
/// other real corpus controls, e.g. <c>/Header/Line/JobNo_Lbl</c>); this factory always uses the latter
/// (present in every corpus layout's <c>word/glossary/document.xml</c>) so it never has to synthesize
/// glossary parts. Real placeholder run text is simply the leaf field name with no
/// <c>xml:space="preserve"</c> — reproduced exactly (see <see cref="BuildField"/>'s default).
/// </remarks>
public static class SdtFactory
{
    /// <summary>
    /// Word's own built-in "Click here to enter text." glossary docPart id — present by default in every
    /// corpus layout's glossary part, unlike the custom GUID-named placeholders real BC field controls
    /// reference. Used for every control this factory builds, keeping it free of glossary-part concerns.
    /// </summary>
    internal const string DefaultPlaceholderDocPart = "DefaultPlaceholder_-1854013440";

    /// <summary>
    /// Builds an inline plain-text FIELD control bound to <paramref name="datasetPath"/> — a slash-delimited
    /// path from the dataset root using the same convention as <see cref="DatasetColumn.Path"/>/
    /// <see cref="DataItem.Path"/> (e.g. <c>/Header/Line/ItemNo_Line</c>; a leading slash is optional).
    /// </summary>
    /// <param name="schema">
    /// The report's parsed dataset (from <see cref="SchemaProvider.FromLayout(string)"/>), supplying both
    /// the <see cref="ReportIdentity"/> (namespace/report name/id/storeItemID) used in the alias/tag/
    /// dataBinding and the schema tree the path is validated against.
    /// </param>
    /// <param name="datasetPath">Slash-delimited path to a leaf column, e.g. <c>/Header/CustomerAddress1</c>.</param>
    /// <param name="placeholderText">
    /// Visible run text for the control's content. Defaults to the path's leaf segment name — matching
    /// real BC-produced controls, whose placeholder text is simply the bound field's own name.
    /// </param>
    /// <param name="id">
    /// The control's <c>w:id</c>. When omitted, a plain random <see cref="int"/> is generated (see
    /// <see cref="ResolveId"/>) — <c>w:id</c> has no schema-level uniqueness requirement, so a collision
    /// cannot break validation. This factory does NOT check the generated id against anything (not a
    /// specific target document's own pre-existing ids, nor any id it has generated before): doc-scoped
    /// uniqueness — the only scope OOXML or this codebase's own guards ever care about — is the caller's
    /// job. Every production caller already supplies its own doc-scoped-unique id explicitly (see
    /// <see cref="LayoutEditor"/>'s <c>GenerateUniqueId</c>/<c>MakeIdGenerator</c>, which check the real
    /// target document); pass one yourself if uniqueness matters to you too.
    /// </param>
    /// <exception cref="ArgumentException">
    /// <paramref name="datasetPath"/> is label-shaped per the active <see cref="LabelConvention.Current"/>
    /// (by default, ends in <c>Lbl</c>/<c>_Lbl</c> — use <see cref="BuildLabel"/> instead), does not resolve
    /// against <paramref name="schema"/>, resolves to a repeating data item rather than a leaf column, or
    /// <paramref name="schema"/>'s report has no storeItemID (e.g. it was parsed via
    /// <see cref="SchemaProvider.FromSchemaXml"/> rather than <see cref="SchemaProvider.FromLayout(string)"/>).
    /// </exception>
    public static SdtRun BuildField(DatasetTree schema, string datasetPath, string? placeholderText = null, int? id = null)
    {
        var segments = ValidatedSegments(schema, datasetPath, expectLabel: false);
        return Build(schema, segments, datasetPath, placeholderText, id);
    }

    /// <summary>
    /// Builds an inline plain-text LABEL control — structurally identical to <see cref="BuildField"/>
    /// (BC's label/field distinction is purely a naming convention — by default <c>*Lbl</c>/<c>*_Lbl</c>,
    /// see <see cref="LabelConvention"/> for the full, configurable rule); kept as a separate method for
    /// caller intent/clarity. All parameters and exceptions match <see cref="BuildField"/>, except
    /// <paramref name="datasetPath"/> must be label-shaped (the reverse check).
    /// </summary>
    public static SdtRun BuildLabel(DatasetTree schema, string datasetPath, string? placeholderText = null, int? id = null)
    {
        var segments = ValidatedSegments(schema, datasetPath, expectLabel: true);
        return Build(schema, segments, datasetPath, placeholderText, id);
    }

    /// <summary>
    /// Builds an inline PICTURE control bound to <paramref name="datasetPath"/> — the placeholder BC fills
    /// with a real image at render time (e.g. <c>/Header/CompanyPicture</c>, a company logo).
    /// </summary>
    /// <remarks>
    /// Mirrored element for element from the real add-in-authored control in
    /// <c>tests/corpus/StandardSalesQuote.docx</c>'s <c>header3.xml</c> (<c>/Header/CompanyPicture</c>):
    /// <code>
    /// &lt;w:sdt&gt;&lt;w:sdtPr&gt;
    ///   &lt;w:alias w:val="#Nav: /Header/CompanyPicture" /&gt;&lt;w:tag w:val="#Nav: Standard_Sales_Quote/1304" /&gt;
    ///   &lt;w:id w:val="-1330981123" /&gt;&lt;w:dataBinding … /&gt;&lt;w:picture /&gt;
    /// &lt;/w:sdtPr&gt;&lt;w:sdtContent&gt;&lt;w:r&gt;&lt;w:rPr&gt;&lt;w:noProof /&gt;&lt;/w:rPr&gt;&lt;w:drawing&gt;
    ///   &lt;wp:inline&gt;&lt;wp:extent cx="1080000" cy="1080000" /&gt;&lt;wp:docPr id="2" name="Picture 2" /&gt;…
    ///   &lt;pic:blipFill&gt;&lt;a:blip r:embed="rId1" /&gt;&lt;a:stretch&gt;&lt;a:fillRect /&gt;&lt;/a:stretch&gt;&lt;/pic:blipFill&gt;…
    /// &lt;/wp:inline&gt;&lt;/w:drawing&gt;&lt;/w:r&gt;&lt;/w:sdtContent&gt;&lt;/w:sdt&gt;
    /// </code>
    /// Two differences from <see cref="BuildField"/> are load-bearing, both taken from that capture: the
    /// marker element is <c>w:picture</c> rather than <c>w:text</c>, and a picture control carries NO
    /// <c>w:placeholder</c> at all (so, unlike every other control this factory builds, it depends on
    /// nothing in the glossary part). The size is the corpus's own 1080000×1080000 EMU (3 cm square) unless
    /// the caller overrides it, and the blip must reference a REAL image part relationship in the hosting
    /// part — Word treats a dangling <c>r:embed</c> as a corrupt document — which is why
    /// <paramref name="embedRelationshipId"/> is required rather than optional (see
    /// <see cref="LayoutEditor.InsertPicture"/>, which adds a <see cref="PlaceholderImage"/> part and passes
    /// its id).
    /// </remarks>
    /// <param name="schema">The report's parsed dataset (see <see cref="BuildField"/>).</param>
    /// <param name="datasetPath">Slash-delimited path to the picture's leaf column, e.g. <c>/Header/CompanyPicture</c>.</param>
    /// <param name="embedRelationshipId">
    /// Relationship id of an <see cref="DocumentFormat.OpenXml.Packaging.ImagePart"/> that ALREADY exists in
    /// the part this control will be inserted into.
    /// </param>
    /// <param name="drawingId">
    /// The drawing's <c>wp:docPr/@id</c>. Must be unique across the whole document — Word (and
    /// <see cref="DocumentFormat.OpenXml.Validation.OpenXmlValidator"/>'s <c>Sem_UniqueAttributeValue</c>)
    /// reject a duplicate; see <see cref="LayoutEditor.InsertPicture"/>, which scans for a free one.
    /// </param>
    /// <param name="widthEmu">Frame width in EMU (914400 per inch, 36000 per mm); the corpus value is 1080000.</param>
    /// <param name="heightEmu">Frame height in EMU; the corpus value is 1080000.</param>
    /// <param name="id">The control's <c>w:id</c> (see <see cref="BuildField"/>).</param>
    /// <exception cref="ArgumentException">
    /// Same rules as <see cref="BuildField"/> (<paramref name="datasetPath"/> must be a non-label-shaped leaf
    /// column that resolves against <paramref name="schema"/>, which must carry a storeItemID), plus a
    /// non-positive <paramref name="widthEmu"/>/<paramref name="heightEmu"/> or an empty
    /// <paramref name="embedRelationshipId"/>.
    /// </exception>
    public static SdtRun BuildPicture(
        DatasetTree schema,
        string datasetPath,
        string embedRelationshipId,
        uint drawingId,
        long widthEmu = CorpusPictureExtentEmu,
        long heightEmu = CorpusPictureExtentEmu,
        int? id = null)
    {
        if (string.IsNullOrEmpty(embedRelationshipId))
        {
            throw new ArgumentException(
                "A picture control's blip must reference a real image part relationship; "
                + $"{nameof(embedRelationshipId)} was empty.",
                nameof(embedRelationshipId));
        }

        if (widthEmu <= 0 || heightEmu <= 0)
        {
            throw new ArgumentException(
                $"A picture's frame must be positive (got {widthEmu}x{heightEmu} EMU).", nameof(widthEmu));
        }

        var segments = ValidatedSegments(schema, datasetPath, expectLabel: false);
        var sdtPr = BuildSdtProperties(schema, segments, id, new SdtContentPicture(), withPlaceholder: false);

        return new SdtRun(sdtPr, new SdtContentRun(BuildPictureRun(embedRelationshipId, drawingId, widthEmu, heightEmu)));
    }

    /// <summary>
    /// The corpus's own picture frame size in EMU (1080000 = 3 cm), from
    /// <c>StandardSalesQuote.docx</c>'s <c>CompanyPicture</c> control — the default for a newly authored one.
    /// </summary>
    public const long CorpusPictureExtentEmu = 1080000;

    /// <summary>English Metric Units per millimetre (914400 per inch / 25.4) — the caller-facing size unit.</summary>
    public const long EmuPerMillimetre = 36000;

    /// <summary>
    /// Builds a complete repeater TABLE for <paramref name="dataItemPath"/> (a repeating data item, e.g.
    /// <c>/Header/Line</c>) with one column per entry in <paramref name="columns"/>: a header row of
    /// label/static-text cells, plus a single data row — wrapped in a row-level <c>w15:repeatingSection</c>
    /// / <c>w15:repeatingSectionItem</c> pair — of per-column <see cref="BuildField"/> controls.
    /// </summary>
    /// <remarks>
    /// The shape mirrored here was extracted from the real <c>Header/Line</c> table in
    /// <c>tests/corpus/SalesInvoiceForSubscriptionBilling.docx</c>: a plain <c>w:tr</c> header row (some cells are
    /// cell-level label <c>SdtCell</c> controls, others plain static text) directly followed, as a sibling
    /// row-level <c>w:sdt</c> in the same <c>w:tbl</c>, by
    /// <c>SdtRow(w15:repeatingSection, w15:dataBinding unindexed) &gt; SdtContentRow &gt;
    /// SdtRow(w15:repeatingSectionItem) &gt; SdtContentRow &gt; w:tr</c> (each data cell a fully-indexed
    /// <see cref="BuildField"/> control). Real corpus <c>w:sdtPr</c> order for the outer/inner repeater sdt
    /// is <c>rPr, alias, tag, id, dataBinding, repeatingSection</c> / <c>rPr, id, placeholder,
    /// repeatingSectionItem</c>; the decorative run-properties (<c>w:rPr</c>) are cosmetic only (absent
    /// from the simpler corpus field example this factory already mirrors in <see cref="Build"/>) and are
    /// omitted here for the same reason. Borders follow <see cref="RepeaterTableOptions.Look"/>: the default
    /// <see cref="TableBorderLook.Bc"/> mirrors the real corpus (NO table-level <c>w:tblBorders</c>; one per-cell
    /// <c>w:tcBorders</c> bottom rule under each header cell), while <see cref="TableBorderLook.Grid"/> emits the
    /// explicit full border grid that makes a freshly inserted table visible without depending on a style.
    /// </remarks>
    /// <param name="schema">The report's parsed dataset (see <see cref="BuildField"/>).</param>
    /// <param name="dataItemPath">
    /// Slash-delimited path to a repeating, non-system data item, e.g. <c>/Header/Line</c> (a leading
    /// slash is optional).
    /// </param>
    /// <param name="columns">
    /// Leaf column names of <paramref name="dataItemPath"/> (not full paths — just the leaf name, e.g.
    /// <c>ItemNo_Line</c>), one per table column, in order. At least one is required.
    /// </param>
    /// <param name="options">Header/style/column-width options — see <see cref="RepeaterTableOptions"/>.</param>
    /// <param name="nextId">
    /// Invoked once per generated sdt's <c>w:id</c> (every header label, every data field, plus the
    /// repeatingSectionItem and repeatingSection wrappers themselves); callers own uniqueness (see
    /// <see cref="LayoutEditor.InsertRepeaterTable"/>, which guarantees uniqueness across the whole doc).
    /// </param>
    /// <exception cref="ArgumentException">
    /// <paramref name="dataItemPath"/> does not resolve, resolves to a leaf column instead of a data item,
    /// or resolves to a system data item; <paramref name="columns"/> is empty, contains an entry that is
    /// not a leaf column of that data item, or (indirectly, via <see cref="BuildField"/>) is label-shaped;
    /// <paramref name="options"/>.ColumnWidths is supplied with a count that does not match
    /// <paramref name="columns"/>; or <paramref name="schema"/>'s report has no storeItemID.
    /// </exception>
    public static Table BuildRepeaterTable(
        DatasetTree schema,
        string dataItemPath,
        IReadOnlyList<string> columns,
        RepeaterTableOptions options,
        Func<int> nextId)
    {
        ArgumentNullException.ThrowIfNull(schema);
        ArgumentNullException.ThrowIfNull(columns);
        ArgumentNullException.ThrowIfNull(options);
        ArgumentNullException.ThrowIfNull(nextId);

        if (columns.Count == 0)
        {
            throw new ArgumentException("At least one column is required.", nameof(columns));
        }

        if (options.ColumnWidths is not null && options.ColumnWidths.Count != columns.Count)
        {
            throw new ArgumentException(
                $"options.{nameof(RepeaterTableOptions.ColumnWidths)} has {options.ColumnWidths.Count} "
                + $"entries but there are {columns.Count} columns; supply exactly one width per column, or "
                + $"omit {nameof(RepeaterTableOptions.ColumnWidths)} for an even default.",
                nameof(options));
        }

        var report = schema.Report;
        if (string.IsNullOrEmpty(report.StoreItemId))
        {
            throw new ArgumentException(
                $"schema.Report.{nameof(ReportIdentity.StoreItemId)} is null/empty; build the schema via "
                + $"{nameof(SchemaProvider)}.{nameof(SchemaProvider.FromLayout)} (not "
                + $"{nameof(SchemaProvider.FromSchemaXml)}) so a real storeItemID is available to bind against.",
                nameof(schema));
        }

        var dataItem = ResolveRepeaterDataItem(schema, dataItemPath);
        foreach (var column in columns)
        {
            ValidateColumnOfItem(dataItem, dataItemPath, column);
        }

        var widths = ResolveColumnWidths(options.ColumnWidths, columns.Count);

        if (options.ColumnAlignments is not null && options.ColumnAlignments.Count != columns.Count)
        {
            throw new ArgumentException(
                $"options.{nameof(RepeaterTableOptions.ColumnAlignments)} has {options.ColumnAlignments.Count} "
                + $"entries but there are {columns.Count} columns; supply exactly one of left/center/right per "
                + $"column, or omit {nameof(RepeaterTableOptions.ColumnAlignments)} entirely.",
                nameof(options));
        }

        var headerCells = new List<TableCell>(columns.Count);
        var dataCells = new List<TableCell>(columns.Count);
        for (var i = 0; i < columns.Count; i++)
        {
            var column = columns[i];
            var width = widths[i];

            var headerCell = BuildHeaderCell(schema, dataItem, column, options, nextId, width);

            // The BC look's one structural rule: a ½-pt line under the header row, drawn per cell rather
            // than by a table-level border grid (see TableBorderLook's corpus evidence).
            if (options.Look == TableBorderLook.Bc)
            {
                TableStructureEditor.ApplyCellBorders(headerCell, new CellBorderOptions { Bottom = true });
            }

            headerCells.Add(headerCell);

            var field = BuildField(schema, $"{dataItem.Path}/{column}", id: nextId());
            dataCells.Add(WrapCell(new Paragraph(field), width));

            // Real BC line tables right-align their numeric columns; the alignment applies to the column
            // as a whole - the header cell AND the data cell.
            if (options.ColumnAlignments?[i] is { } alignment)
            {
                var jc = CellTextEditor.ParseAlignment(alignment, nameof(options));
                foreach (var paragraph in new[] { headerCells[i], dataCells[i] }
                             .Select(cell => cell.Descendants<Paragraph>().First()))
                {
                    paragraph.ParagraphProperties ??= new ParagraphProperties();
                    paragraph.ParagraphProperties.Justification = new Justification { Val = jc };
                }
            }
        }

        var tblPr = BuildTableProperties(options, widths);
        var tblGrid = new TableGrid(widths.Select(w => (OpenXmlElement)new GridColumn { Width = w.ToString(CultureInfo.InvariantCulture) }));

        var headerRowChildren = new List<OpenXmlElement> { new TableRowProperties(new TableHeader()) };
        headerRowChildren.AddRange(headerCells);
        var headerRow = new TableRow(headerRowChildren);

        var dataRow = new TableRow(dataCells.Select(c => (OpenXmlElement)c));
        var repeaterRow = BuildRepeaterRow(schema, dataItem, dataRow, nextId);

        return new Table(tblPr, tblGrid, headerRow, repeaterRow);
    }

    /// <summary>
    /// Builds a nested DETAIL ROW repeater — the standard BC shape for per-line detail: inside an outer
    /// repeater's <c>repeatingSectionItem</c>, each nested repeater is a SIBLING ROW after the line's own
    /// <c>w:tr</c>, aligned to the SAME table grid via per-cell <c>gridSpan</c>s (corpus-verified: serial
    /// nos, lot nos, and assembly-component lines all take this shape, nesting several levels deep).
    /// Each <see cref="RepeaterRowCell"/> becomes one <c>w:tc</c> spanning <see cref="RepeaterRowCell.Span"/>
    /// grid columns and carrying its child-item columns as inline controls chained in one paragraph
    /// (label-shaped names become label controls; the corpus chains them with no separator run). The row is
    /// wrapped in the same <c>repeatingSection</c>/<c>repeatingSectionItem</c> pair
    /// <see cref="BuildRepeaterTable"/> emits, bound to <paramref name="dataItemPath"/>.
    /// </summary>
    /// <param name="cellWidths">Explicit <c>w:tcW</c> per cell (the sum of the parent grid columns it covers), one per cell.</param>
    /// <exception cref="ArgumentException">
    /// <paramref name="dataItemPath"/> does not resolve to a repeating, non-system data item; a cell names
    /// a column that is not a leaf column of that item; a span is &lt; 1; or an alignment is unknown.
    /// </exception>
    public static SdtRow BuildDetailRepeaterRow(
        DatasetTree schema,
        string dataItemPath,
        IReadOnlyList<RepeaterRowCell> cells,
        IReadOnlyList<int> cellWidths,
        Func<int> nextId)
    {
        ArgumentNullException.ThrowIfNull(schema);
        ArgumentNullException.ThrowIfNull(cells);
        ArgumentNullException.ThrowIfNull(cellWidths);
        ArgumentNullException.ThrowIfNull(nextId);

        if (cells.Count == 0)
        {
            throw new ArgumentException("At least one cell is required.", nameof(cells));
        }

        var dataItem = ResolveRepeaterDataItem(schema, dataItemPath);

        var row = BuildRowFromCells(schema, cells, cellWidths, nextId, column =>
        {
            ValidateColumnOfItem(dataItem, dataItemPath, column);
            return $"{dataItem.Path}/{column}";
        });

        return BuildRepeaterRow(schema, dataItem, row, nextId);
    }

    /// <summary>
    /// Builds one STATIC (non-repeating) <c>w:tr</c> whose cells carry inline bound controls — the row
    /// shape the stock totals blocks INSIDE a line-items table are made of (GitHub issue #28): every stock
    /// document layout ends its lines table with right-anchored trailing rows (leading empty cells,
    /// <c>gridSpan</c>-merged content cells, bound totals fields) rather than a separate table —
    /// corpus-verified in <c>StandardSalesQuote.docx</c> (1304), <c>StandardPurchaseOrder.docx</c> (1322)
    /// and <c>StandardSalesInvoiceVatSpec.docx</c> (1306), and the same static-rows-in-a-table shape carries
    /// <c>SalespersonCommission.docx</c>'s (115) grand-totals row. Unlike
    /// <see cref="BuildDetailRepeaterRow"/>, each <see cref="RepeaterRowCell.Columns"/> entry here is a
    /// FULL dataset path (e.g. <c>/Header/Totals/TotalAmountIncludingVAT</c> — typically a leaf of a
    /// NON-repeating data item), not a leaf name of some parent item — a totals row binds whatever the
    /// dataset provides, not the repeater's own columns. Label-shaped paths become label controls;
    /// placement is the caller's job (see <see cref="TableStructureEditor.InsertStaticRow"/>).
    /// </summary>
    /// <param name="cellWidths">Explicit <c>w:tcW</c> per cell (the sum of the grid columns it covers), one per cell.</param>
    /// <exception cref="ArgumentException">
    /// <paramref name="cells"/> is empty; a span is &lt; 1; an alignment is unknown; or (propagated from
    /// <see cref="BuildField"/>/<see cref="BuildLabel"/>) a path does not resolve to a leaf column, or
    /// <paramref name="schema"/>'s report has no storeItemID.
    /// </exception>
    public static TableRow BuildStaticRow(
        DatasetTree schema,
        IReadOnlyList<RepeaterRowCell> cells,
        IReadOnlyList<int> cellWidths,
        Func<int> nextId)
    {
        ArgumentNullException.ThrowIfNull(schema);
        ArgumentNullException.ThrowIfNull(cells);
        ArgumentNullException.ThrowIfNull(cellWidths);
        ArgumentNullException.ThrowIfNull(nextId);

        if (cells.Count == 0)
        {
            throw new ArgumentException("At least one cell is required.", nameof(cells));
        }

        return BuildRowFromCells(schema, cells, cellWidths, nextId, path => path);
    }

    /// <summary>
    /// The one row-from-cell-specs builder <see cref="BuildDetailRepeaterRow"/> and
    /// <see cref="BuildStaticRow"/> share: one <c>w:tc</c> per <see cref="RepeaterRowCell"/> (explicit
    /// <c>w:tcW</c>, <c>gridSpan</c> when spanning, compact single-spaced paragraph, optional
    /// justification), each cell's entries resolved to full dataset paths by
    /// <paramref name="resolveEntryToPath"/> and chained inline in one paragraph — label-shaped paths (per
    /// the active <see cref="LabelConvention"/>) as label controls, the rest as fields, with no separator
    /// between them (the corpus shape).
    /// </summary>
    private static TableRow BuildRowFromCells(
        DatasetTree schema,
        IReadOnlyList<RepeaterRowCell> cells,
        IReadOnlyList<int> cellWidths,
        Func<int> nextId,
        Func<string, string> resolveEntryToPath)
    {
        var convention = LabelConvention.Current;

        var row = new TableRow();
        for (var i = 0; i < cells.Count; i++)
        {
            var cell = cells[i];
            if (cell.Span < 1)
            {
                throw new ArgumentException($"cells[{i}].Span must be >= 1 (got {cell.Span}).", nameof(cells));
            }

            var paragraph = new Paragraph
            {
                ParagraphProperties = new ParagraphProperties
                {
                    SpacingBetweenLines = new SpacingBetweenLines { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto },
                },
            };
            if (cell.Alignment is { } alignment)
            {
                paragraph.ParagraphProperties.Justification =
                    new Justification { Val = CellTextEditor.ParseAlignment(alignment, nameof(cells)) };
            }

            foreach (var entry in cell.Columns)
            {
                var path = resolveEntryToPath(entry);
                var segments = path.Split('/', StringSplitOptions.RemoveEmptyEntries);
                paragraph.AppendChild<OpenXmlElement>(convention.IsLabelPath(segments)
                    ? BuildLabel(schema, path, id: nextId())
                    : BuildField(schema, path, id: nextId()));
            }

            var tcPr = new TableCellProperties(new TableCellWidth
            {
                Type = TableWidthUnitValues.Dxa,
                Width = cellWidths[i].ToString(CultureInfo.InvariantCulture),
            });
            if (cell.Span > 1)
            {
                tcPr.Append(new GridSpan { Val = cell.Span });
            }

            row.AppendChild(new TableCell(tcPr, paragraph));
        }

        return row;
    }

    /// <summary>Splits, naming-convention-checks, and schema-validates a dataset path; returns its segments.</summary>
    private static string[] ValidatedSegments(DatasetTree schema, string datasetPath, bool expectLabel)
    {
        ArgumentNullException.ThrowIfNull(schema);
        if (string.IsNullOrWhiteSpace(datasetPath))
        {
            throw new ArgumentException("Dataset path must not be empty.", nameof(datasetPath));
        }

        var segments = datasetPath.Split('/', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (segments.Length == 0)
        {
            throw new ArgumentException($"Dataset path '{datasetPath}' has no segments.", nameof(datasetPath));
        }

        var convention = LabelConvention.Current;
        var isLabel = convention.IsLabelPath(segments);
        if (expectLabel && !isLabel)
        {
            throw new ArgumentException(
                $"'{datasetPath}' is not label-shaped per the active label convention "
                + $"({convention.Describe()}); bind it as a field, not a label (insert_field).",
                nameof(datasetPath));
        }

        if (!expectLabel && isLabel)
        {
            throw new ArgumentException(
                $"'{datasetPath}' is label-shaped per the active label convention "
                + $"({convention.Describe()}); bind it as a label, not a field (insert_label).",
                nameof(datasetPath));
        }

        ValidateLeafColumn(schema, segments, datasetPath);
        return segments;
    }

    /// <summary>
    /// Walks <paramref name="segments"/> down <paramref name="schema"/>'s data-item tree exactly like
    /// <see cref="LayoutValidator"/>'s own xpath-resolves check, but with richer failure detail: distinct
    /// messages for "segment not found at all" vs. "found, but it's a repeating data item, not a leaf
    /// column" (a field/label must bind to a leaf column).
    /// </summary>
    private static void ValidateLeafColumn(DatasetTree schema, string[] segments, string datasetPath)
    {
        var node = schema.Root;
        for (var i = 0; i < segments.Length; i++)
        {
            var name = segments[i];
            var isLast = i == segments.Length - 1;

            var childItem = node.FindChildItem(name);
            if (childItem is not null)
            {
                if (isLast)
                {
                    throw new ArgumentException(
                        $"Dataset path '{datasetPath}' resolves to a repeating data item ('{name}'), not a leaf "
                        + "column; field/label controls must bind to a leaf column (use a repeater control for "
                        + "data items).",
                        nameof(datasetPath));
                }

                node = childItem;
                continue;
            }

            if (isLast && node.FindChildColumn(name) is not null)
            {
                return;
            }

            throw new ArgumentException(
                $"Dataset path '{datasetPath}' does not resolve against the report schema: segment '{name}' was "
                + $"not found under '{node.Path}'.",
                nameof(datasetPath));
        }
    }

    private static SdtRun Build(
        DatasetTree schema, IReadOnlyList<string> segments, string datasetPath, string? placeholderText, int? id)
    {
        var sdtPr = BuildSdtProperties(schema, segments, id, new SdtContentText(), withPlaceholder: true);
        var text = string.IsNullOrEmpty(placeholderText) ? segments[^1] : placeholderText;

        return new SdtRun(sdtPr, new SdtContentRun(new Run(new Text(text))));
    }

    /// <summary>
    /// The <c>w:sdtPr</c> every bound control this factory builds shares — alias, tag, id, (optionally) the
    /// built-in placeholder reference, and the <c>w:dataBinding</c> — closed by
    /// <paramref name="typeMarker"/>, the element that says WHICH kind of control it is (<c>w:text</c> for a
    /// field/label, <c>w:picture</c> for a picture). <paramref name="withPlaceholder"/> is false only for a
    /// picture, whose real add-in-authored corpus counterpart carries no <c>w:placeholder</c> (see
    /// <see cref="BuildPicture"/>).
    /// </summary>
    private static SdtProperties BuildSdtProperties(
        DatasetTree schema, IReadOnlyList<string> segments, int? id, OpenXmlElement typeMarker, bool withPlaceholder)
    {
        var report = schema.Report;
        if (string.IsNullOrEmpty(report.StoreItemId))
        {
            throw new ArgumentException(
                $"schema.Report.{nameof(ReportIdentity.StoreItemId)} is null/empty; build the schema via "
                + $"{nameof(SchemaProvider)}.{nameof(SchemaProvider.FromLayout)} (not "
                + $"{nameof(SchemaProvider.FromSchemaXml)}) so a real storeItemID is available to bind against.",
                nameof(schema));
        }

        var sdtPr = new SdtProperties(
            new SdtAlias { Val = $"#Nav: /{string.Join('/', segments)}" },
            new Tag { Val = $"#Nav: {report.ReportName}/{report.ReportId}" },
            new SdtId { Val = ResolveId(id) });

        if (withPlaceholder)
        {
            sdtPr.AppendChild(new SdtPlaceholder(new DocPartReference { Val = DefaultPlaceholderDocPart }));
        }

        sdtPr.AppendChild(new DataBinding
        {
            PrefixMappings = $"xmlns:ns0='{report.Namespace}'",
            XPath = BuildIndexedXPath(schema.Root.Name, segments),
            StoreItemId = report.StoreItemId,
        });
        sdtPr.AppendChild(typeMarker);

        return sdtPr;
    }

    /// <summary>
    /// The <c>w:r</c> holding a picture control's inline drawing, element for element as the corpus capture
    /// has it (see <see cref="BuildPicture"/>): a <c>w:noProof</c> run, a <c>wp:inline</c> frame of the given
    /// extent, and a <c>pic:pic</c> whose blip references <paramref name="embedRelationshipId"/> and whose
    /// shape properties repeat the same extent.
    /// </summary>
    private static Run BuildPictureRun(string embedRelationshipId, uint drawingId, long widthEmu, long heightEmu)
    {
        var name = $"Picture {drawingId}";

        var picture = new Pic.Picture(
            new Pic.NonVisualPictureProperties(
                new Pic.NonVisualDrawingProperties { Id = drawingId, Name = name },
                new Pic.NonVisualPictureDrawingProperties(
                    new A.PictureLocks { NoChangeAspect = true, NoChangeArrowheads = true })
                {
                    PreferRelativeResize = false,
                }),
            new Pic.BlipFill(
                new A.Blip { Embed = embedRelationshipId },
                new A.Stretch(new A.FillRectangle())),
            new Pic.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 0, Y = 0 },
                    new A.Extents { Cx = widthEmu, Cy = heightEmu }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle },
                new A.NoFill(),
                new A.Outline(new A.NoFill()))
            {
                BlackWhiteMode = A.BlackWhiteModeValues.Auto,
            });

        var inline = new Dw.Inline(
            new Dw.Extent { Cx = widthEmu, Cy = heightEmu },
            new Dw.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
            new Dw.DocProperties { Id = drawingId, Name = name },
            new Dw.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
            new A.Graphic(new A.GraphicData(picture) { Uri = PictureGraphicDataUri }))
        {
            DistanceFromTop = 0U,
            DistanceFromBottom = 0U,
            DistanceFromLeft = 0U,
            DistanceFromRight = 0U,
        };

        return new Run(new RunProperties(new NoProof()), new Drawing(inline));
    }

    /// <summary>The DrawingML picture namespace a <c>a:graphicData/@uri</c> must name for a picture frame.</summary>
    private const string PictureGraphicDataUri = "http://schemas.openxmlformats.org/drawingml/2006/picture";

    /// <summary>Fully-indexed XPath: root + each path step, each prefixed <c>ns0:</c> and suffixed <c>[1]</c>.</summary>
    private static string BuildIndexedXPath(string rootName, IReadOnlyList<string> segments)
    {
        var steps = new List<string>(segments.Count + 1) { $"ns0:{rootName}[1]" };
        steps.AddRange(segments.Select(s => $"ns0:{s}[1]"));
        return "/" + string.Join('/', steps);
    }

    /// <summary>
    /// Resolves the <c>w:id</c> to stamp onto a built control: <paramref name="id"/> itself when supplied,
    /// otherwise a plain random <see cref="int"/>. Deliberately untracked — earlier
    /// revisions kept a process-lifetime static <c>HashSet&lt;int&gt;</c> of every id ever issued or passed
    /// explicitly, growing unboundedly across a long host session for a uniqueness scope no caller actually
    /// needs: <c>w:id</c> has no OOXML schema-level uniqueness requirement, and the one scope that DOES
    /// matter — not colliding with a specific target document's own ids — was never served by that set
    /// anyway (it only ever compared against other ids this factory itself had produced, in any document,
    /// not the document a caller is about to insert into). <see cref="LayoutEditor"/>'s
    /// <c>GenerateUniqueId</c>/<c>MakeIdGenerator</c> already do the real, doc-scoped check by scanning the
    /// actual target document, and every production call site already supplies an explicit id from one of
    /// those — so this method drops the dead bookkeeping rather than fix its scope.
    /// </summary>
    private static int ResolveId(int? id) => id ?? Random.Shared.Next(int.MinValue, int.MaxValue);

    // ---- BuildRepeaterTable: data-item / column validation ----

    /// <summary>
    /// Walks <paramref name="dataItemPath"/> down <paramref name="schema"/>'s tree, requiring every segment
    /// to be a data item (not a leaf column) and the final node to be a non-system data item. Distinct,
    /// specific messages mirror <see cref="ValidateLeafColumn"/>'s own style but for the reverse
    /// expectation (a repeater table binds a whole repeating item, not one leaf column).
    /// </summary>
    private static DataItem ResolveRepeaterDataItem(DatasetTree schema, string dataItemPath)
    {
        if (string.IsNullOrWhiteSpace(dataItemPath))
        {
            throw new ArgumentException("Data item path must not be empty.", nameof(dataItemPath));
        }

        var segments = dataItemPath.Split('/', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (segments.Length == 0)
        {
            throw new ArgumentException($"Data item path '{dataItemPath}' has no segments.", nameof(dataItemPath));
        }

        var node = schema.Root;
        foreach (var name in segments)
        {
            var child = node.FindChildItem(name);
            if (child is null)
            {
                if (node.FindChildColumn(name) is not null)
                {
                    throw new ArgumentException(
                        $"Data item path '{dataItemPath}' resolves to a leaf column ('{name}'), not a "
                        + "repeating data item; a repeater table must bind to a data item (bind a single "
                        + "leaf column as a field/label instead - insert_field/insert_label).",
                        nameof(dataItemPath));
                }

                throw new ArgumentException(
                    $"Data item path '{dataItemPath}' does not resolve against the report schema: segment "
                    + $"'{name}' was not found under '{node.Path}'.",
                    nameof(dataItemPath));
            }

            node = child;
        }

        if (node.IsSystem)
        {
            throw new ArgumentException(
                $"Data item path '{dataItemPath}' resolves to the system data item ('{node.Name}'); a "
                + "repeater table must bind to a business data item, not BCReportInformation.",
                nameof(dataItemPath));
        }

        return node;
    }

    /// <summary>
    /// Requires <paramref name="column"/> to be a leaf column directly under <paramref name="dataItem"/>
    /// (the label-shaped check itself is left to <see cref="BuildField"/>, which every data cell already
    /// goes through — see <see cref="BuildRepeaterTable"/>).
    /// </summary>
    private static void ValidateColumnOfItem(DataItem dataItem, string dataItemPath, string column)
    {
        if (string.IsNullOrWhiteSpace(column))
        {
            throw new ArgumentException("Column name must not be empty.", nameof(column));
        }

        if (dataItem.FindChildColumn(column) is not null)
        {
            return;
        }

        if (dataItem.FindChildItem(column) is not null)
        {
            throw new ArgumentException(
                $"Column '{column}' resolves to a nested repeating data item under '{dataItemPath}', not a "
                + "leaf column; repeater table columns must be leaf columns of the bound data item.",
                nameof(column));
        }

        throw new ArgumentException(
            $"Column '{column}' is not a leaf column of data item '{dataItemPath}'.",
            nameof(column));
    }

    // ---- BuildRepeaterTable: column widths ----

    private const int DefaultColumnWidthTwips = 2000;

    private static IReadOnlyList<int> ResolveColumnWidths(IReadOnlyList<int>? explicitWidths, int columnCount) =>
        explicitWidths ?? Enumerable.Repeat(DefaultColumnWidthTwips, columnCount).ToList();

    // ---- BuildRepeaterTable: table shell (tblPr/tblGrid/cells) ----

    /// <summary>
    /// An optional <c>w:tblStyle</c> reference plus — for <see cref="TableBorderLook.Grid"/> only — explicit
    /// single-line borders (top/left/bottom/right/inside). <see cref="TableBorderLook.Bc"/> emits no
    /// <c>w:tblBorders</c> at all, exactly like every real corpus lines table: its look comes from the
    /// per-cell header rule <see cref="BuildRepeaterTable"/> applies instead.
    /// </summary>
    private static TableProperties BuildTableProperties(RepeaterTableOptions options, IReadOnlyList<int> widths)
    {
        var tblPr = new TableProperties();
        if (!string.IsNullOrEmpty(options.TableStyle))
        {
            tblPr.TableStyle = new TableStyle { Val = options.TableStyle };
        }

        // Pin the grid the caller asked for; without this Word autofits and recomputes every width from
        // cell content (see FixedTableLayout). Assigned through typed properties throughout this method so
        // each element lands at its schema position regardless of the order they are set in.
        FixedTableLayout.ApplyTo(tblPr, widths);

        if (options.Look != TableBorderLook.Grid)
        {
            return tblPr;
        }

        tblPr.TableBorders = new TableBorders
        {
            TopBorder = new TopBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
            LeftBorder = new LeftBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
            BottomBorder = new BottomBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
            RightBorder = new RightBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
            InsideHorizontalBorder =
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
            InsideVerticalBorder =
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
        };

        return tblPr;
    }

    private static TableCell WrapCell(Paragraph paragraph, int width)
    {
        // Compact, single-spaced cells — the shape real BC layout cells render with (their paragraphs
        // override the document's airy defaults explicitly). Without this a blank document's Word
        // defaults (~8pt after each paragraph, 1.08 line) inflate every repeater row.
        paragraph.ParagraphProperties ??= new ParagraphProperties();
        paragraph.ParagraphProperties.SpacingBetweenLines ??=
            new SpacingBetweenLines { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };
        return WrapCellCore(paragraph, width);
    }

    private static TableCell WrapCellCore(Paragraph paragraph, int width) =>
        new(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = width.ToString(CultureInfo.InvariantCulture) }), paragraph);

    /// <summary>
    /// Builds one header cell for <paramref name="column"/>: a bound label control when
    /// <see cref="RepeaterTableOptions.HeaderFromLabels"/> is true and a label column can be found (see
    /// <see cref="FindLabelColumn"/>), otherwise static humanized text (see <see cref="Humanize"/>).
    /// </summary>
    private static TableCell BuildHeaderCell(
        DatasetTree schema, DataItem dataItem, string column, RepeaterTableOptions options, Func<int> nextId, int width)
    {
        if (options.HeaderFromLabels)
        {
            var labelColumn = FindLabelColumn(schema, dataItem, column);
            if (labelColumn is not null)
            {
                var label = BuildLabel(schema, labelColumn.Path, id: nextId());
                return WrapCell(new Paragraph(label), width);
            }
        }

        return WrapCell(new Paragraph(new Run(new Text(Humanize(column)))), width);
    }

    /// <summary>
    /// Finds <paramref name="column"/>'s label column: first a same-item sibling named
    /// <c>&lt;column&gt;_&lt;suffix&gt;</c> or <c>&lt;column&gt;&lt;suffix&gt;</c> for each suffix in the
    /// active <see cref="LabelConvention.Current"/>'s <see cref="LabelConvention.Suffixes"/> (in order — the
    /// BC-default convention tries <c>_Lbl</c> then <c>Lbl</c>); failing that, any label column anywhere in
    /// the schema whose de-suffixed name (see <see cref="LabelConvention.StripSuffix"/>) equals
    /// <paramref name="column"/> exactly, in schema tree order (first match wins). Returns null when neither
    /// finds anything — the caller falls back to static text.
    /// </summary>
    private static DatasetColumn? FindLabelColumn(DatasetTree schema, DataItem dataItem, string column)
    {
        var convention = LabelConvention.Current;
        foreach (var suffix in convention.Suffixes)
        {
            var underscoreSibling = dataItem.FindChildColumn(column + "_" + suffix);
            if (underscoreSibling is not null)
            {
                return underscoreSibling;
            }

            var suffixOnlySibling = dataItem.FindChildColumn(column + suffix);
            if (suffixOnlySibling is not null)
            {
                return suffixOnlySibling;
            }
        }

        return schema.AllColumns(includeSystem: false)
            .FirstOrDefault(c => c.IsLabel && string.Equals(convention.StripSuffix(c.Name), column, StringComparison.Ordinal));
    }

    private static readonly Regex HumanizeCamelBoundary = new(@"(?<=[a-z0-9])(?=[A-Z])", RegexOptions.Compiled);
    private static readonly Regex HumanizeAcronymBoundary = new(@"(?<=[A-Z])(?=[A-Z][a-z])", RegexOptions.Compiled);

    /// <summary>
    /// Turns a column name into plain humanized text for a static header cell, e.g.
    /// <c>CompanyVATRegNo</c> -&gt; <c>Company VAT Reg No</c>: splits camelCase/acronym-then-word
    /// boundaries and underscores into spaces.
    /// </summary>
    internal static string Humanize(string columnName)
    {
        var split = HumanizeAcronymBoundary.Replace(HumanizeCamelBoundary.Replace(columnName, " "), " ");
        var words = split.Replace('_', ' ').Split(' ', StringSplitOptions.RemoveEmptyEntries);
        return string.Join(' ', words);
    }

    // ---- BuildRepeaterTable: repeatingSection / repeatingSectionItem wrapper ----

    /// <summary>
    /// Wraps <paramref name="dataRow"/> in <c>SdtRow(w15:repeatingSectionItem) &gt; SdtContentRow</c>, then
    /// wraps THAT in <c>SdtRow(w15:repeatingSection, w15:dataBinding) &gt; SdtContentRow</c> bound to
    /// <paramref name="dataItem"/> with an UNINDEXED final XPath step (see <see cref="BuildRepeaterXPath"/>)
    /// — the shape verified against the real corpus (see <see cref="BuildRepeaterTable"/>'s remarks).
    /// </summary>
    private static SdtRow BuildRepeaterRow(DatasetTree schema, DataItem dataItem, TableRow dataRow, Func<int> nextId)
    {
        var report = schema.Report;
        var segments = dataItem.Path.Split('/', StringSplitOptions.RemoveEmptyEntries);
        var xpath = BuildRepeaterXPath(schema.Root.Name, segments);
        var alias = $"#Nav: /{string.Join('/', segments)}";
        var tag = $"#Nav: {report.ReportName}/{report.ReportId}";

        var itemSdtPr = new SdtProperties(
            new SdtId { Val = nextId() },
            new SdtPlaceholder(new DocPartReference { Val = DefaultPlaceholderDocPart }),
            new Office2013Word.SdtRepeatedSectionItem());
        var itemRow = new SdtRow(itemSdtPr, new SdtContentRow(dataRow));

        var repeaterSdtPr = new SdtProperties(
            new SdtAlias { Val = alias },
            new Tag { Val = tag },
            new SdtId { Val = nextId() },
            new Office2013Word.DataBinding
            {
                PrefixMappings = $"xmlns:ns0='{report.Namespace}'",
                XPath = xpath,
                StoreItemId = report.StoreItemId,
            },
            new Office2013Word.SdtRepeatedSection());

        return new SdtRow(repeaterSdtPr, new SdtContentRow(itemRow));
    }

    /// <summary>
    /// Same shape as <see cref="BuildIndexedXPath"/> except the FINAL step is left unindexed (no
    /// <c>[1]</c>) — the data-item xpath a <c>w15:dataBinding</c> repeater binding requires, e.g.
    /// <c>/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line</c>.
    /// </summary>
    private static string BuildRepeaterXPath(string rootName, string[] segments)
    {
        var steps = new List<string>(segments.Length + 1) { $"ns0:{rootName}[1]" };
        for (var i = 0; i < segments.Length; i++)
        {
            var isLast = i == segments.Length - 1;
            steps.Add(isLast ? $"ns0:{segments[i]}" : $"ns0:{segments[i]}[1]");
        }

        return "/" + string.Join('/', steps);
    }
}
