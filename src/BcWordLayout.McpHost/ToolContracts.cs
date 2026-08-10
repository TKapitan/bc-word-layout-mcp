namespace BcWordLayout.McpHost;

/// <summary>
/// Structured error carried in every tool failure (design doc §4). Chosen error model for this server:
/// tools never throw across the MCP boundary; instead they return a <see cref="ToolResponse"/> envelope
/// with <c>Ok = false</c> and a populated <see cref="ToolError"/> (<c>Code</c>, <c>Message</c>, <c>Hint</c>).
/// This keeps failures deterministic and machine-actionable for the calling agent.
/// </summary>
/// <remarks>
/// <para>
/// ENVELOPE BOUNDARY. "Tools never throw" describes every tool BODY that actually
/// runs — it is a guarantee about <c>Tools.ToolGuards.Guard</c>/<c>GuardMutate</c>'s exception mapping,
/// which only executes once the MCP C# SDK has already bound the incoming JSON-RPC arguments onto the
/// tool method's typed parameters. That binding step happens INSIDE the SDK, before any of this
/// assembly's code runs, and is NOT covered: a caller passing a wrong-typed argument (e.g. a string where
/// an <c>int</c> parameter is expected), omitting a required parameter, or sending malformed params never
/// reaches a <see cref="ToolResponse"/> at all. Verified empirically against `ModelContextProtocol`
/// 1.4.1: such a call returns an MCP <c>CallToolResult</c> with <c>IsError = true</c> whose content is a
/// generic SDK-generated message (e.g. <c>"An error occurred invoking 'insert_field'."</c>) with no
/// <c>code</c>/<c>hint</c> — never this type's JSON shape. This is distinct from calling an entirely
/// unregistered tool NAME, which the SDK instead surfaces as a client-visible protocol-level error.
/// </para>
/// <para>
/// No proportionate code fix exists for this SDK version: <c>McpServerOptions.Filters.Request.CallToolFilters</c>
/// only wraps calls to a tool name that is NOT in the server's <c>ToolCollection</c> (i.e. never fires for
/// any of this server's registered tools, which are always found there first); the message-level
/// <c>Filters.Message.IncomingFilters</c>/<c>OutgoingFilters</c> only ever see the already-formed,
/// non-exceptional <c>isError:true</c> result — nothing throws through them either, so there is no
/// exception to catch-and-rewrap, only a generic response body that would have to be pattern-matched by
/// its (undocumented, SDK-internal, version-fragile) message text to be reshaped into this envelope. That
/// trade — hand-parsing an internal SDK string to recover a `code`/`hint` we have no real information for
/// — was judged worse than documenting the boundary plainly here and in README.md's "Response envelope"
/// section. Revisit if a future SDK version exposes a binding-failure hook.
/// </para>
/// </remarks>
/// <remarks>
/// Structural guarantee: <see cref="Hint"/> is a non-nullable <see cref="string"/>, and
/// <see cref="ToolResponse.Failure"/> - the only way to build a failing <see cref="ToolResponse"/> - requires
/// a <c>hint</c> argument with no default value. A hintless failure therefore cannot compile; it is not
/// merely a convention every call site happens to follow. Every call site (see <c>Tools.ToolGuards</c>'s
/// <c>Guard</c>/<c>GuardMutate</c> exception-to-code mappings and its direct <c>Failure</c> calls) supplies a
/// hint that is agent-actionable: it tells the calling agent what to DO next - which argument to fix and its
/// valid values, or which inspection tool (<c>get_layout_info</c>/<c>list_dataset_fields</c>) to call - not
/// just a restatement that something went wrong.
/// </remarks>
public sealed record ToolError(string Code, string Message, string Hint);

/// <summary>Uniform envelope returned by every tool.</summary>
public sealed class ToolResponse
{
    public bool Ok { get; init; }
    public object? Data { get; init; }
    public ToolError? Error { get; init; }

    public static ToolResponse Success(object data) => new() { Ok = true, Data = data };

    /// <summary>
    /// Builds a failing envelope. <paramref name="hint"/> is required (no default value, not nullable) -
    /// see <see cref="ToolError"/>'s own remarks for why that is a deliberate structural guarantee
    /// rather than an omission every caller must remember on its own.
    /// </summary>
    public static ToolResponse Failure(string code, string message, string hint) =>
        new() { Ok = false, Error = new ToolError(code, message, hint) };
}

// ---- Output DTOs (stable JSON shape for agents) ----

public sealed record ReportInfoDto(string ReportName, string ReportId, string Namespace, string? StoreItemId);

public sealed record ValidationSummaryDto(string Level, bool Passed, int ErrorCount, int WarningCount);

/// <summary>
/// One control in the flat inventory. <see cref="Level"/> is the structural <c>w:sdt</c> kind
/// (<c>run</c>/<c>block</c>/<c>cell</c>/<c>row</c>/<c>runRuby</c>) — a <c>cell</c>/<c>row</c> control wraps a
/// table cell/row that defines the grid, which is what a caller must know before removing it.
/// <see cref="TableIndex"/>/<see cref="RowIndex"/>/<see cref="ColIndex"/> locate the control in a table
/// (all null when it is not in one; <see cref="ColIndex"/> alone is null for a row-level control) — the
/// table index is the SAME index the insert tools' <c>tableCell</c> addressing uses.
/// </summary>
public sealed record ControlDto(
    string Kind,
    string? Alias,
    string? Tag,
    string? XPath,
    string? StoreItemId,
    string Part,
    int? SdtId,
    bool UsesW15Binding,
    string? ParentRepeaterXPath,
    string Level,
    int? TableIndex,
    int? RowIndex,
    int? ColIndex);

public sealed record ControlSummaryDto(int Field, int Label, int Repeater, int Picture, int Unbound, int Total);

/// <summary>One cell of a <see cref="TableDto"/> — its column index, the control that owns it (if any), and its visible text.</summary>
public sealed record TableCellDto(
    int ColIndex,
    bool IsControlCell,
    int? ControlId,
    string? ControlKind,
    string? Alias,
    string? XPath,
    string Text,
    IReadOnlyList<int> InnerControlIds);

/// <summary>One row of a <see cref="TableDto"/>. <see cref="IsControlRow"/> marks a row wrapped by a row-level control (e.g. a repeater).</summary>
public sealed record TableRowDto(
    int RowIndex,
    bool IsControlRow,
    int? ControlId,
    IReadOnlyList<TableCellDto> Cells);

/// <summary>
/// One table, described structurally. <see cref="TableIndex"/> is the 0-based, per-part, document-order
/// index the insert tools' <c>tableCell</c> addressing uses, so a table here can be targeted by it directly.
/// </summary>
public sealed record TableDto(
    string Part,
    int TableIndex,
    int RowCount,
    int ColumnCount,
    IReadOnlyList<int> GridColumnWidths,
    IReadOnlyList<TableRowDto> Rows);

/// <summary>
/// One content part of the layout, with its RENDERING facts: <see cref="Kind"/> is
/// <c>document</c>/<c>header</c>/<c>footer</c>; <see cref="Role"/> is which pages the section reference
/// assigns it (<c>default</c> = everyday pages, <c>first</c> = the first page — only rendered when the
/// layout sets <c>w:titlePg</c>, see <see cref="LayoutInfoDto.HasTitlePage"/> — <c>even</c> = even pages;
/// null for the document part or a part no section references). <see cref="IsDefaultTarget"/> marks the
/// part a <c>layoutPart='header'/'footer'</c> location WITHOUT an explicit <c>partName</c> resolves to —
/// computed through the same selection logic the edit tools use, so it cannot disagree with where an edit
/// lands.
/// </summary>
public sealed record PartInfoDto(string Name, string Kind, string? Role, bool IsDefaultTarget);

/// <summary>
/// <c>get_layout_info</c>'s payload. <see cref="Parts"/> is the flat part-name list (kept stable);
/// <see cref="PartDetails"/> is the same list with each part's kind/role/default-target facts — what a
/// caller must consult before addressing a header/footer edit, because the package-order FIRST part is
/// frequently NOT the everyday default one (see <see cref="PartInfoDto"/>). <see cref="HasTitlePage"/> is
/// true when the first section renders a DIFFERENT FIRST PAGE (<c>w:titlePg</c>): page 1 then shows the
/// <c>first</c>-role header/footer (or none), not the <c>default</c> one — the reason a correctly
/// inserted default-header field can be invisible on a one-page render (GitHub issue #5).
/// </summary>
public sealed record LayoutInfoDto(
    ReportInfoDto Report,
    ValidationSummaryDto Validation,
    ControlSummaryDto ControlSummary,
    IReadOnlyList<string> Parts,
    IReadOnlyList<ControlDto> Controls,
    IReadOnlyList<TableDto> Tables,
    IReadOnlyList<PartInfoDto> PartDetails,
    bool HasTitlePage);

public sealed record ColumnDto(string Name, string Path, bool IsLabel, bool? Bound);

public sealed record DataItemDto(
    string Name,
    string Path,
    bool IsSystem,
    IReadOnlyList<ColumnDto> Columns,
    IReadOnlyList<DataItemDto> Children);

public sealed record DatasetFieldsDto(string SourceType, ReportInfoDto Report, DataItemDto Root);

public sealed record FindingDto(string Check, string Severity, string Message, string? Location);

public sealed record ValidationResultDto(
    string Level,
    bool Passed,
    int ErrorCount,
    int WarningCount,
    IReadOnlyList<FindingDto> Findings);

public sealed record PreviewStatsDto(int FieldsFilled, int RepeatersExpanded, int RowsGenerated, int Unresolved, int PicturesFilled);

public sealed record MergeWarningDto(string Kind, string Message, string? Location);

public sealed record PreviewResultDto(
    string MergedDocxPath,
    string? PdfPath,
    string ConverterUsed,
    bool ConverterAvailable,
    bool ConversionOk,
    string? ConversionError,
    PreviewStatsDto Stats,
    IReadOnlyList<MergeWarningDto> Warnings,
    ValidationSummaryDto QuickValidation,
    string Disclaimer);

/// <summary>
/// Per-page metadata for one image returned by <c>render_preview_pages</c>. The pixels themselves are NOT
/// in this JSON — each page travels as its own MCP image content block after the JSON text block; this
/// record only tells the caller what to expect there (<see cref="PngByteCount"/> is the encoded PNG size,
/// pre-base64). <see cref="Path"/> is the on-disk PNG when the call supplied an <c>outputDir</c>, and
/// <c>null</c> otherwise — the one piece of this response a human (rather than the calling agent) can open.
/// </summary>
public sealed record PreviewPageDto(
    int PageNumber, int WidthPx, int HeightPx, int PngByteCount, string? Path = null);

/// <summary>
/// JSON half of a successful <c>render_preview_pages</c> response. Unlike every other tool this one's
/// envelope shares the MCP result with non-JSON content (the image blocks), so <see cref="Truncated"/> +
/// <see cref="PageCount"/> exist to keep an agent from mistaking "the first N pages" for "the whole
/// document" — it can page onward via <c>firstPage</c>.
/// </summary>
public sealed record PreviewPagesResultDto(
    string PdfPath,
    int PageCount,
    int FirstPage,
    int PagesRendered,
    int EffectiveDpi,
    bool Truncated,
    IReadOnlyList<PreviewPageDto> Pages);

/// <summary>
/// Result of a mutating edit tool (<c>insert_field</c>, <c>insert_label</c>, <c>insert_picture</c>,
/// <c>remove_control</c>, <c>insert_repeater_table</c>). <see cref="ColumnCount"/>,
/// <see cref="TableIndex"/> and <see cref="DataRowIndex"/> are populated for <c>insert_repeater_table</c>
/// only; null for every other operation. <see cref="QuickValidation"/> is always populated on a successful
/// response: <c>Tools.ToolGuards.GuardEdit</c> is the one and only place this record is constructed, and it
/// always computes a fresh post-edit <c>LayoutValidator.Quick</c> summary before building it - so every
/// mutating tool that routes through <c>GuardEdit</c> is structurally incapable of reporting success
/// without one (a guarantee, not a per-tool convention - see <c>GuardEdit</c>'s own doc-comment).
/// </summary>
public sealed record EditResultDto(
    string Operation,
    int ControlId,
    string? Alias,
    string? XPath,
    string? Kind,
    int? ColumnCount,
    string Part,
    string Summary,
    ValidationSummaryDto QuickValidation,
    int? TableIndex = null,
    int? DataRowIndex = null);

/// <summary>
/// Result of a plain-text cell edit tool (<c>set_cell_text</c>, <c>clear_cell_text</c>). The cell is
/// addressed by its (table, row, column) coordinates rather than a control id (a plain-text cell has no
/// control). <see cref="NewText"/> is the text after the edit (empty for a clear). <see cref="QuickValidation"/>
/// is always populated on a successful response — <c>Tools.ToolGuards.GuardCellEdit</c> is the one and only place
/// this record is constructed, and it always computes a fresh post-edit <c>LayoutValidator.Quick</c> summary
/// before building it (the same structural guarantee <see cref="EditResultDto"/> has via <c>GuardEdit</c>).
/// </summary>
public sealed record CellEditResultDto(
    string Operation,
    string Part,
    int TableIndex,
    int Row,
    int Col,
    string PreviousText,
    string NewText,
    string Summary,
    ValidationSummaryDto QuickValidation);

/// <summary>
/// Result of a table-STRUCTURE edit tool (<c>set_column_widths</c>, <c>insert_column</c>,
/// <c>remove_column</c>). The table is addressed by its 0-based document-order index (the same one
/// <c>get_layout_info</c> reports); <see cref="ColumnIndex"/> is the GRID column the edit added/removed
/// (null for a whole-table width set). <see cref="ColumnCountBefore"/>/<see cref="ColumnCountAfter"/> are
/// the grid column counts around the edit. <see cref="QuickValidation"/> is always populated on a
/// successful response — <c>Tools.ToolGuards.GuardTableEdit</c> is the one and only place this record is built,
/// with the same structural guarantee <see cref="EditResultDto"/>/<see cref="CellEditResultDto"/> have.
/// </summary>
public sealed record TableEditResultDto(
    string Operation,
    string Part,
    int TableIndex,
    int? ColumnIndex,
    int RowsAffected,
    int ColumnCountBefore,
    int ColumnCountAfter,
    string Summary,
    ValidationSummaryDto QuickValidation);

/// <summary>
/// Result of a SUCCESSFUL <c>create_layout</c> call: what was written, whether a template's own BC part was
/// replaced, and a <see cref="QuickValidation"/> summary of the built layout (computed while the package was
/// still open, not re-derived from anything the caller does afterward). A <c>templatePath</c> that already
/// carried its own BC dataset part AND bound content controls that would go stale once that part is replaced
/// never reaches this DTO at all — the call fails outright with error code <c>template_not_unbound</c>
/// instead (see <c>Tools.ToolGuards.Guard</c>'s <c>TemplateNotUnboundException</c> catch and
/// <c>LifecycleTools.CreateLayout</c>'s own <c>[Description]</c>). <c>Ok=true</c> is therefore not quite
/// "guaranteed fully clean" — <see cref="QuickValidation"/> can still carry a WARNING-level finding (e.g. an
/// <c>attached-template</c> relationship the source layout/template happened to carry) — but it is guaranteed
/// free of the specific stale-binding damage this refusal exists to catch.
/// </summary>
public sealed record CreateResultDto(
    string OutputPath,
    string ReportName,
    string ReportId,
    string Namespace,
    string StoreItemId,
    bool UsedTemplate,
    bool ReplacedExistingBcPart,
    ValidationSummaryDto QuickValidation);

/// <summary>An existing control binding that no longer resolves against the NEW schema — left in place, reported.</summary>
/// <remarks><c>SdtId</c> is the control's <c>w:id</c>, ready to pass straight to <c>remove_control</c> (which
/// requires it) without a second <c>get_layout_info</c> lookup.</remarks>
public sealed record OrphanedBindingDto(string? Alias, string XPath, string Part, int? SdtId);

/// <summary>
/// Result of <c>refresh_xml_part</c>: old vs new report identity, whether the dataset namespace itself
/// changed (triggering a <c>w:prefixMappings</c>/<c>w:tag</c> remap), how many existing bindings still
/// resolve against the new schema (<see cref="RemappedCount"/>), which no longer do
/// (<see cref="OrphanedBindings"/> — renamed/deleted fields, left in place for the caller to act on), which
/// new-schema fields have no control bound to them yet (<see cref="NewUnboundFields"/>), and a post-refresh
/// <see cref="QuickValidation"/> summary. A non-zero <see cref="ValidationSummaryDto.ErrorCount"/> here is
/// EXPECTED when there are orphaned bindings (they surface there too, as <c>xpath-resolves</c> findings) —
/// that is the orphan report corroborated by an independent check, not a failed refresh; <c>Ok=true</c>
/// stands regardless.
/// </summary>
public sealed record RefreshResultDto(
    string OldReportName,
    string OldReportId,
    string OldNamespace,
    string NewReportName,
    string NewReportId,
    string NewNamespace,
    string? StoreItemId,
    bool NamespaceChanged,
    int RemappedCount,
    IReadOnlyList<OrphanedBindingDto> OrphanedBindings,
    IReadOnlyList<string> NewUnboundFields,
    ValidationSummaryDto QuickValidation);
