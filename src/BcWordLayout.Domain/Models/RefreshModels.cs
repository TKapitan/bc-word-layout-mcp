namespace BcWordLayout.Domain.Models;

/// <summary>
/// One existing control binding that no longer resolves against the NEW schema after a
/// <see cref="BcWordLayout.Domain.LayoutRefresher.Refresh"/> — its dataset column/data item was renamed or
/// removed. The control itself is left in place (never auto-removed or auto-rebound by
/// <see cref="BcWordLayout.Domain.LayoutRefresher"/>); the caller decides what to do about it (rebind via
/// <c>remove_control</c> + <c>insert_field</c>/<c>insert_label</c>, or fix the AL report dataset instead).
/// </summary>
public sealed class OrphanedBinding
{
    /// <summary>The control's <c>w:alias</c>, when present (e.g. <c>#Nav: /Header/CustomerAddress1</c>).</summary>
    public string? Alias { get; init; }

    /// <summary>The control's binding XPath (raw, indexed, prefixed) that no longer resolves against the new schema.</summary>
    public required string XPath { get; init; }

    /// <summary>Which OOXML part the control lives in (e.g. <c>document.xml</c>, <c>header1.xml</c>).</summary>
    public required string Part { get; init; }

    /// <summary>
    /// The orphaned control's <c>w:id</c> (<see cref="BcWordLayout.Domain.Models.LayoutControl.SdtId"/>), so the
    /// caller can act on it directly — e.g. hand it straight to <c>remove_control</c> (which requires this id),
    /// then rebind with <c>insert_field</c>/<c>insert_label</c> — without a second <c>get_layout_info</c> round
    /// trip and an ambiguous alias/XPath match (the same alias/XPath can legitimately be bound by more than one
    /// control). Null only for a control that carries no <c>w:id</c> at all (not expected for a BC-authored control).
    /// </summary>
    public int? SdtId { get; init; }
}

/// <summary>
/// The result of <see cref="BcWordLayout.Domain.LayoutRefresher.Refresh"/>: old vs new report identity,
/// whether the dataset namespace itself changed (triggering a <c>w:prefixMappings</c>/<c>w:tag</c> remap),
/// how many existing bindings still resolve against the new schema ("remapped"), which no longer do
/// ("orphaned"), which fields in the new schema have no control bound to them yet, and a post-refresh quick
/// validation summary.
/// </summary>
public sealed class RefreshResult
{
    public required string OldReportName { get; init; }
    public required string OldReportId { get; init; }
    public required string OldNamespace { get; init; }

    public required string NewReportName { get; init; }
    public required string NewReportId { get; init; }
    public required string NewNamespace { get; init; }

    /// <summary>
    /// The BC custom XML part's <c>ds:itemID</c> — UNCHANGED by refresh (the whole point of refreshing in
    /// place: every existing binding's <c>w:storeItemID</c> keeps pointing at the same part after its
    /// content is swapped for the new schema). Null only when the layout's BC part itself had no item ID to
    /// begin with (malformed input; not expected for a real BC-produced layout).
    /// </summary>
    public string? StoreItemId { get; init; }

    /// <summary>
    /// True when <see cref="NewNamespace"/> differs from <see cref="OldNamespace"/> (the new schema source
    /// carries a different report name/id) — in that case every binding's <c>w:prefixMappings</c> URI and
    /// every BC-authored control's <c>w:tag</c> were rewritten to the new identity. The XPath element-name
    /// steps themselves are NEVER rewritten either way — that is the "remap where element names match": a
    /// binding survives the refresh as long as the same chain of element names still exists in the new
    /// schema, regardless of whether the report's own name/id changed.
    /// </summary>
    public required bool NamespaceChanged { get; init; }

    /// <summary>
    /// Number of existing control bindings (any kind with a non-null XPath — field, label, repeater,
    /// picture) whose XPath still resolves against the NEW schema by element name. Every bound control is
    /// counted in exactly one of <see cref="RemappedCount"/> or <see cref="OrphanedBindings"/>.
    /// </summary>
    public required int RemappedCount { get; init; }

    /// <summary>
    /// Existing control bindings whose XPath no longer resolves against the NEW schema — the column/data
    /// item they pointed at was renamed or removed in the new schema. Left in place, never auto-deleted or
    /// auto-rebound.
    /// </summary>
    public required IReadOnlyList<OrphanedBinding> OrphanedBindings { get; init; }

    /// <summary>
    /// Dataset paths (<see cref="DatasetColumn.Path"/>) of non-label leaf columns that are GENUINELY NEW as
    /// of this refresh: present in the NEW schema, ABSENT from the OLD schema (an old-vs-new diff, computed
    /// from the OLD schema BEFORE the BC part's content is overwritten), and not yet bound by any existing
    /// control. Label columns (<see cref="DatasetColumn.IsLabel"/>) are always SKIPPED here (a deliberate,
    /// documented choice, not an oversight): most labels are only ever consumed as repeater header cells
    /// found by <see cref="BcWordLayout.Domain.SdtFactory.BuildRepeaterTable"/>'s own label lookup rather
    /// than bound 1:1 by name, so treating every unused one as noteworthy would swamp genuinely actionable
    /// data-field gaps with noise.
    /// <para>
    /// Because this IS a diff against the OLD schema, a field that already existed (bound or not) before the
    /// refresh is NEVER reported here again, even if it is still unbound afterward — a same-schema
    /// (idempotent) refresh therefore always yields an EMPTY list. A field that was already unbound and
    /// still is remains fully discoverable via <c>list_dataset_fields</c>'s own bound/unbound flag; this
    /// list's whole value is the DELTA the AL report just introduced, not a standing inventory of every gap.
    /// </para>
    /// </summary>
    public required IReadOnlyList<string> NewUnboundFields { get; init; }

    /// <summary>
    /// A fresh <see cref="BcWordLayout.Domain.LayoutValidator.Quick"/> summary of the layout AFTER the
    /// refresh (computed in-memory, before the caller persists anything). Every entry in
    /// <see cref="OrphanedBindings"/> also surfaces here as an <c>xpath-resolves</c> finding — that overlap
    /// is EXPECTED (the same fact, reported by an independent code path) and must not be misread as the
    /// refresh having failed; a non-zero <see cref="BcWordLayout.Domain.Models.ValidationResult.ErrorCount"/>
    /// here is the normal, correct outcome of a refresh that produced orphans.
    /// </summary>
    public required ValidationResult QuickValidation { get; init; }
}
