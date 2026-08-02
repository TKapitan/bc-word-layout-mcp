namespace BcWordLayout.Domain;

/// <summary>
/// Discriminates WHAT KIND of lookup a <see cref="NotFoundException"/> reports, so a consumer (today, only
/// <c>BcWordLayout.McpHost.Tools.ToolGuards.Guard</c>) can pick an agent-actionable hint by SWITCHING ON
/// THIS VALUE instead of pattern-matching the exception's free-text <see cref="Exception.Message"/>. Every
/// member here mirrors one lookup family that <see cref="LocationResolver"/>/<see cref="LayoutEditor"/>/
/// <see cref="TableGridNavigator"/>/<see cref="TableStructureEditor"/> can fail to resolve; the mapping is
/// 1:1 with the hint branches the message-sniffing <c>NotFoundHint(string)</c> method used to cover.
/// <para>
/// WHY A DEDICATED EXCEPTION TYPE EXISTS. The host used to map <see cref="InvalidOperationException"/>
/// wholesale to <c>not_found</c>. But that is also the BCL's default failure type and what a stray LINQ
/// <c>First()</c>/<c>Single()</c> throws, so a genuine INTERNAL bug was reported to the agent as "the
/// referenced control id does not exist" — sending it into a futile <c>get_layout_info</c> retry loop
/// against a layout that was never the problem. Only this type maps to <c>not_found</c> now; a raw
/// <see cref="InvalidOperationException"/> falls through to <c>internal_error</c>, where it belongs.
/// </para>
/// </summary>
public enum NotFoundTarget
{
    /// <summary>An <c>AfterControl</c>/<c>remove_control</c> lookup by <c>w:id</c> found no matching control
    /// anywhere the caller searched (a specific part, or the whole document).</summary>
    Control,

    /// <summary>A <see cref="Location.PartName"/> was supplied but did not match any header/footer part
    /// actually present in the layout (the layout DOES have header/footer parts — just not one with this name).</summary>
    NamedHeaderFooterPart,

    /// <summary>A location targeted <see cref="LayoutPart.Header"/>/<see cref="LayoutPart.Footer"/> but the
    /// layout has NO header/footer parts at all to resolve against (a blank/body-only layout).</summary>
    HeaderFooterParts,

    /// <summary>A <c>tableIndex</c>/<c>row</c>/<c>col</c> (or grid-column) address is out of range for the
    /// table it names, or resolves to a control wrapper with no inner row/cell to operate on.</summary>
    TableCoordinate,

    /// <summary>An <c>atText</c> location's <c>searchText</c> substring does not appear in any run of text
    /// in the part searched.</summary>
    SearchText,

    /// <summary>An <c>AfterControl</c> location found the control, but its position in the tree cannot
    /// safely host an inserted sibling (e.g. a row-level repeater control, whose parent is a
    /// <c>w:tbl</c> — a table cannot take a paragraph or block sdt as a direct child) — a deliberate
    /// "refuse rather than guess" rejection, not a document defect.</summary>
    AfterControlPosition,

    /// <summary>
    /// Fallback for a lookup failure that does not cleanly fit any of the buckets above. Not thrown by any
    /// site in this codebase today (every throw site picks a specific value); kept so a future Domain
    /// lookup can still participate in the typed <c>not_found</c> contract on day one instead of reverting
    /// to message-sniffing or a raw <see cref="InvalidOperationException"/> while a real bucket is designed.
    /// </summary>
    General,
}

/// <summary>
/// Thrown when a well-formed request (already past <see cref="ArgumentException"/>-shaped validation) names
/// something that does not exist in THIS document: an unknown control <c>w:id</c>, an out-of-range table/
/// row/column index, a header/footer part name that is not present, a layout with no header/footer parts at
/// all, or a <c>searchText</c> substring that is not actually present. <see cref="TargetKind"/> says WHICH of
/// those this particular failure was, so a caller can react without parsing <see cref="Exception.Message"/>.
/// </summary>
/// <remarks>
/// <para>
/// WHY THIS TYPE EXISTS. Before this type, every one of the lookup
/// failures above was a plain <see cref="InvalidOperationException"/> — but that is also the BCL's default
/// failure type, thrown by a stray LINQ <c>.First()</c>/<c>.Single()</c>, a null-coalescing assertion
/// (<c>?? throw new InvalidOperationException(...)</c>) guarding an internal invariant, or any other
/// "this should be impossible" assertion anywhere in <c>src/</c>. <c>ToolGuards.Guard</c> mapped EVERY
/// <see cref="InvalidOperationException"/> to the <c>not_found</c> error code with a "the referenced control/
/// location does not exist — call <c>get_layout_info</c> and retry" hint. For a genuine lookup failure that
/// hint is exactly right; for an internal-bug-shaped throw (e.g. "enclosing paragraph has no parent",
/// "no w:id (unexpected)") it actively misleads the calling agent into a futile retry loop instead of
/// reporting the failure as the tool bug it actually is. The distinction held together only by throw-site
/// discipline — a hand-rolled convention with no compiler or test enforcing it, one every future Domain
/// method silently joins whether or not its author intends to. A dedicated type makes the distinction
/// STRUCTURAL: only a throw site that deliberately constructs a <see cref="NotFoundException"/> ever
/// produces <c>not_found</c>; every other <see cref="InvalidOperationException"/> (the BCL default) falls
/// through <c>Guard</c>'s generic <c>catch (Exception)</c> to <c>internal_error</c> instead, exactly where an
/// unreviewed assertion belongs.
/// </para>
/// <para>
/// WHY <see cref="TargetKind"/> INSTEAD OF JUST THIS TYPE. <c>ToolGuards.Guard</c> tailors its <c>not_found</c>
/// hint to WHICH kind of thing was not found (a bad control id gets "ids are per-document, not sequential or
/// guessable"; an out-of-range table coordinate gets "call get_layout_info... before retrying tableCell";
/// etc.) — the pre-B11 code picked that hint by pattern-matching substrings of <c>ex.Message</c>
/// (<c>message.Contains("No control with id", ...)</c>), which silently degrades to the generic fallback
/// hint the moment a Domain throw site's wording is reworded for clarity, with no compile or test signal
/// that the specific hint stopped firing. <see cref="TargetKind"/> replaces that string coupling with an
/// enum every throw site sets explicitly, so no rewording of a Domain message can silently change which
/// hint fires.
/// </para>
/// </remarks>
public sealed class NotFoundException : Exception
{
    /// <summary>Constructs a lookup failure carrying <paramref name="targetKind"/> for hint selection.</summary>
    public NotFoundException(string message, NotFoundTarget targetKind)
        : base(message)
    {
        TargetKind = targetKind;
    }

    /// <summary>Which family of lookup failed — see <see cref="NotFoundTarget"/>'s own members.</summary>
    public NotFoundTarget TargetKind { get; }
}
