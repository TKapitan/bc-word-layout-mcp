namespace BcWordLayout.Domain.Models;

/// <summary>Classification of a Word content control (<c>w:sdt</c>) found in a BC layout.</summary>
public enum ControlKind
{
    /// <summary>Plain field control bound to a non-label dataset column.</summary>
    Field,

    /// <summary>Field control bound to a label column (by default a name following the <c>*Lbl</c>
    /// convention; see <c>BcWordLayout.Domain.LabelConvention</c> for the full, configurable rule).</summary>
    Label,

    /// <summary>Repeating-section control (<c>w15:repeatingSection</c> + <c>w15:dataBinding</c>).</summary>
    Repeater,

    /// <summary>Picture control (<c>w:picture</c> marker alongside a <c>w:dataBinding</c>).</summary>
    Picture,

    /// <summary>
    /// Control with no dataBinding: page-number fields (<c>w:docPartObj</c>), locked
    /// structural <c>w:group</c> wrappers, and other unbound content controls. Not an error.
    /// </summary>
    Unbound,
}

/// <summary>
/// One entry in a layout's control inventory. XPath is preserved raw (as found); nested repeater
/// paths are NOT re-anchored here — that is the merge engine's job.
/// </summary>
public sealed class LayoutControl
{
    public required ControlKind Kind { get; init; }

    /// <summary>The <c>w:alias</c> value, e.g. <c>#Nav: /Header/CustomerAddress1</c>.</summary>
    public string? Alias { get; init; }

    /// <summary>The <c>w:tag</c> value, e.g. <c>#Nav: Standard_Sales_Invoice/1306</c>.</summary>
    public string? Tag { get; init; }

    /// <summary>The binding XPath (raw, indexed, prefixed), or null for unbound controls.</summary>
    public string? XPath { get; init; }

    /// <summary>The <c>w:storeItemID</c> of the binding, or null for unbound controls.</summary>
    public string? StoreItemId { get; init; }

    /// <summary>
    /// The BC dataset namespace URI the binding's <c>w:prefixMappings</c> declares, or null when the control
    /// is unbound or its prefixMappings names no BC namespace. A value here that differs from the layout's
    /// own BC part namespace means the binding is orphaned — see <c>LayoutValidator</c>'s
    /// <c>binding-namespace</c> check.
    /// </summary>
    public string? BindingNamespace { get; init; }

    /// <summary>Which OOXML part the control lives in (e.g. <c>document.xml</c>, <c>header1.xml</c>).</summary>
    public required string Part { get; init; }

    /// <summary>The <c>w:id</c> of the SDT, if present.</summary>
    public int? SdtId { get; init; }

    /// <summary>True when the binding uses <c>w15:dataBinding</c> (repeaters) rather than <c>w:dataBinding</c>.</summary>
    public bool UsesW15Binding { get; init; }

    /// <summary>
    /// The enclosing repeater control when this control sits inside a repeating section, so callers
    /// can see nesting. Null for top-level controls.
    /// </summary>
    public LayoutControl? ParentRepeater { get; init; }

    /// <summary>
    /// The structural level of the underlying <c>w:sdt</c> (run/block/cell/row/runRuby) — tells a caller
    /// what a whole-control removal would take with it. A <see cref="SdtLevel.Cell"/>/<see cref="SdtLevel.Row"/>
    /// control wraps a table cell/row that defines the grid; see <see cref="SdtLevel"/>.
    /// </summary>
    public SdtLevel Level { get; init; } = SdtLevel.Unknown;

    /// <summary>
    /// 0-based, per-part, document-order index of the table this control sits in (the SAME index the insert
    /// tools' <c>tableCell</c> addressing uses); null when the control is not inside any table.
    /// </summary>
    public int? TableIndex { get; init; }

    /// <summary>0-based row index within <see cref="TableIndex"/>; null when not inside a table.</summary>
    public int? RowIndex { get; init; }

    /// <summary>
    /// 0-based column index within the row; null when not inside a table cell (e.g. a row-level control,
    /// which spans the whole row rather than sitting in a single column).
    /// </summary>
    public int? ColIndex { get; init; }
}

/// <summary>What kind of content part a <see cref="LayoutPartInfo"/> describes.</summary>
public enum LayoutPartKind
{
    /// <summary>The main document part (<c>document.xml</c>).</summary>
    Document,

    /// <summary>A header part (<c>headerN.xml</c>).</summary>
    Header,

    /// <summary>A footer part (<c>footerN.xml</c>).</summary>
    Footer,
}

/// <summary>
/// The role a section's <c>w:headerReference</c>/<c>w:footerReference</c> assigns to a header/footer part
/// — which pages actually render it. See <see cref="BcWordLayout.Domain.HeaderFooterPartRoles"/> for how
/// this is read (and why a part LIST alone misled callers — GitHub issue #5).
/// </summary>
public enum HeaderFooterRole
{
    /// <summary>The everyday header/footer (an absent <c>w:type</c> counts as default per the spec).</summary>
    Default,

    /// <summary>Rendered on the FIRST page only — and only when the section sets <c>w:titlePg</c>.</summary>
    First,

    /// <summary>Rendered on EVEN pages only — and only when the document enables even/odd headers.</summary>
    Even,
}

/// <summary>
/// One content part of a layout, with the rendering facts a caller needs before addressing an edit at it:
/// which role its section reference assigns it (<see cref="Role"/>), and whether it is the part a
/// <c>layoutPart='header'/'footer'</c> location with no explicit <c>partName</c> resolves to
/// (<see cref="IsDefaultTarget"/>).
/// </summary>
public sealed class LayoutPartInfo
{
    /// <summary>The OOXML part file name, e.g. <c>document.xml</c>, <c>header2.xml</c>.</summary>
    public required string Name { get; init; }

    /// <summary>Whether this is the main document, a header, or a footer part.</summary>
    public required LayoutPartKind Kind { get; init; }

    /// <summary>
    /// The role the first section-reference in document order assigns this part (see
    /// <see cref="HeaderFooterRole"/>); null for the document part, and for a header/footer part no
    /// section references (present in the package, rendered on no page).
    /// </summary>
    public HeaderFooterRole? Role { get; init; }

    /// <summary>
    /// True when this is the part a header/footer-targeted location WITHOUT an explicit part name resolves
    /// to — computed through the SAME selection logic <see cref="BcWordLayout.Domain.LocationResolver"/>
    /// applies, so it can never disagree with where an edit actually lands.
    /// </summary>
    public bool IsDefaultTarget { get; init; }
}

/// <summary>The result of reading a layout: the flat control inventory plus the parts that were walked.</summary>
public sealed class LayoutInventory
{
    public required IReadOnlyList<LayoutControl> Controls { get; init; }

    /// <summary>Names of the OOXML parts walked (document.xml + header/footer parts).</summary>
    public required IReadOnlyList<string> Parts { get; init; }

    /// <summary>
    /// Per-part detail for every entry in <see cref="Parts"/>, same order: part kind, header/footer role,
    /// and whether a partName-less header/footer location resolves to it. See <see cref="LayoutPartInfo"/>.
    /// </summary>
    public IReadOnlyList<LayoutPartInfo> PartDetails { get; init; } = Array.Empty<LayoutPartInfo>();

    /// <summary>
    /// True when the first section sets <c>w:titlePg</c> (a distinct first page): page 1 renders the
    /// <c>first</c>-role header/footer — or none when no such part is referenced — NOT the default one.
    /// </summary>
    public bool HasTitlePage { get; init; }

    /// <summary>
    /// Every table found across all walked parts, described structurally (grid, rows, cells, and which
    /// cell/row holds which control). Empty when the layout has no tables.
    /// </summary>
    public IReadOnlyList<TableStructure> Tables { get; init; } = Array.Empty<TableStructure>();

    public int RepeaterCount => Controls.Count(c => c.Kind == ControlKind.Repeater);
}
