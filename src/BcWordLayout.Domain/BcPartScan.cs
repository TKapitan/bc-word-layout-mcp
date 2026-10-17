using DocumentFormat.OpenXml.Packaging;

namespace BcWordLayout.Domain;

/// <summary>
/// What <see cref="SchemaProvider.Scan"/> found among a package's custom XML parts: every readable BC dataset
/// part and every part that is recognisably one but cannot be parsed. Holds part references and strings
/// only — never a parsed tree (see <see cref="SchemaProvider.Scan"/>'s security remarks).
/// </summary>
internal sealed class BcPartScan
{
    /// <summary>Every custom XML part whose root parses and is in a BC report namespace, in relationship order.</summary>
    public required IReadOnlyList<BcPartInfo> Parts { get; init; }

    /// <summary>
    /// Every custom XML part that failed to parse but is recognisably a BC dataset part (GitHub issue #96),
    /// in relationship order. Unrelated broken parts are not listed.
    /// </summary>
    public required IReadOnlyList<UnreadableBcPart> Unreadable { get; init; }

    /// <summary>
    /// The readable part the layout is READ against: the only one, or — when there are several — the one its
    /// bindings use (GitHub issue #101, see <see cref="SchemaProvider.ChooseBcPart"/>). Null when no part is readable.
    /// </summary>
    public required BcPartInfo? Chosen { get; init; }

    /// <summary>Why <see cref="Chosen"/> was chosen, as a clause that reads after its name; null unless there were several readable parts.</summary>
    public required string? ChoiceReason { get; init; }

    /// <summary>How many BC dataset parts the package carries, readable or not — "exactly one" is the valid shape.</summary>
    public int Count => Parts.Count + Unreadable.Count;

    /// <summary>
    /// Every BC dataset part, by package path and report (or as unreadable), in relationship order — e.g.
    /// <c>'customXml/item1.xml' (Job_Quote/1016, itemID {...}), 'customXml/item2.xml' (not well-formed XML)</c>.
    /// </summary>
    public string DescribeParts() =>
        string.Join(
            ", ",
            Parts.Select(p =>
                {
                    var identity = SchemaProvider.ParseIdentity(p.Namespace, p.ItemId);
                    return $"'{p.Name}' ({identity.ReportName}/{identity.ReportId}, "
                           + (p.ItemId is null ? "no itemProps" : $"itemID {p.ItemId}") + ")";
                })
                .Concat(Unreadable.Select(u => $"'{u.Name}' (not well-formed XML)")));
}

/// <summary>A readable BC dataset custom XML part, with the facts a caller needs without re-parsing it.</summary>
/// <param name="Part">The part itself.</param>
/// <param name="Name">Its path inside the package, e.g. <c>customXml/item1.xml</c>.</param>
/// <param name="Namespace">Its root element's namespace — the report it describes.</param>
/// <param name="ItemId">Its itemProps <c>ds:itemID</c>; null when it carries no itemProps (SOLUTION-DESIGN §2.1).</param>
internal sealed record BcPartInfo(CustomXmlPart Part, string Name, string Namespace, string? ItemId);

/// <summary>
/// A custom XML part that is recognisably the BC dataset part — by its itemProps <c>ds:schemaRef</c> or by the
/// BC report namespace in its raw bytes — but is not well-formed XML (GitHub issue #96).
/// </summary>
internal sealed class UnreadableBcPart
{
    /// <summary>The part itself — what refresh_xml_part replaces.</summary>
    public required CustomXmlPart Part { get; init; }

    /// <summary>Its path inside the package, e.g. <c>customXml/item1.xml</c>.</summary>
    public required string Name { get; init; }

    /// <summary>The BC report namespace its schemaRef or raw content names — the report it describes.</summary>
    public required string Namespace { get; init; }

    /// <summary>Its itemProps <c>ds:itemID</c>, when it has one — kept by a refresh that replaces it.</summary>
    public string? ItemId { get; init; }

    /// <summary>The evidence it was recognised by, phrased to follow "recognised by".</summary>
    public required string RecognisedBy { get; init; }

    /// <summary>The parser's own message (with its line/position whenever the parser reports one), plus any encoding contradiction found in the bytes.</summary>
    public required string Error { get; init; }

    /// <summary>One sentence naming the part, the evidence and the parse error — the wording every tool reports.</summary>
    public string Describe() =>
        $"The BC dataset custom XML part '{Name}' (recognised by {RecognisedBy}, '{Namespace}') is not "
        + $"well-formed XML, so no dataset can be read from it: {Error}";
}
