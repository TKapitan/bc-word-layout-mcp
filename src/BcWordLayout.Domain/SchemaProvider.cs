using System.Xml.Linq;
using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml.Packaging;

namespace BcWordLayout.Domain;

/// <summary>
/// Parses a Business Central report dataset into a <see cref="DatasetTree"/> from either a
/// <c>.docx</c> layout (locating the BC custom XML part) or a standalone exported schema <c>.xml</c>.
/// Both paths converge on the same tree-building logic.
/// </summary>
public static class SchemaProvider
{
    /// <summary>
    /// Parses the dataset from the BC custom XML part inside a <c>.docx</c> layout. Skips any unrelated
    /// custom XML part (e.g. the Office bibliography part) by matching the BC namespace prefix.
    /// </summary>
    public static DatasetTree FromLayout(string docxPath)
    {
        if (!File.Exists(docxPath))
        {
            throw new FileNotFoundException("The layout file does not exist.", docxPath);
        }

        using var doc = WordprocessingDocument.Open(docxPath, false);
        return FromLayout(doc);
    }

    /// <summary>Parses the dataset from an already-open layout package (lets callers share one handle).</summary>
    public static DatasetTree FromLayout(WordprocessingDocument doc)
    {
        var main = doc.MainDocumentPart
            ?? throw new InvalidDataException("Layout has no main document part.");

        return FromMainPart(main);
    }

    /// <summary>Builds the dataset tree from an already-open main document part.</summary>
    internal static DatasetTree FromMainPart(MainDocumentPart main)
    {
        var (part, root) = FindBcPart(main)
            ?? throw new InvalidDataException(
                $"No BC dataset custom XML part (namespace starting '{OoxmlNames.BcNamespacePrefix}') was found.");

        var storeItemId = part.CustomXmlPropertiesPart?.DataStoreItem?.ItemId?.Value;
        return Build(root, storeItemId);
    }

    /// <summary>Parses the dataset from a standalone exported schema XML file (no OPC package).</summary>
    public static DatasetTree FromSchemaXml(string xmlPath)
    {
        if (!File.Exists(xmlPath))
        {
            throw new FileNotFoundException("The schema XML file does not exist.", xmlPath);
        }

        // ResourceLimits.LoadXDocumentCapped reads the encoding declaration (incl. UTF-16 LE + BOM) from the
        // stream itself, same as a bare XDocument.Load(stream) would - it just also enforces the size
        // cap via a length-limiting wrapper before handing the stream to the XmlReader.
        using var stream = File.OpenRead(xmlPath);
        var xdoc = ResourceLimits.LoadXDocumentCapped(stream, $"Schema XML file '{xmlPath}'");
        var root = xdoc.Root
            ?? throw new InvalidDataException("Schema XML has no root element.");

        if (root.Name.LocalName != OoxmlNames.RootElementName)
        {
            throw new InvalidDataException(
                $"Schema root element is '{root.Name.LocalName}', expected '{OoxmlNames.RootElementName}'.");
        }

        // The local name alone is not enough: a non-BC XML could coincidentally use the same root element
        // name. Require the namespace to actually be a BC dataset namespace too, matching the same check
        // FindBcParts applies to a .docx's custom XML parts - otherwise a bogus schema would be accepted
        // here and silently orphan every existing binding once used as a refresh_xml_part/create_layout
        // source.
        if (!root.Name.NamespaceName.StartsWith(OoxmlNames.BcNamespacePrefix, StringComparison.Ordinal))
        {
            throw new InvalidDataException(
                $"Schema root namespace '{root.Name.NamespaceName}' does not start with "
                + $"'{OoxmlNames.BcNamespacePrefix}'; this does not look like a BC dataset schema.");
        }

        return Build(root, storeItemId: null);
    }

    /// <summary>
    /// Locates the first custom XML part whose root namespace starts with the BC prefix, returning the part
    /// together with its loaded root element. Returns null when none is present.
    /// </summary>
    /// <remarks>
    /// SECURITY. Parts are parsed ONE AT A TIME, stopping the
    /// moment a match is found — this method never retains more than one part's parsed <see cref="XElement"/>
    /// tree simultaneously, and (for the common case where the BC part is found early) may not even attempt
    /// to parse every part in the package. This is deliberately a SEPARATE walk from <see cref="FindBcParts"/>
    /// rather than <c>FindBcParts(main).FirstOrDefault()</c>: the latter would force <see cref="FindBcParts"/>
    /// to finish enumerating (and therefore transiently parsing) every part in the package before this method
    /// could return, for no benefit — this method only ever needs the FIRST match.
    /// </remarks>
    /// <exception cref="ResourceLimitExceededException">
    /// The package has more than <see cref="ResourceLimits.MaxCustomXmlParts"/> custom XML parts, or one
    /// exceeds <see cref="ResourceLimits.MaxCustomXmlPartBytes"/>.
    /// </exception>
    internal static (CustomXmlPart Part, XElement Root)? FindBcPart(MainDocumentPart main)
    {
        EnsurePartCountWithinLimit(main);

        foreach (var part in main.CustomXmlParts)
        {
            if (TryLoadBcNamespacedRoot(part, out var root))
            {
                return (part, root!);
            }
        }

        return null;
    }

    /// <summary>
    /// Enumerates every custom XML part whose root namespace starts with the BC prefix — the PART REFERENCE
    /// only, NOT its parsed root. Malformed / non-XML parts are skipped. This is the single shared
    /// implementation used both by <see cref="LayoutValidator"/> (to count/report duplicate BC parts) and by
    /// <see cref="LayoutBuilder"/> (to remove every existing BC part before attaching a new one) — neither
    /// caller needs a parsed root for any part beyond the first (see <see cref="FindBcPart"/> for that one).
    /// </summary>
    /// <remarks>
    /// SECURITY. The
    /// PREVIOUS version of this method returned <c>(CustomXmlPart, XElement)</c> pairs for every matching
    /// part, retaining every one's fully parsed root SIMULTANEOUSLY in the returned list — the per-part
    /// <see cref="ResourceLimits.MaxCustomXmlPartBytes"/> cap bounds any ONE part's cost, but not the SUM: a
    /// package with many large BC-namespaced parts (each individually under the cap) could still exhaust
    /// memory in aggregate. This method now parses each part's root TRANSIENTLY just to check its namespace
    /// (via <see cref="TryLoadBcNamespacedRoot"/>) and discards it immediately either way — only the
    /// <see cref="CustomXmlPart"/> reference itself is retained in the returned list, so at most one part's
    /// parsed tree is ever alive at a time regardless of how many parts match or how large they are.
    /// </remarks>
    /// <exception cref="ResourceLimitExceededException">
    /// The package has more than <see cref="ResourceLimits.MaxCustomXmlParts"/> custom XML parts, or one
    /// exceeds <see cref="ResourceLimits.MaxCustomXmlPartBytes"/> .
    /// </exception>
    internal static IReadOnlyList<CustomXmlPart> FindBcParts(MainDocumentPart main)
    {
        EnsurePartCountWithinLimit(main);

        var result = new List<CustomXmlPart>();
        foreach (var part in main.CustomXmlParts)
        {
            if (TryLoadBcNamespacedRoot(part, out _))
            {
                result.Add(part);
            }
        }

        return result;
    }

    /// <summary>
    /// Rejects up front (before any part is even opened) a package carrying more than
    /// <see cref="ResourceLimits.MaxCustomXmlParts"/> custom XML parts — bounds the per-part iteration/parse
    /// LOOP itself, not just any one part's size.
    /// </summary>
    private static void EnsurePartCountWithinLimit(MainDocumentPart main)
    {
        var count = main.CustomXmlParts.Count();
        if (count > ResourceLimits.MaxCustomXmlParts)
        {
            throw ResourceLimits.PartCountExceeded(count, ResourceLimits.MaxCustomXmlParts);
        }
    }

    /// <summary>
    /// Parses <paramref name="part"/>'s root through the same size-capped loader <see cref="FromSchemaXml"/>
    /// uses. Returns true with <paramref name="root"/> set when it parses cleanly AND its namespace starts
    /// with the BC prefix; false (<paramref name="root"/> null) for a malformed/non-XML part, or a
    /// well-formed part in some other namespace (e.g. the Office bibliography part some corpus layouts
    /// carry) — either way, the parsed tree is not retained by this method once it returns.
    /// </summary>
    /// <exception cref="ResourceLimitExceededException">
    /// <paramref name="part"/> exceeds <see cref="ResourceLimits.MaxCustomXmlPartBytes"/> (a decompression
    /// "zip bomb" part) — deliberately NOT swallowed by the generic catch below that skips an ordinary
    /// malformed/non-XML part: an oversized part is reported as the crafted-file problem it is, rather than
    /// silently treated as "not the BC part" (which would otherwise mislead the caller with a plain "no BC
    /// part found" once the real cause was a bomb).
    /// </exception>
    private static bool TryLoadBcNamespacedRoot(CustomXmlPart part, out XElement? root)
    {
        XElement? parsed;
        try
        {
            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            parsed = ResourceLimits.LoadXDocumentCapped(stream, $"Custom XML part '{PartWalker.PartFileName(part)}'").Root;
        }
        catch (ResourceLimitExceededException)
        {
            // Propagate distinctly rather than falling into the generic catch below, which exists only for
            // a genuinely malformed/non-XML part - see this method's own exception doc.
            throw;
        }
        catch
        {
            // A malformed / non-XML custom part is simply not the BC part.
            root = null;
            return false;
        }

        if (parsed is not null &&
            parsed.Name.NamespaceName.StartsWith(OoxmlNames.BcNamespacePrefix, StringComparison.Ordinal))
        {
            root = parsed;
            return true;
        }

        root = null;
        return false;
    }

    /// <summary>Builds the dataset tree from the loaded <c>NavWordReportXmlPart</c> root element.</summary>
    private static DatasetTree Build(XElement root, string? storeItemId)
    {
        var ns = root.Name.NamespaceName;
        var identity = ParseIdentity(ns, storeItemId);

        var rootItem = new DataItem { Name = root.Name.LocalName, Path = "/", IsSystem = false };
        foreach (var child in root.Elements())
        {
            var isSystem = child.Name.LocalName == OoxmlNames.SystemNodeName;
            BuildNode(child, rootItem, "", isSystem, depth: 1);
        }

        return new DatasetTree { Report = identity, Root = rootItem };
    }

    /// <summary>
    /// Recursively classifies <paramref name="element"/> as a data item (has element children) or a
    /// leaf column, attaching it under <paramref name="parent"/>. The system flag propagates down the
    /// <c>BCReportInformation</c> subtree.
    /// </summary>
    /// <remarks>
    /// <paramref name="depth"/> (1 at the root's direct children) is the ONE enforcement point for
    /// <see cref="ResourceLimits.MaxSchemaDepth"/> (an uncatchable <see cref="StackOverflowException"/> would
    /// kill the server; a cap fails the call instead): every
    /// <see cref="DataItem"/> tree in the process is built HERE and nowhere else (see
    /// <see cref="Models.DataItem"/>'s own remarks - no other production code constructs one), so capping
    /// recursion in this one method structurally bounds every later walk of the SAME tree too -
    /// <c>BcWordLayout.Merge.SampleDataGenerator.BuildInstance</c> and
    /// <c>BcWordLayout.McpHost.Tools.ToolGuards.ToDataItemDto</c> each recurse over a <see cref="DataItem"/>
    /// tree this method already produced within the cap, so neither needs (or has) its own separate depth
    /// counter.
    /// </remarks>
    private static void BuildNode(XElement element, DataItem parent, string parentPath, bool isSystem, int depth)
    {
        if (depth > ResourceLimits.MaxSchemaDepth)
        {
            throw ResourceLimits.DepthExceeded("Schema data-item", ResourceLimits.MaxSchemaDepth);
        }

        var name = element.Name.LocalName;
        var path = parentPath + "/" + name;
        var childElements = element.Elements().ToList();

        if (childElements.Count == 0)
        {
            parent.Columns.Add(new DatasetColumn { Name = name, Path = path });
            return;
        }

        var item = new DataItem { Name = name, Path = path, IsSystem = isSystem };
        parent.Children.Add(item);
        foreach (var child in childElements)
        {
            BuildNode(child, item, path, isSystem, depth + 1);
        }
    }

    /// <summary>
    /// Parses <c>urn:microsoft-dynamics-nav/reports/&lt;ReportName&gt;/&lt;ReportId&gt;/</c> into an identity.
    /// Tolerates a trailing slash and unexpected shapes (falls back to best-effort values).
    /// </summary>
    internal static ReportIdentity ParseIdentity(string ns, string? storeItemId)
    {
        var reportName = "";
        var reportId = "";

        var marker = ns.IndexOf(OoxmlNames.BcNamespacePrefix, StringComparison.Ordinal);
        if (marker >= 0)
        {
            var tail = ns[(marker + OoxmlNames.BcNamespacePrefix.Length)..].Trim('/');
            var parts = tail.Split('/', StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length >= 1)
            {
                reportName = parts[0];
            }

            if (parts.Length >= 2)
            {
                reportId = parts[1];
            }
        }

        return new ReportIdentity
        {
            ReportName = reportName,
            ReportId = reportId,
            Namespace = ns,
            StoreItemId = storeItemId,
        };
    }
}
