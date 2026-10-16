using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml.Packaging;
using Ds = DocumentFormat.OpenXml.CustomXmlDataProperties;

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

        using var doc = LayoutPackage.Open(docxPath, isEditable: false);
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
    /// <exception cref="InvalidDataException">The package has no BC dataset custom XML part at all.</exception>
    /// <exception cref="LayoutDefectException">
    /// Its BC dataset part is there but cannot be parsed (see <see cref="FindBcPart"/>).
    /// </exception>
    internal static DatasetTree FromMainPart(MainDocumentPart main)
    {
        var (part, root) = FindBcPart(main)
            ?? throw new InvalidDataException(
                $"No BC dataset custom XML part (namespace starting '{OoxmlNames.BcNamespacePrefix}') was found.");

        var storeItemId = part.CustomXmlPropertiesPart?.DataStoreItem?.ItemId?.Value;
        return Build(root, storeItemId);
    }

    /// <summary>Parses the dataset from a standalone exported schema XML file (no OPC package).</summary>
    /// <exception cref="InvalidInputFileException">
    /// The file is not well-formed XML (truncated, an encoding declaration that does not match its bytes, a
    /// DTD, or not XML at all, e.g. a <c>.docx</c> under another extension). The reader's own
    /// <see cref="System.Xml.XmlException"/> names neither the file nor that it is the caller's input, and
    /// used to reach the caller as <c>internal_error</c> (GitHub issue #95).
    /// </exception>
    public static DatasetTree FromSchemaXml(string xmlPath)
    {
        if (!File.Exists(xmlPath))
        {
            throw new FileNotFoundException("The schema XML file does not exist.", xmlPath);
        }

        // ResourceLimits.LoadXDocumentCapped reads the encoding declaration (incl. UTF-16 LE + BOM) from the
        // stream itself, same as a bare XDocument.Load(stream) would - it just also enforces the size
        // cap via a length-limiting wrapper before handing the stream to the XmlReader.
        XDocument xdoc;
        using (var stream = File.OpenRead(xmlPath))
        {
            try
            {
                xdoc = ResourceLimits.LoadXDocumentCapped(stream, $"Schema XML file '{xmlPath}'");
            }
            catch (System.Xml.XmlException ex)
            {
                throw new InvalidInputFileException(
                    xmlPath,
                    InputFileProblem.MalformedXmlFile,
                    $"Schema XML file '{xmlPath}' is not well-formed XML: {ex.Message}",
                    ex);
            }
        }

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
    /// Locates the BC dataset custom XML part — the first readable part whose root namespace starts with the
    /// BC prefix, in relationship order — returning the part together with its loaded root element. Returns
    /// null when the package has no BC dataset part at all.
    /// </summary>
    /// <remarks>
    /// A part that IS recognisably the BC dataset part but does not parse is not "no BC part": that used to
    /// be the answer (the parse exception was swallowed), so every tool misdiagnosed a broken part as a
    /// missing one and never showed the parser's error (GitHub issue #96). When no part parses as BC but
    /// <see cref="Scan"/> recognised an unreadable candidate, this throws a
    /// <see cref="LayoutDefectException"/> naming the part and the parse error instead of returning null.
    /// <para>
    /// SECURITY: the classification walk is <see cref="Scan"/>'s, which parses each part transiently and keeps
    /// none of the trees; the chosen part's root is parsed again here, so no more than one parsed tree is ever
    /// alive at a time however many parts the package carries.
    /// </para>
    /// </remarks>
    /// <exception cref="LayoutDefectException">
    /// No part parses as a BC dataset part, but at least one is recognisably one
    /// (<see cref="LayoutDefect.UnreadableDatasetPart"/>).
    /// </exception>
    /// <exception cref="ResourceLimitExceededException">
    /// The package has more than <see cref="ResourceLimits.MaxCustomXmlParts"/> custom XML parts, or one
    /// exceeds <see cref="ResourceLimits.MaxCustomXmlPartBytes"/>.
    /// </exception>
    internal static (CustomXmlPart Part, XElement Root)? FindBcPart(MainDocumentPart main)
    {
        var scan = Scan(main);
        if (scan.Parts.Count == 0)
        {
            return scan.Unreadable.Count == 0 ? null : throw UnreadableDefect(scan.Unreadable);
        }

        var part = scan.Parts[0].Part;
        return (part, LoadRoot(part));
    }

    /// <summary>
    /// Enumerates every custom XML part whose root parses and whose namespace starts with the BC prefix — the
    /// PART REFERENCE only, NOT its parsed root (see <see cref="Scan"/>). Used by <see cref="LayoutBuilder"/>
    /// to remove every existing BC part before attaching a new one.
    /// </summary>
    /// <exception cref="ResourceLimitExceededException">See <see cref="Scan"/>.</exception>
    internal static IReadOnlyList<CustomXmlPart> FindBcParts(MainDocumentPart main) =>
        Scan(main).Parts.Select(p => p.Part).ToList();

    /// <summary>
    /// Classifies every custom XML part of <paramref name="main"/>, in relationship order, as a readable BC
    /// dataset part, a recognisable-but-unreadable BC dataset part, or unrelated (the Office bibliography part,
    /// MSIP metadata, ...). The ONE walk every lookup above derives from, so the validator, the read tools and
    /// refresh_xml_part cannot disagree about what the package contains.
    /// </summary>
    /// <remarks>
    /// SECURITY. Each part's root is parsed TRANSIENTLY — just long enough to read its namespace — and dropped
    /// before the next part is opened; the result keeps only part references and a few strings, never a tree.
    /// The per-part <see cref="ResourceLimits.MaxCustomXmlPartBytes"/> cap bounds any ONE part's cost, and
    /// keeping no tree is what bounds the SUM: an earlier version of <see cref="FindBcParts"/> returned every
    /// matching part's parsed root at once, so a package of many large BC-namespaced parts (each under the
    /// cap) could exhaust memory in aggregate. A part that fails to parse costs one further bounded read of
    /// its first <see cref="SniffWindowBytes"/> bytes (see <see cref="RecogniseUnreadableBcPart"/>).
    /// </remarks>
    /// <exception cref="ResourceLimitExceededException">
    /// The package has more than <see cref="ResourceLimits.MaxCustomXmlParts"/> custom XML parts, or one
    /// exceeds <see cref="ResourceLimits.MaxCustomXmlPartBytes"/> — deliberately NOT treated as "unreadable":
    /// an oversized part is reported as the crafted-file problem it is.
    /// </exception>
    internal static BcPartScan Scan(MainDocumentPart main)
    {
        EnsurePartCountWithinLimit(main);

        var parts = new List<BcPartInfo>();
        var unreadable = new List<UnreadableBcPart>();
        foreach (var part in main.CustomXmlParts)
        {
            if (TryReadRootNamespace(part, out var ns, out var parseFailure))
            {
                if (ns.StartsWith(OoxmlNames.BcNamespacePrefix, StringComparison.Ordinal))
                {
                    parts.Add(new BcPartInfo(part, PartPath(part), ns, ReadItemId(part)));
                }
            }
            else if (RecogniseUnreadableBcPart(part, parseFailure!) is { } candidate)
            {
                unreadable.Add(candidate);
            }
        }

        return new BcPartScan { Parts = parts, Unreadable = unreadable };
    }

    /// <summary>
    /// Parses <paramref name="part"/>'s root again and returns it — for a part <see cref="Scan"/> has already
    /// classified as a readable BC dataset part, so the parse is expected to succeed.
    /// </summary>
    internal static XElement LoadRoot(CustomXmlPart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        return ResourceLimits.LoadXDocumentCapped(stream, $"Custom XML part '{PartWalker.PartFileName(part)}'").Root
            ?? throw new InvalidDataException($"Custom XML part '{PartPath(part)}' has no root element.");
    }

    /// <summary>The part's own <c>ds:itemID</c> (its itemProps <c>DataStoreItem</c>), or null when it carries none.</summary>
    internal static string? ReadItemId(CustomXmlPart part) =>
        part.CustomXmlPropertiesPart?.DataStoreItem?.ItemId?.Value;

    /// <summary>The part's path inside the package, e.g. <c>customXml/item1.xml</c> — what a person repairing the zip sees.</summary>
    internal static string PartPath(CustomXmlPart part) => part.Uri.OriginalString.TrimStart('/');

    /// <summary>
    /// The <see cref="LayoutDefectException"/> every caller throws for a package whose only BC dataset part(s)
    /// cannot be read — one wording, so get_layout_info, the insert tools and create_layout's schemaSource
    /// all report the same part, the same parser error and the same recognition evidence.
    /// </summary>
    internal static LayoutDefectException UnreadableDefect(IReadOnlyList<UnreadableBcPart> unreadable) =>
        new(string.Join(" ", unreadable.Select(u => u.Describe())), LayoutDefect.UnreadableDatasetPart);

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
    /// uses and returns its namespace (empty for a part with no root), discarding the tree. False with
    /// <paramref name="parseFailure"/> set when the part does not parse — whose BC-ness
    /// <see cref="RecogniseUnreadableBcPart"/> then decides without the parser.
    /// </summary>
    /// <exception cref="ResourceLimitExceededException">
    /// <paramref name="part"/> exceeds <see cref="ResourceLimits.MaxCustomXmlPartBytes"/> (a decompression
    /// "zip bomb" part) — deliberately NOT caught with an ordinary parse failure: an oversized part is reported
    /// as the crafted-file problem it is, rather than as "unreadable" or "not the BC part".
    /// </exception>
    private static bool TryReadRootNamespace(CustomXmlPart part, out string ns, out Exception? parseFailure)
    {
        try
        {
            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            ns = ResourceLimits.LoadXDocumentCapped(stream, $"Custom XML part '{PartWalker.PartFileName(part)}'")
                .Root?.Name.NamespaceName ?? string.Empty;
            parseFailure = null;
            return true;
        }
        catch (ResourceLimitExceededException)
        {
            throw;
        }
        catch (Exception ex)
        {
            ns = string.Empty;
            parseFailure = ex;
            return false;
        }
    }

    /// <summary>
    /// How much of an unparsable part <see cref="RecogniseUnreadableBcPart"/> reads to look for the BC
    /// namespace. The namespace is declared on the root element, which follows nothing longer than an XML
    /// declaration (and, in a broken part, perhaps a DOCTYPE) — a few hundred bytes in every corpus part — so
    /// 64 KB is generous while keeping the read bounded whatever the part's size.
    /// </summary>
    private const int SniffWindowBytes = 64 * 1024;

    /// <summary>
    /// Decides whether a custom XML part that failed to parse is the BC dataset part, WITHOUT the parser:
    /// first from its itemProps <c>ds:schemaRef</c> (every corpus layout that has an itemProps names the BC
    /// namespace there), else from the BC namespace prefix appearing in its first <see cref="SniffWindowBytes"/>
    /// bytes, decoded as UTF-8, UTF-16 LE and UTF-16 BE in turn (the part's own encoding declaration may be
    /// the very thing that is wrong, so it cannot be trusted to say which). Null for a broken part that is
    /// neither — an unrelated part stays "not the BC part", as before GitHub issue #96.
    /// </summary>
    private static UnreadableBcPart? RecogniseUnreadableBcPart(CustomXmlPart part, Exception parseFailure)
    {
        var head = ReadHead(part);

        string? ns;
        string recognisedBy;
        if (SchemaRefBcNamespace(part) is { } fromSchemaRef)
        {
            ns = fromSchemaRef;
            recognisedBy = "its itemProps schemaRef";
        }
        else if (SniffBcNamespace(head) is { } fromBytes)
        {
            ns = fromBytes;
            recognisedBy = "the BC report namespace in its raw content";
        }
        else
        {
            return null;
        }

        return new UnreadableBcPart
        {
            Part = part,
            Name = PartPath(part),
            Namespace = ns,
            ItemId = ReadItemId(part),
            RecognisedBy = recognisedBy,
            Error = parseFailure.Message + DescribeEncodingContradiction(head) + DescribeDoctype(head),
        };
    }

    /// <summary>
    /// A sentence saying the part carries a DOCTYPE, when its head does (in any of the three encodings
    /// <see cref="SniffBcNamespace"/> tries); empty otherwise. The parser's own message for that case tells the
    /// reader to enable DTD processing in its settings - advice no tool caller can act on, and the wrong
    /// repair besides: a BC dataset part never carries a DTD, and DTDs stay prohibited because they can expand
    /// entities.
    /// </summary>
    private static string DescribeDoctype(byte[] head) =>
        new[] { Encoding.UTF8, Encoding.Unicode, Encoding.BigEndianUnicode }
            .Any(e => e.GetString(head).Contains("<!DOCTYPE", StringComparison.Ordinal))
            ? " The part carries a DOCTYPE declaration, which Business Central never writes into a dataset part "
              + "and which is refused here because a DTD can expand entities."
            : string.Empty;

    /// <summary>The first <see cref="SniffWindowBytes"/> bytes of <paramref name="part"/>; empty when even its stream cannot be read.</summary>
    private static byte[] ReadHead(CustomXmlPart part)
    {
        try
        {
            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            var buffer = new byte[SniffWindowBytes];
            var read = stream.ReadAtLeast(buffer, buffer.Length, throwOnEndOfStream: false);
            return buffer[..read];
        }
        catch (Exception)
        {
            // Best effort only: a part whose stream cannot even be opened (a corrupt zip entry) can still be
            // recognised by its schemaRef, and otherwise stays "not the BC part".
            return [];
        }
    }

    /// <summary>The first BC namespace among <paramref name="part"/>'s itemProps <c>ds:schemaRef</c> URIs, or null.</summary>
    private static string? SchemaRefBcNamespace(CustomXmlPart part)
    {
        try
        {
            return part.CustomXmlPropertiesPart?.DataStoreItem?.SchemaReferences?
                .Elements<Ds.SchemaReference>()
                .Select(r => r.Uri?.Value)
                .FirstOrDefault(uri => uri?.StartsWith(OoxmlNames.BcNamespacePrefix, StringComparison.Ordinal) == true);
        }
        catch (Exception)
        {
            // A malformed itemProps is no evidence either way; the raw-content sniff still runs.
            return null;
        }
    }

    /// <summary>
    /// The BC namespace URI found in <paramref name="head"/> — decoded as UTF-8, then UTF-16 LE, then UTF-16
    /// BE, since a wrong decoding cannot produce the ASCII prefix by accident — read up to the first character
    /// that cannot be part of it (a quote, whitespace or markup); null when none of the three contains it.
    /// </summary>
    private static string? SniffBcNamespace(byte[] head)
    {
        foreach (var encoding in new[] { Encoding.UTF8, Encoding.Unicode, Encoding.BigEndianUnicode })
        {
            var text = encoding.GetString(head);
            var start = text.IndexOf(OoxmlNames.BcNamespacePrefix, StringComparison.Ordinal);
            if (start < 0)
            {
                continue;
            }

            var end = start;
            while (end < text.Length && text[end] is not ('"' or '\'' or '<' or '>') && !char.IsWhiteSpace(text[end]))
            {
                end++;
            }

            return text[start..end];
        }

        return null;
    }

    /// <summary>Matches the <c>encoding</c> pseudo-attribute of an XML declaration at the very start of a part.</summary>
    private static readonly Regex EncodingDeclaration =
        new(@"^﻿?\s*<\?xml[^>]*?\bencoding\s*=\s*[""'](?<name>[^""']+)[""']", RegexOptions.Compiled);

    /// <summary>
    /// A sentence naming the contradiction when <paramref name="head"/>'s XML declaration names one encoding
    /// family (UTF-8 or UTF-16) while its bytes — by byte order mark, or by the zero bytes UTF-16 puts around
    /// an ASCII <c>&lt;</c> — are the other; empty otherwise. Three of the six broken shapes GitHub issue #96
    /// reproduced are exactly this, and two of them fail with the parser's least helpful message ("There is
    /// no Unicode byte order mark. Cannot switch to Unicode.", which carries no line or position), so naming
    /// the two encodings is what makes the error actionable.
    /// </summary>
    private static string DescribeEncodingContradiction(byte[] head)
    {
        (Encoding Encoding, bool IsUtf16, string Description)? actual = head switch
        {
            [0xEF, 0xBB, 0xBF, ..] => (Encoding.UTF8, false, "UTF-8 (with a UTF-8 byte order mark)"),
            [0xFF, 0xFE, ..] => (Encoding.Unicode, true, "UTF-16 (with a UTF-16 byte order mark)"),
            [0xFE, 0xFF, ..] => (Encoding.BigEndianUnicode, true, "UTF-16 big-endian (with a byte order mark)"),
            [(byte)'<', 0, ..] => (Encoding.Unicode, true, "UTF-16"),
            [0, (byte)'<', ..] => (Encoding.BigEndianUnicode, true, "UTF-16 big-endian"),
            [(byte)'<', ..] => (Encoding.UTF8, false, "UTF-8"),
            _ => null,
        };

        if (actual is not { } bytes)
        {
            return string.Empty;
        }

        var match = EncodingDeclaration.Match(bytes.Encoding.GetString(head));
        if (!match.Success)
        {
            return string.Empty;
        }

        var declared = match.Groups["name"].Value;
        var declaresUtf16 = declared.StartsWith("utf-16", StringComparison.OrdinalIgnoreCase)
                            || declared.Equals("unicode", StringComparison.OrdinalIgnoreCase);
        var declaresUtf8 = declared.Equals("utf-8", StringComparison.OrdinalIgnoreCase);
        if ((declaresUtf16 && !bytes.IsUtf16) || (declaresUtf8 && bytes.IsUtf16))
        {
            return $" Its XML declaration says encoding=\"{declared}\" but its bytes are {bytes.Description} - what "
                   + "re-saving the part in a text editor typically leaves behind.";
        }

        return string.Empty;
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
