using BcWordLayout.Domain;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Merge;

/// <summary>
/// Strips external relationships a renderer would otherwise dereference the moment it OPENS a merged
/// preview copy. A layout cloned from
/// an untrusted repo can carry an <c>attachedTemplate</c> relationship pointing at a UNC path or URL, an
/// externally-LINKED image (<c>a:blip/@r:link</c>, as opposed to an embedded <c>r:embed</c> image part), a
/// linked OLE object, or a mail-merge/subDocument/movie reference. Word reaches out over the network/SMB for
/// any of these purely by OPENING the file — independent of macros: <c>AutomationSecurity = ForceDisable</c>
/// in <see cref="BcWordLayout.Render.WordComConverter"/> blocks VBA execution only, not relationship/template
/// fetches. <c>preview_layout</c> merges into a separate working copy and only ever hands THAT COPY to a
/// converter, so stripping here can never touch the layout the caller passed in (the never-open-the-
/// original-writable invariant).
/// <para>
/// STRIP SET: every relationship surfaced by <see cref="OpenXmlPartContainer.ExternalRelationships"/> — on
/// the PACKAGE ROOT itself (<see cref="OpenXmlPackage"/>, e.g. a relationship in <c>/_rels/.rels</c> not
/// owned by any part — rare, but <see cref="OpenXmlPackage"/> derives from the SAME
/// <see cref="OpenXmlPartContainer"/> base a part does and exposes the identical API, so this costs nothing
/// to cover) AND recursively across every part reachable from it — document.xml, headers/footers,
/// <see cref="DocumentSettingsPart"/> (where <c>attachedTemplate</c> actually lives), glossary document,
/// docProps, etc. (the same traversal shape <see cref="BcWordLayout.Domain.LayoutValidator"/>'s own
/// attachedTemplate check uses for the part-tree portion, extended one level up to the package). This one
/// collection covers <c>attachedTemplate</c>, externally-linked images, linked OLE objects, and mail-merge/
/// subDocument/movie references alike: the OpenXml SDK only asks a linked/external target to become a
/// distinct C# type (<see cref="HyperlinkRelationship"/>) for hyperlinks (see KEEP SET below) — every other
/// external-target relationship, whatever element references it, is a plain
/// <see cref="ExternalRelationship"/> in this one collection.
/// </para>
/// <para>
/// KEEP SET: plain <c>hyperlink</c> relationships (the click-to-follow link a BC layout's body text can
/// legitimately carry, e.g. a "view online" link). These are never touched, for two independent reasons:
/// (1) the OpenXml SDK itself tracks them in a wholly separate
/// <see cref="OpenXmlPartContainer.HyperlinkRelationships"/> collection — a <see cref="HyperlinkRelationship"/>
/// is never surfaced via <see cref="OpenXmlPartContainer.ExternalRelationships"/>, so this stripper
/// structurally cannot reach one even by accident (verified empirically: a hyperlink relationship added via
/// <c>AddHyperlinkRelationship</c> never appears in <c>ExternalRelationships</c>); (2) belt-and-braces, the
/// relationship-type check below explicitly skips anything typed <c>hyperlink</c> in case a future SDK
/// change ever collapsed the two collections. Following a hyperlink is a deliberate user click Word never
/// performs on its own while opening a file — unlike attachedTemplate/linked-image/OLE fetches — so
/// stripping them would only degrade preview fidelity for a legitimate layout with no security upside.
/// </para>
/// <para>
/// ELEMENT CLEANUP: deleting a relationship never touches the XML element that referenced it
/// (<see cref="OpenXmlPartContainer.DeleteExternalRelationship(string)"/> only removes the package-level
/// .rels entry), so a dangling reference would otherwise survive. Every class in
/// <c>DocumentFormat.OpenXml.Wordprocessing</c> deriving from the SDK's own internal
/// <c>RelationshipType</c> base (<see cref="AttachedTemplate"/> / <c>w:attachedTemplate</c>,
/// <see cref="SubDocumentReference"/> / <c>w:subDoc</c>, <see cref="MovieReference"/> / <c>w:movie</c>,
/// <see cref="PrinterSettingsReference"/>, <see cref="SourceReference"/>,
/// <see cref="RecipientDataReference"/>, <see cref="DataSourceReference"/>, <see cref="HeaderSource"/>,
/// <see cref="SourceFileReference"/>) is the SAME "CT_Rel" shape: a standalone element whose ONLY content is
/// a SCHEMA-REQUIRED <c>r:id</c> attribute — confirmed empirically for every one of these nine types
/// individually against <see cref="DocumentFormat.OpenXml.Validation.OpenXmlValidator"/> (an instance built
/// with no <c>Id</c> set always reports "The required attribute 'id' is missing"; merely clearing the
/// attribute would reproduce that on a real document the same way). Any such element is therefore removed
/// OUTRIGHT rather than left with a dangling/empty required attribute. Every OTHER reference —
/// <c>a:blip/@r:link</c> (confirmed: both <c>r:embed</c> and <c>r:link</c> are optional on <c>a:blip</c>), a
/// linked <c>o:OLEObject</c>'s <c>r:id</c>, or a <c>w:frame</c>'s <c>r:id</c> (both independently confirmed
/// optional the same way) — only has its r:-namespaced attribute cleared, leaving the hosting element intact.
/// </para>
/// </summary>
internal static class ExternalRelationshipStripper
{
    private const string RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string HyperlinkRelationshipType = RelationshipsNamespace + "/hyperlink";

    /// <summary>Longest relationship target echoed back in a warning message before truncation — a
    /// crafted URL can itself carry an embedded token/credential in its query string, so the FULL target is
    /// never echoed back verbatim, only enough to identify it.</summary>
    private const int MaxTargetLength = 120;

    /// <summary>
    /// Recursively strips every non-hyperlink external relationship reachable from <paramref name="document"/>
    /// — the package root itself, plus every part reachable from it, direct or nested — scrubbing any
    /// now-dangling reference left in each affected part's own content, and returns one
    /// <see cref="MergeWarning"/> (kind <c>external-relationship-stripped</c>) per relationship removed so
    /// the caller can surface that the preview differs from the source layout.
    /// </summary>
    public static List<MergeWarning> Strip(DocumentFormat.OpenXml.Packaging.OpenXmlPackage document)
    {
        var warnings = new List<MergeWarning>();

        // The package root (OpenXmlPackage) shares OpenXmlPartContainer's relationship API with every part,
        // so a package-root-only relationship (in /_rels/.rels, owned by no part - rare, but structurally
        // possible) is covered too. There is no XML content model at the package level itself (only parts
        // have one), so only the relationship is removed here - nothing to scrub for a dangling reference.
        StripRelationships(document, "(package root)", warnings);

        var visited = new HashSet<OpenXmlPart>();
        foreach (var child in document.Parts)
        {
            StripPart(child.OpenXmlPart, visited, warnings);
        }

        return warnings;
    }

    /// <summary>
    /// <paramref name="depth"/> (0 at the package root's direct parts) enforces
    /// <see cref="ResourceLimits.MaxPartGraphDepth"/> (an uncatchable <see cref="StackOverflowException"/>
    /// would kill the server; a cap fails the call instead) — see
    /// <c>BcWordLayout.Domain.LayoutValidator</c>'s own external-relationship walk for why the
    /// <paramref name="visited"/> cycle guard alone is not enough.
    /// </summary>
    private static void StripPart(OpenXmlPart part, HashSet<OpenXmlPart> visited, List<MergeWarning> warnings, int depth = 0)
    {
        if (depth > ResourceLimits.MaxPartGraphDepth)
        {
            throw ResourceLimits.DepthExceeded("Package part graph", ResourceLimits.MaxPartGraphDepth);
        }

        if (!visited.Add(part))
        {
            return;
        }

        var partName = PartWalker.PartFileName(part);
        var strippedIds = StripRelationships(part, partName, warnings);

        if (strippedIds.Count > 0 && part.RootElement is { } root)
        {
            ScrubDanglingReferences(root, strippedIds);
        }

        foreach (var child in part.Parts)
        {
            StripPart(child.OpenXmlPart, visited, warnings, depth + 1);
        }
    }

    /// <summary>
    /// Deletes every non-hyperlink external relationship directly on <paramref name="container"/> (a part OR
    /// the package root — both are an <see cref="OpenXmlPartContainer"/>), raising one
    /// <c>external-relationship-stripped</c> warning per relationship removed, and returns the set of ids
    /// actually removed (for the caller to scrub dangling references against, when it has content to scrub).
    /// </summary>
    private static HashSet<string> StripRelationships(
        OpenXmlPartContainer container, string location, List<MergeWarning> warnings)
    {
        var strippedIds = new HashSet<string>(StringComparer.Ordinal);

        // Materialize before deleting - ExternalRelationships reflects the container's live relationship
        // collection, and deleting mid-enumeration over it is unsafe.
        foreach (var rel in container.ExternalRelationships.ToList())
        {
            // Defensive only - see class remarks (KEEP SET): a plain hyperlink is structurally unable to
            // reach this collection today, but the check costs nothing and survives a future SDK change
            // that might otherwise silently start stripping legitimate click-to-follow links.
            if (string.Equals(rel.RelationshipType, HyperlinkRelationshipType, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            container.DeleteExternalRelationship(rel.Id);
            strippedIds.Add(rel.Id);

            warnings.Add(new MergeWarning
            {
                Kind = "external-relationship-stripped",
                Message = $"Stripped an external {RelationshipKindName(rel.RelationshipType)} "
                    + $"relationship targeting '{TruncateTarget(rel.Uri)}' before rendering, so the "
                    + "preview's PDF converter never opens/fetches it and this preview's PDF/merged docx "
                    + "differ from the source layout in this respect (a poisoned layout could "
                    + "otherwise make the renderer fetch it on open).",
                Location = location,
            });
        }

        return strippedIds;
    }

    /// <summary>
    /// Walks every descendant of <paramref name="root"/> looking for a now-dangling reference to one of
    /// <paramref name="strippedIds"/>: any <see cref="RelationshipType"/>-derived element (see class remarks
    /// — <see cref="AttachedTemplate"/>/<see cref="SubDocumentReference"/>/<see cref="MovieReference"/> and
    /// the rest of that nine-member family, ALL of which have a schema-required <c>r:id</c> and nothing
    /// else) is removed outright; any other element merely has its matching r:-namespaced attribute (e.g.
    /// <c>a:blip/@r:embed</c> or <c>@r:link</c>, a linked OLE object's or frame's <c>r:id</c>) removed,
    /// leaving the element itself intact.
    /// </summary>
    private static void ScrubDanglingReferences(OpenXmlElement root, HashSet<string> strippedIds)
    {
        foreach (var element in root.Descendants().ToList())
        {
            if (element is RelationshipType requiredIdElement)
            {
                if (requiredIdElement.Id?.Value is { } requiredRelId && strippedIds.Contains(requiredRelId))
                {
                    element.Remove();
                }

                continue;
            }

            foreach (var attr in element.GetAttributes()
                .Where(a => a.NamespaceUri == RelationshipsNamespace && a.Value is not null && strippedIds.Contains(a.Value))
                .ToList())
            {
                element.RemoveAttribute(attr.LocalName, attr.NamespaceUri);
            }
        }
    }

    /// <summary>The last URI segment of a relationship type, e.g. <c>"attachedTemplate"</c> or
    /// <c>"image"</c> - used only to make the warning message readable.</summary>
    private static string RelationshipKindName(string relationshipType)
    {
        var slash = relationshipType.LastIndexOf('/');
        return slash >= 0 && slash + 1 < relationshipType.Length ? relationshipType[(slash + 1)..] : relationshipType;
    }

    private static string TruncateTarget(Uri? uri)
    {
        var text = uri?.ToString() ?? "(no target)";
        return text.Length <= MaxTargetLength ? text : text[..MaxTargetLength] + "…";
    }
}
