using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Domain;

/// <summary>
/// The external attached template a layout's settings can link: the <c>w:attachedTemplate</c> setting and the
/// external relationship it names, where Word records the template a document was created from. Microsoft's own
/// stock layouts carry one into the folders of the developers who designed them (four corpus captures, e.g.
/// <c>Simplified Sales Document.dotx</c>), and so does any design started from a custom template.
/// </summary>
/// <remarks>
/// <para>
/// Business Central refuses to upload a layout that links one: "UnsupportedExternalContentFields: The Word layout
/// contains unsupported fields or external references: ExternalRelationship:attachedTemplate" (a BC 28 sandbox,
/// 2026-10-24 - the same stock layouts uploaded in August). Nothing BC renders comes from the link: the
/// template's styles are already copied into the package. So every tool that writes a layout removes it
/// (<c>ToolGuards</c>' save choreography and <see cref="LayoutBuilder"/>), and <see cref="LayoutValidator"/>
/// reports one still there as an error.
/// </para>
/// <para>
/// The walk covers every part reachable from the main document part, not only its settings part (a glossary
/// document has a settings part of its own), with a cycle guard and the
/// <see cref="ResourceLimits.MaxPartGraphDepth"/> cap: the visited set rules out an infinite CYCLE, but a
/// crafted package could still chain thousands of acyclic parts, and an uncatchable
/// <see cref="StackOverflowException"/> would kill the server where the cap fails the call.
/// </para>
/// </remarks>
public static class AttachedTemplateLink
{
    /// <summary>Longest target echoed back before it is cut: a crafted URL can carry a token in its query.</summary>
    private const int MaxTargetLength = 120;

    /// <summary>
    /// Removes every external attached-template relationship reachable from <paramref name="doc"/>'s main
    /// document part, together with the <c>w:attachedTemplate</c> element naming it, saves each part it changed,
    /// and returns the targets it unlinked (see <see cref="Describe"/>); empty when there was none.
    /// </summary>
    public static IReadOnlyList<string> Remove(WordprocessingDocument doc)
    {
        if (doc.MainDocumentPart is not { } main)
        {
            return [];
        }

        var unlinked = new List<string>();
        foreach (var group in Find(main).ToList().GroupBy(link => link.Part))
        {
            var ids = new HashSet<string>(StringComparer.Ordinal);
            foreach (var (part, relationship) in group)
            {
                unlinked.Add(Describe(relationship.Uri));
                ids.Add(relationship.Id);
                part.DeleteExternalRelationship(relationship.Id);
            }

            if (group.Key.RootElement is { } root)
            {
                foreach (var setting in root.Descendants<AttachedTemplate>().Where(a => a.Id?.Value is { } id && ids.Contains(id)).ToList())
                {
                    setting.Remove();
                }

                root.Save();
            }
        }

        return unlinked;
    }

    /// <summary>
    /// Every external attached-template relationship reachable from <paramref name="main"/>, with the part that
    /// owns it, in part-tree order.
    /// </summary>
    internal static IEnumerable<(OpenXmlPart Part, ExternalRelationship Relationship)> Find(MainDocumentPart main) =>
        Find(main, new HashSet<OpenXmlPart>(), 0);

    /// <summary>A link's target as messages show it: cut to <see cref="MaxTargetLength"/> characters.</summary>
    internal static string Describe(Uri? target)
    {
        var text = target?.ToString() ?? "(no target)";
        return text.Length <= MaxTargetLength ? text : text[..MaxTargetLength] + "…";
    }

    private static IEnumerable<(OpenXmlPart Part, ExternalRelationship Relationship)> Find(
        OpenXmlPart part, HashSet<OpenXmlPart> visited, int depth)
    {
        if (depth > ResourceLimits.MaxPartGraphDepth)
        {
            throw ResourceLimits.DepthExceeded("Package part graph", ResourceLimits.MaxPartGraphDepth);
        }

        if (!visited.Add(part))
        {
            yield break;
        }

        foreach (var relationship in part.ExternalRelationships)
        {
            if (relationship.RelationshipType.EndsWith("/attachedTemplate", StringComparison.Ordinal))
            {
                yield return (part, relationship);
            }
        }

        foreach (var child in part.Parts)
        {
            foreach (var link in Find(child.OpenXmlPart, visited, depth + 1))
            {
                yield return link;
            }
        }
    }
}
