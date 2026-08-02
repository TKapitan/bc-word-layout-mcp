using BcWordLayout.Domain;

namespace BcWordLayout.Merge;

/// <summary>
/// Re-anchoring helper for nested repeater/field XPaths. Bindings store fully absolute XPaths (e.g.
/// <c>/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:ItemNo_Line[1]</c>), but once a repeater
/// row is cloned for a specific data node, any inner binding must resolve against THAT row's node rather
/// than the document root. Dropping the leading location steps already represented by the current
/// context node turns the absolute XPath into the correct relative remainder to evaluate against it —
/// PROVIDED those leading steps actually describe the same element path as the enclosing repeater's own
/// binding (see <see cref="Remainder"/>'s <c>contextXPath</c> parameter); a step COUNT match alone is not
/// enough, since an equal-depth sibling binding has the right count but names different elements.
/// </summary>
internal static class XPathReanchor
{
    /// <summary>
    /// Splits <paramref name="xpath"/> into raw location steps, retaining namespace prefixes and
    /// positional predicates (e.g. <c>ns0:Line[1]</c>) — unlike <see cref="BcWordLayout.Domain.BindingXPath.Segments"/>,
    /// which strips them for structural comparison. Returns an empty list for null/blank input.
    /// </summary>
    internal static IReadOnlyList<string> RawSteps(string? xpath) =>
        string.IsNullOrWhiteSpace(xpath)
            ? Array.Empty<string>()
            : xpath.Split('/', StringSplitOptions.RemoveEmptyEntries);

    /// <summary>
    /// True when the path uses the descendant-or-self axis (<c>//</c>). Splitting on <c>/</c> with
    /// <see cref="StringSplitOptions.RemoveEmptyEntries"/> would silently collapse that into an ordinary
    /// step boundary, so callers must check this first and treat such paths as unsupported.
    /// </summary>
    internal static bool HasDescendantAxis(string? xpath) =>
        !string.IsNullOrEmpty(xpath) && xpath.Contains("//", StringComparison.Ordinal);

    /// <summary>
    /// Returns the relative XPath remaining after dropping the first <paramref name="consumedSteps"/> raw
    /// steps of <paramref name="xpath"/> (empty string when every step is consumed — the binding targets
    /// the context node itself), or null when re-anchoring is not possible: a <c>//</c> axis, a step count
    /// that is negative or exceeds the path's own length, or — when <paramref name="contextXPath"/> is
    /// supplied — a STRUCTURAL mismatch between the dropped prefix and the path <paramref name="contextXPath"/>
    /// itself describes (see <see cref="StructuralPrefixMatches"/>). Callers should treat null as "fall back
    /// to evaluating the original absolute XPath from the document root" and raise a warning.
    /// </summary>
    /// <param name="xpath">The (fully absolute) binding XPath to re-anchor.</param>
    /// <param name="consumedSteps">
    /// How many of <paramref name="xpath"/>'s leading raw steps the current context node already
    /// represents — always the enclosing repeater's own raw step count (see
    /// <see cref="BcWordLayout.Merge.MergeEngine.ProcessRepeater"/>).
    /// </param>
    /// <param name="contextXPath">
    /// The absolute XPath of the repeater whose row context node <paramref name="consumedSteps"/> is
    /// counted against, or null at the top of a walk (document root, no enclosing repeater — in which case
    /// <paramref name="consumedSteps"/> must be 0 and no structural check is needed). Required whenever
    /// <paramref name="consumedSteps"/> is greater than zero: dropping steps without knowing what they were
    /// supposed to be is exactly the bug this parameter exists to prevent.
    /// </param>
    internal static string? Remainder(string? xpath, int consumedSteps, string? contextXPath)
    {
        if (string.IsNullOrWhiteSpace(xpath) || HasDescendantAxis(xpath))
        {
            return null;
        }

        var steps = RawSteps(xpath);
        if (consumedSteps < 0 || consumedSteps > steps.Count)
        {
            return null;
        }

        if (consumedSteps > 0 && !StructuralPrefixMatches(xpath, consumedSteps, contextXPath))
        {
            return null;
        }

        return string.Join('/', steps.Skip(consumedSteps));
    }

    /// <summary>
    /// True when the first <paramref name="consumedSteps"/> STRUCTURAL segments of <paramref name="xpath"/>
    /// (element local names, namespace prefixes and positional predicates ignored — see
    /// <see cref="BindingXPath.Segments"/>) are IDENTICAL, in order, to <paramref name="contextXPath"/>'s
    /// own full structural segment list. This is the guard the step-count-only version of re-anchoring was
    /// missing: an equal-depth binding that names a DIFFERENT element (e.g. a Header-level sibling field
    /// bound inside a same-depth Line repeater row) has the right step COUNT but the wrong path, and must
    /// not be treated as "already inside this context." A null/blank <paramref name="contextXPath"/>, or one
    /// whose own segment count does not equal <paramref name="consumedSteps"/>, is treated as a mismatch —
    /// re-anchoring only ever drops a prefix it can positively confirm.
    /// </summary>
    private static bool StructuralPrefixMatches(string xpath, int consumedSteps, string? contextXPath)
    {
        if (string.IsNullOrWhiteSpace(contextXPath))
        {
            return false;
        }

        var contextSegments = BindingXPath.Segments(contextXPath);
        if (contextSegments.Count != consumedSteps)
        {
            return false;
        }

        var candidateSegments = BindingXPath.Segments(xpath);
        if (candidateSegments.Count < consumedSteps)
        {
            return false;
        }

        for (var i = 0; i < consumedSteps; i++)
        {
            if (!string.Equals(candidateSegments[i], contextSegments[i], StringComparison.Ordinal))
            {
                return false;
            }
        }

        return true;
    }
}
