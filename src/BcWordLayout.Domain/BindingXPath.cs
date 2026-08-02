namespace BcWordLayout.Domain;

/// <summary>
/// Helpers for interpreting the (prefixed, indexed) XPath expressions found in Word data bindings,
/// e.g. <c>/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:ItemNo_Line[1]</c>.
/// Only the structural segment names are needed; namespace prefixes and array indices are stripped.
/// (Re-anchoring nested repeater paths to their enclosing item is a merge-engine concern.)
/// </summary>
public static class BindingXPath
{
    /// <summary>
    /// Splits a binding XPath into its ordered element-name segments, dropping namespace prefixes
    /// (<c>ns0:</c>) and positional predicates (<c>[1]</c>). Returns an empty list for null/empty input.
    /// </summary>
    public static IReadOnlyList<string> Segments(string? xpath)
    {
        if (string.IsNullOrWhiteSpace(xpath))
        {
            return Array.Empty<string>();
        }

        var result = new List<string>();
        foreach (var rawStep in xpath.Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            var step = rawStep;

            // Drop a positional predicate such as "[1]".
            var bracket = step.IndexOf('[');
            if (bracket >= 0)
            {
                step = step[..bracket];
            }

            // Drop a namespace prefix such as "ns0:".
            var colon = step.IndexOf(':');
            if (colon >= 0)
            {
                step = step[(colon + 1)..];
            }

            step = step.Trim();
            if (step.Length > 0)
            {
                result.Add(step);
            }
        }

        return result;
    }

    /// <summary>Returns the final (leaf) element name of a binding XPath, or null if none.</summary>
    public static string? LeafName(string? xpath)
    {
        var segments = Segments(xpath);
        return segments.Count == 0 ? null : segments[^1];
    }
}
