using System.Text.RegularExpressions;

namespace BcWordLayout.Domain;

/// <summary>
/// Helpers for interpreting the (prefixed, indexed) XPath expressions found in Word data bindings,
/// e.g. <c>/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:ItemNo_Line[1]</c>.
/// Resolution needs only the structural segment names, with namespace prefixes and array indices stripped;
/// the prefixes themselves, and the binding's <c>w:prefixMappings</c> that must declare them, are read
/// separately (<see cref="Prefixes"/>, <see cref="PrefixDeclarations"/>).
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

    /// <summary>
    /// True when <paramref name="xpath"/>'s root step names the BC dataset root element
    /// (<c>NavWordReportXmlPart</c>), whatever its prefix — the mark of a binding into the BC dataset part
    /// rather than into some other data store (a document-property control binds <c>coreProperties</c>).
    /// </summary>
    internal static bool IsBcDatasetBinding(string? xpath)
    {
        var segments = Segments(xpath);
        return segments.Count > 0 && string.Equals(segments[0], OoxmlNames.RootElementName, StringComparison.Ordinal);
    }

    /// <summary>
    /// The distinct namespace prefixes <paramref name="xpath"/>'s steps use (<c>ns0</c> for
    /// <c>/ns0:NavWordReportXmlPart[1]/ns0:Header[1]</c>), in first-use order; empty for null/empty input or
    /// an XPath whose steps carry no prefix. Positional predicates are dropped first, exactly as
    /// <see cref="Segments"/> drops them.
    /// </summary>
    internal static IReadOnlyList<string> Prefixes(string? xpath)
    {
        var result = new List<string>();
        if (string.IsNullOrWhiteSpace(xpath))
        {
            return result;
        }

        foreach (var rawStep in xpath.Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            var step = rawStep;
            var bracket = step.IndexOf('[');
            if (bracket >= 0)
            {
                step = step[..bracket];
            }

            var colon = step.IndexOf(':');
            if (colon > 0)
            {
                var prefix = step[..colon].Trim();
                if (prefix.Length > 0 && !result.Contains(prefix, StringComparer.Ordinal))
                {
                    result.Add(prefix);
                }
            }
        }

        return result;
    }

    /// <summary>
    /// Matches one namespace declaration inside a <c>w:prefixMappings</c> value — <c>xmlns:ns0='&lt;uri&gt;'</c>,
    /// the form Word and every corpus layout write. Either XML quote is accepted, as in any XML namespace
    /// declaration.
    /// </summary>
    private static readonly Regex PrefixDeclarationPattern =
        new(@"xmlns:(?<prefix>\w+)\s*=\s*(?<q>['""])(?<uri>.*?)\k<q>", RegexOptions.Compiled);

    /// <summary>
    /// Parses the <c>xmlns:&lt;prefix&gt;='&lt;uri&gt;'</c> declarations out of a <c>w:prefixMappings</c>
    /// value, in order (a prefix declared twice keeps its FIRST uri, the one an XPath engine would bind).
    /// Empty when the value is null/empty or declares nothing — notably a BARE uri with no <c>xmlns:</c>
    /// around it, a shape real base-app layouts ship (see <see cref="SdtInspector.ExtractBcNamespace"/>).
    /// The ONE parser both the merge (registering prefixes for XPath evaluation) and validation
    /// (<c>binding-prefix-mapping</c>) use, so the two cannot disagree about what a binding declares.
    /// </summary>
    internal static IReadOnlyList<(string Prefix, string Uri)> PrefixDeclarations(string? prefixMappings)
    {
        var result = new List<(string Prefix, string Uri)>();
        if (string.IsNullOrWhiteSpace(prefixMappings))
        {
            return result;
        }

        foreach (Match m in PrefixDeclarationPattern.Matches(prefixMappings))
        {
            var prefix = m.Groups["prefix"].Value;
            if (!result.Exists(d => string.Equals(d.Prefix, prefix, StringComparison.Ordinal)))
            {
                result.Add((prefix, m.Groups["uri"].Value));
            }
        }

        return result;
    }
}
