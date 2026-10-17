using System.Text.RegularExpressions;

namespace BcWordLayout.Tests;

/// <summary>
/// GitHub issue #99: tool descriptions, hints and messages are prose split into string literals joined with
/// <c>+</c> at line ends, and nine joins in <c>validate_layout</c>'s description had no space on either side.
/// MCP clients show that text to the model verbatim, so it read "anda", "thecontrol" and
/// "level='full'additionally". This scans every literal-to-literal join under <c>src/</c> and requires
/// whitespace on one side of it.
/// </summary>
public class StringSeamGuardTests
{
    // Two string literals (optionally interpolated or verbatim) joined by '+', across any whitespace.
    private static readonly Regex Seam = new(
        @"(?<left>\$?@?""(?:[^""\\\r\n]|\\.)*"")\s*\+\s*(?<right>\$?@?""(?:[^""\\\r\n]|\\.)*"")",
        RegexOptions.Compiled);

    // An encoded blob split over lines (PlaceholderImage's base64 PNG), where a join without spaces is the point.
    private static readonly Regex EncodedRun = new("^[A-Za-z0-9+/=]{16,}$", RegexOptions.Compiled);

    [Fact]
    public void Every_join_of_two_string_literals_keeps_a_space_at_the_seam()
    {
        var offenders = new List<string>();
        foreach (var file in RepoFiles.SourceFiles("src"))
        {
            var text = File.ReadAllText(file);
            foreach (Match match in Seam.Matches(text))
            {
                var left = Body(match.Groups["left"].Value);
                var right = Body(match.Groups["right"].Value);
                if (left.Length == 0 || right.Length == 0 || !GluesWords(left, right)
                    || (EncodedRun.IsMatch(left) && EncodedRun.IsMatch(right)))
                {
                    continue;
                }

                var line = text.AsSpan(0, match.Index).Count('\n') + 1;
                offenders.Add($"{RepoFiles.Relative(file)}:{line}: …{Tail(left)}\" + \"{Head(right)}…");
            }
        }

        Assert.True(
            offenders.Count == 0,
            "These joined string literals glue two words together (add a space at the end of the left one):"
            + Environment.NewLine + string.Join(Environment.NewLine, offenders));
    }

    /// <summary>
    /// The shape of the bug: the left literal ends a word (a letter, digit, quote, closing parenthesis or
    /// sentence punctuation) and the right one starts one ("and" + "a", "error:" + "w:rPr",
    /// "level='full'" + "additionally"). A path, alternative or hyphenated word split at its '/' or '-', a list
    /// example split after a comma ("Description_Line," + "2:Quantity"), and a join at an interpolation hole
    /// are deliberate. After a comma only a letter counts, since prose wants a space there.
    /// </summary>
    private static bool GluesWords(string left, string right)
    {
        var last = left[^1];
        var first = right[0];
        if (last == ',')
        {
            return char.IsLetter(first);
        }

        var endsWord = char.IsLetterOrDigit(last) || last is '\'' or '"' or ')' or ':' or '.' or ';' or '!' or '?';
        var startsWord = char.IsLetterOrDigit(first) || first is '\'' or '(';
        return endsWord && startsWord;
    }

    private static string Body(string literal)
    {
        var start = literal.IndexOf('"', StringComparison.Ordinal) + 1;
        return literal[start..^1];
    }

    private static string Tail(string text) => text.Length <= 20 ? text : text[^20..];

    private static string Head(string text) => text.Length <= 20 ? text : text[..20];
}
