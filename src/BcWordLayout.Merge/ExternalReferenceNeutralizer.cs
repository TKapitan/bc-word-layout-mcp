using System.Collections.Frozen;
using System.Text;
using BcWordLayout.Domain;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Merge;

/// <summary>
/// The second half of the render-copy hardening (after <see cref="ExternalRelationshipStripper"/>): removes
/// every way a preview copy can point a converter at a resource OUTSIDE the copy that is not an OPC
/// relationship, so neither Word nor LibreOffice fetches a caller-chosen UNC path or URL while opening,
/// laying out or exporting it. Like the stripper it only ever runs on <c>preview_layout</c>'s merged working
/// copy, never on the caller's layout.
/// <list type="bullet">
/// <item><b>Fields.</b> <c>INCLUDEPICTURE</c> (a "web bug" picture Word fetches to render it),
/// <c>INCLUDETEXT</c>, <c>LINK</c>, <c>DDE</c>/<c>DDEAUTO</c>, <c>DATABASE</c> and their kin name a file,
/// URL or data source in the field INSTRUCTION, not in a relationship. Every field outside
/// <see cref="SelfContainedFieldKeywords"/> is UNLINKED — replaced by its last cached result, exactly what
/// Word's own Ctrl+Shift+F9 does — so the copy still looks the same but has nothing left to update. An
/// ALLOWLIST rather than a denylist, because the keyword itself can be computed: the DDE attacks of 2017 hid
/// <c>DDEAUTO</c> behind a nested <c>QUOTE</c> field. So a field is kept live only when its instruction BEGINS
/// with a literal, allowlisted keyword (see <see cref="LiteralKeyword"/>), with no tracked change inside its
/// instruction.</item>
/// <item><b>Mail-merge, update-on-open and schema-library settings.</b> <c>w:mailMerge</c> carries a
/// connection string, an ODSO <c>.udl</c> path and a query Word runs against the data source when it opens a
/// merge main document; <c>sl:schemaLibrary</c> carries schema/manifest locations. <c>w:updateFields</c> makes
/// Word update every field on open: observed with the Word COM converter, it raises Word's "fields that may
/// refer to other files" prompt on the desktop despite the hidden, alerts-off session, so a crafted layout
/// could both stall a preview on a modal dialog and get one "Yes" away from fetching. No BC layout uses any
/// of the three (BC binds content controls to its own custom XML part).</item>
/// <item><b>Imported content (<c>w:altChunk</c>).</b> The chunk's part is INTERNAL, but Word converts its
/// content (HTML, MHT, RTF, another .docx) on open, and that content can reference anything. The chunk and
/// its part are removed; an empty paragraph keeps the containing block valid.</item>
/// <item><b>VML URL attributes.</b> <c>src</c>, <c>o:href</c> and <c>o:althref</c> on VML elements hold a
/// location directly rather than through <c>r:id</c>; DOCX rendering never needs them.</item>
/// </list>
/// Each removal raises a warning (<c>field-unlinked</c> or <c>external-reference-stripped</c>), grouped per
/// part, so the caller knows the preview differs from the source layout in that respect.
/// </summary>
internal static class ExternalReferenceNeutralizer
{
    private const string VmlNamespace = "urn:schemas-microsoft-com:vml";
    private const string OfficeNamespace = "urn:schemas-microsoft-com:office:office";
    private const string SchemaLibraryNamespace = "http://schemas.openxmlformats.org/schemaLibrary/2006/main";

    /// <summary>Grouping key and message text for a field whose keyword is not a literal at all.</summary>
    private const string ComputedKeyword = "(computed-keyword)";

    /// <summary>Grouping key and message text for field markup inside markup-compatibility alternate content.</summary>
    private const string AlternateContentKeyword = "(alternate-content)";

    /// <summary>
    /// Field types computed only from the document itself — its text, bookmarks, styles, properties,
    /// pagination and the clock — which therefore have nothing outside the copy to load. <c>HYPERLINK</c>
    /// is included because following a link is a click, never part of opening or exporting (the same reason
    /// <see cref="ExternalRelationshipStripper"/> keeps hyperlink relationships). Everything else is unlinked,
    /// including fields that are harmless but not needed here (<c>MERGEFIELD</c>, <c>CITATION</c>), and the
    /// prompting ones (<c>ASK</c>, <c>FILLIN</c>) that would block a headless converter on a dialog.
    /// </summary>
    internal static readonly FrozenSet<string> SelfContainedFieldKeywords = new[]
    {
        "=", "ADVANCE", "AUTHOR", "AUTONUM", "AUTONUMLGL", "AUTONUMOUT", "BARCODE", "BIDIOUTLINE", "COMMENTS",
        "COMPARE", "CREATEDATE", "DATE", "DISPLAYBARCODE", "DOCPROPERTY", "DOCVARIABLE", "EDITTIME", "EQ",
        "FILENAME", "FILESIZE", "FORMCHECKBOX", "FORMDROPDOWN", "FORMTEXT", "HYPERLINK", "IF", "INDEX", "INFO",
        "KEYWORDS", "LASTSAVEDBY", "LISTNUM", "NOTEREF", "NUMCHARS", "NUMPAGES", "NUMWORDS", "PAGE", "PAGEREF",
        "PRINTDATE", "QUOTE", "REF", "REVNUM", "SAVEDATE", "SECTION", "SECTIONPAGES", "SEQ", "SET", "STYLEREF",
        "SUBJECT", "SYMBOL", "TA", "TC", "TEMPLATE", "TIME", "TITLE", "TOA", "TOC", "USERADDRESS",
        "USERINITIALS", "USERNAME", "XE",
    }.ToFrozenSet(StringComparer.Ordinal);

    /// <summary>
    /// Neutralizes every non-relationship external reference reachable from <paramref name="document"/>
    /// (see class remarks) and returns one warning per kind of removal per part.
    /// </summary>
    public static List<MergeWarning> Neutralize(WordprocessingDocument document)
    {
        var warnings = new List<MergeWarning>();

        foreach (var part in AllParts(document))
        {
            switch (part)
            {
                case DocumentSettingsPart settingsPart:
                    StripSettings(settingsPart, warnings);
                    break;
                case MainDocumentPart or HeaderPart or FooterPart or FootnotesPart or EndnotesPart
                    or WordprocessingCommentsPart or GlossaryDocumentPart:
                    if (part.RootElement is { } root)
                    {
                        NeutralizeContent(part, root, warnings);
                    }

                    break;
            }
        }

        return warnings;
    }

    /// <summary>
    /// Every part reachable from the package, each once — iterative (the package gate already bounded the
    /// count, see <see cref="ResourceLimits.MaxPackageParts"/>), and materialized up front because
    /// <see cref="RemoveAltChunks"/> deletes parts.
    /// </summary>
    private static List<OpenXmlPart> AllParts(OpenXmlPackage package)
    {
        var visited = new HashSet<OpenXmlPart>();
        var ordered = new List<OpenXmlPart>();
        var pending = new Stack<OpenXmlPart>(package.Parts.Select(p => p.OpenXmlPart).Reverse());
        while (pending.TryPop(out var part))
        {
            if (!visited.Add(part))
            {
                continue;
            }

            ordered.Add(part);
            foreach (var child in part.Parts.Reverse())
            {
                pending.Push(child.OpenXmlPart);
            }
        }

        return ordered;
    }

    private static void StripSettings(DocumentSettingsPart settingsPart, List<MergeWarning> warnings)
    {
        if (settingsPart.Settings is not { } settings)
        {
            return;
        }

        var partName = PartWalker.PartFileName(settingsPart);

        foreach (var mailMerge in settings.Elements<MailMerge>().ToList())
        {
            mailMerge.Remove();
            warnings.Add(Stripped(
                partName,
                "Removed the mail-merge settings (w:mailMerge) before rendering: they name a data source "
                + "(connection string, .udl file, query) the converter would query when it opens the copy. Business "
                + "Central layouts never use Word mail merge, so the preview is unaffected."));
        }

        foreach (var updateFields in settings.Elements<UpdateFieldsOnOpen>().ToList())
        {
            updateFields.Remove();
            warnings.Add(Stripped(
                partName,
                "Removed the update-fields-on-open setting (w:updateFields) before rendering: it makes Word stop "
                + "on a \"fields that may refer to other files\" prompt even in a hidden, alerts-off session, "
                + "where one \"Yes\" would update every field. The page-number fields a layout needs are "
                + "updated by the converter's own layout pass regardless."));
        }

        foreach (var library in settings.ChildElements
            .Where(e => e.LocalName == "schemaLibrary" && e.NamespaceUri == SchemaLibraryNamespace).ToList())
        {
            library.Remove();
            warnings.Add(Stripped(
                partName,
                "Removed the XML schema library (sl:schemaLibrary) before rendering: it names schema and "
                + "manifest locations outside the document. Business Central layouts never use one."));
        }
    }

    private static void NeutralizeContent(OpenXmlPart part, OpenXmlPartRootElement root, List<MergeWarning> warnings)
    {
        var partName = PartWalker.PartFileName(part);
        var unlinked = new SortedDictionary<string, int>(StringComparer.Ordinal);

        // A text box is its own story: a field can never span into or out of one, so each is read
        // separately and nested text boxes are skipped by their enclosing story (see StoryElements).
        foreach (var story in root.Descendants<TextBoxContent>().Cast<OpenXmlElement>().Prepend(root).ToList())
        {
            var elements = StoryElements(story);
            UnlinkComplexFields(elements, unlinked);
            UnlinkSimpleFields(elements, unlinked);
        }

        foreach (var (keyword, count) in unlinked)
        {
            var what = keyword switch
            {
                ComputedKeyword => $"{count} field(s) whose type is computed rather than written literally (or whose instruction carries a tracked change)",
                AlternateContentKeyword => $"{count} field(s) inside markup-compatibility alternate content (the converter picks a branch at open time, so the instruction cannot be read reliably)",
                _ => $"{count} {keyword} field(s)",
            };
            warnings.Add(new MergeWarning
            {
                Kind = "field-unlinked",
                Message = $"Unlinked {what} before rendering: each now shows its last cached result. The preview "
                    + "copy keeps live only fields computed from the document itself (PAGE, NUMPAGES, DATE, REF, "
                    + "IF, ...), so a field that loads a file, picture, link or data source (INCLUDEPICTURE, "
                    + "INCLUDETEXT, LINK, DDE, ...) can never make the converter fetch it.",
                Location = partName,
            });
        }

        var altChunks = RemoveAltChunks(part, root);
        if (altChunks > 0)
        {
            warnings.Add(Stripped(
                partName,
                $"Removed {altChunks} imported content chunk(s) (w:altChunk) before rendering: the converter "
                + "imports their HTML/RTF/document content on open, and that content can reference external "
                + "files or URLs. The preview does not show that content."));
        }

        var vmlAttributes = StripVmlUrlAttributes(root);
        if (vmlAttributes > 0)
        {
            warnings.Add(Stripped(
                partName,
                $"Removed {vmlAttributes} VML location attribute(s) (src, o:href, o:althref) before rendering: "
                + "they name an image location directly rather than through a relationship."));
        }
    }

    // ---- fields ----

    /// <summary>
    /// The elements of one story in document order, each flagged with whether it sits inside
    /// markup-compatibility alternate content WITHIN this story. Nested text boxes are not descended into:
    /// they are stories of their own. Iterative, so a deep tree cannot overflow the stack here.
    /// </summary>
    private static List<(OpenXmlElement Element, bool InAlternateContent)> StoryElements(OpenXmlElement story)
    {
        var result = new List<(OpenXmlElement, bool)>();
        var pending = new Stack<(OpenXmlElement Element, bool InAlternateContent)>();
        PushChildren(story, inAlternateContent: false);

        while (pending.TryPop(out var item))
        {
            result.Add(item);
            if (item.Element is not TextBoxContent)
            {
                PushChildren(item.Element, item.InAlternateContent || item.Element is AlternateContent);
            }
        }

        return result;

        void PushChildren(OpenXmlElement parent, bool inAlternateContent)
        {
            for (var child = parent.LastChild; child is not null; child = child.PreviousSibling())
            {
                pending.Push((child, inAlternateContent));
            }
        }
    }

    /// <summary>One complex field (<c>w:fldChar</c> begin … separate … end) being read.</summary>
    private sealed class FieldFrame(FieldChar begin)
    {
        public FieldChar Begin { get; } = begin;

        public FieldChar? Separate { get; set; }

        public FieldChar? End { get; set; }

        /// <summary>Instruction text up to the first nested field — the only part that can name the type.</summary>
        public StringBuilder Literal { get; } = new();

        /// <summary>A nested field began inside the instruction: <see cref="Literal"/> is complete.</summary>
        public bool LiteralClosed { get; set; }

        public bool TrackedChange { get; set; }

        /// <summary>
        /// Every run-level element between <see cref="Begin"/> and <see cref="Separate"/> (or <see cref="End"/>
        /// when there is no separator): the field CODE, nested fields included. What unlinking removes.
        /// </summary>
        public List<OpenXmlElement> Code { get; } = [];

        public bool InCode => Separate is null;
    }

    /// <summary>
    /// Reads the complex fields of one story in a single pass and unlinks every one that does not qualify
    /// (see class remarks): its begin/separate/end markers and its whole code are removed, its result runs
    /// stay where they are. Fields nest, so a field's code can contain whole fields; each element of a code
    /// belongs to EVERY enclosing field still reading its code. Inner fields are decided first; an unlinked
    /// outer field removes its inner fields along with the rest of its code.
    /// </summary>
    private static void UnlinkComplexFields(
        List<(OpenXmlElement Element, bool InAlternateContent)> elements, SortedDictionary<string, int> unlinked)
    {
        var open = new List<FieldFrame>();
        var alternateContentFields = 0;
        var alternateContentMarkup = false;

        foreach (var (element, inAlternateContent) in elements)
        {
            if (element is not (FieldChar or FieldCode or DeletedFieldCode)
                && (element.Parent is not Run || element is RunProperties))
            {
                continue;
            }

            if (inAlternateContent)
            {
                // Removed outright, in every branch: the linear order read here is not the branch the
                // converter will pick, so no instruction inside can be trusted (see class remarks).
                if (element is FieldChar or FieldCode or DeletedFieldCode)
                {
                    alternateContentMarkup = true;
                    if (element is FieldChar marker && marker.FieldCharType?.Value == FieldCharValues.Begin)
                    {
                        alternateContentFields++;
                    }

                    RemoveRunContent(element);
                }

                continue;
            }

            if (element is FieldChar fieldChar)
            {
                var type = fieldChar.FieldCharType?.Value;
                if (type == FieldCharValues.Begin)
                {
                    AddToOpenCodes(open, element);
                    if (open.Count > 0 && open[^1].InCode)
                    {
                        open[^1].LiteralClosed = true;
                    }

                    if (open.Count >= ResourceLimits.MaxFieldNestingDepth)
                    {
                        throw ResourceLimits.DepthExceeded("Field", ResourceLimits.MaxFieldNestingDepth);
                    }

                    open.Add(new FieldFrame(fieldChar));
                    continue;
                }

                if (type == FieldCharValues.Separate && open.Count > 0 && open[^1].InCode)
                {
                    open[^1].Separate = fieldChar;
                    AddToOpenCodes(open, element);
                    continue;
                }

                if (type == FieldCharValues.End && open.Count > 0)
                {
                    var closed = open[^1];
                    open.RemoveAt(open.Count - 1);
                    closed.End = fieldChar;
                    AddToOpenCodes(open, element);
                    Decide(closed, unlinked);
                    continue;
                }

                // A stray separator or end belongs to no field: just content of any enclosing code.
            }

            if (open.Count > 0 && open[^1].InCode && element is FieldCode or DeletedFieldCode)
            {
                var current = open[^1];
                if (element is DeletedFieldCode || element.Ancestors().Any(IsTrackedChange))
                {
                    current.TrackedChange = true;
                }

                if (element is FieldCode code && !current.LiteralClosed)
                {
                    current.Literal.Append(code.Text);
                }
            }

            AddToOpenCodes(open, element);
        }

        // Fields never terminated within the story: unlinked, since there is no complete instruction to judge.
        for (var i = open.Count - 1; i >= 0; i--)
        {
            Decide(open[i], unlinked);
        }

        if (alternateContentMarkup)
        {
            Count(unlinked, AlternateContentKeyword, Math.Max(1, alternateContentFields));
        }
    }

    private static void AddToOpenCodes(List<FieldFrame> open, OpenXmlElement element)
    {
        foreach (var frame in open)
        {
            if (frame.InCode)
            {
                frame.Code.Add(element);
            }
        }
    }

    private static void Decide(FieldFrame field, SortedDictionary<string, int> unlinked)
    {
        var keyword = field.TrackedChange ? null : LiteralKeyword(field.Literal.ToString(), field.LiteralClosed);
        if (field.End is not null && keyword is not null && SelfContainedFieldKeywords.Contains(keyword))
        {
            return;
        }

        RemoveRunContent(field.Begin);
        foreach (var element in field.Code)
        {
            RemoveRunContent(element);
        }

        if (field.Separate is not null)
        {
            RemoveRunContent(field.Separate);
        }

        if (field.End is not null)
        {
            RemoveRunContent(field.End);
        }

        Count(unlinked, keyword ?? ComputedKeyword, 1);
    }

    /// <summary>
    /// Unwraps every <c>w:fldSimple</c> that does not qualify, leaving its result content in its place.
    /// A simple field's instruction is one attribute (no nested fields), so only its literal keyword matters
    /// — unless it sits inside alternate content, where it is unwrapped regardless.
    /// </summary>
    private static void UnlinkSimpleFields(
        List<(OpenXmlElement Element, bool InAlternateContent)> elements, SortedDictionary<string, int> unlinked)
    {
        foreach (var (element, inAlternateContent) in elements)
        {
            if (element is not SimpleField field || field.Parent is null)
            {
                continue;
            }

            var keyword = inAlternateContent ? null : LiteralKeyword(field.Instruction?.Value ?? string.Empty, false);
            if (keyword is not null && SelfContainedFieldKeywords.Contains(keyword))
            {
                continue;
            }

            foreach (var child in field.ChildElements.ToList())
            {
                child.Remove();
                if (child is not FieldData)
                {
                    field.InsertBeforeSelf(child);
                }
            }

            field.Remove();
            Count(unlinked, inAlternateContent ? AlternateContentKeyword : keyword ?? ComputedKeyword, 1);
        }
    }

    /// <summary>
    /// The field type an instruction names, upper-cased, or null when it does not name one LITERALLY: the
    /// leading token must be ASCII letters (or the formula sign <c>=</c>) and must end at whitespace, a
    /// switch (<c>\</c>), a quote or the end of the instruction. When the instruction continues into a
    /// nested field (<paramref name="continuesIntoNestedField"/>) right where the token ends, the real type is
    /// only known once that field is evaluated (<c>INCLUDE</c> + <c>{ QUOTE "PICTURE" }</c>), so it is null.
    /// Any other character (a Cyrillic look-alike, a zero-width space) also makes it null. Null means unlink.
    /// </summary>
    internal static string? LiteralKeyword(string instruction, bool continuesIntoNestedField)
    {
        var text = instruction.TrimStart();
        if (text.Length == 0)
        {
            return null;
        }

        if (text[0] == '=')
        {
            return "=";
        }

        var end = 0;
        while (end < text.Length && char.IsAsciiLetter(text[end]))
        {
            end++;
        }

        if (end == 0)
        {
            return null;
        }

        if (end == text.Length)
        {
            return continuesIntoNestedField ? null : text.ToUpperInvariant();
        }

        var next = text[end];
        return char.IsWhiteSpace(next) || next is '\\' or '"' ? text[..end].ToUpperInvariant() : null;
    }

    private static bool IsTrackedChange(OpenXmlElement element) =>
        element is DeletedRun or InsertedRun or MoveFromRun or MoveToRun;

    /// <summary>Removes a run-level element, and its run when nothing but run properties is left.</summary>
    private static void RemoveRunContent(OpenXmlElement element)
    {
        if (element.Parent is not { } parent)
        {
            return;
        }

        element.Remove();
        if (parent is Run run && run.ChildElements.All(c => c is RunProperties))
        {
            run.Remove();
        }
    }

    private static void Count(SortedDictionary<string, int> counts, string key, int by) =>
        counts[key] = counts.GetValueOrDefault(key) + by;

    // ---- imported content and VML locations ----

    private static int RemoveAltChunks(OpenXmlPart part, OpenXmlElement root)
    {
        var removed = 0;
        foreach (var chunk in root.Descendants<AltChunk>().ToList())
        {
            if (chunk.Id?.Value is { } id && part.TryGetPartById(id, out var chunkPart))
            {
                part.DeletePart(chunkPart);
            }

            // An empty paragraph rather than nothing: a table cell (among others) must keep one block.
            chunk.InsertBeforeSelf(new Paragraph());
            chunk.Remove();
            removed++;
        }

        return removed;
    }

    private static int StripVmlUrlAttributes(OpenXmlElement root)
    {
        var removed = 0;
        foreach (var element in root.Descendants()
            .Where(e => e.NamespaceUri is VmlNamespace or OfficeNamespace).ToList())
        {
            foreach (var attribute in element.GetAttributes()
                .Where(a => (a.LocalName == "src" && a.NamespaceUri.Length == 0)
                    || (a.NamespaceUri == OfficeNamespace && a.LocalName is "href" or "althref"))
                .ToList())
            {
                element.RemoveAttribute(attribute.LocalName, attribute.NamespaceUri);
                removed++;
            }
        }

        return removed;
    }

    private static MergeWarning Stripped(string location, string message) => new()
    {
        Kind = "external-reference-stripped",
        Message = message,
        Location = location,
    };
}
