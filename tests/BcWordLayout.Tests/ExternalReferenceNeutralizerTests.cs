using BcWordLayout.Domain;
using BcWordLayout.Merge;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// Covers <see cref="ExternalReferenceNeutralizer"/>, the half of <see cref="MergeOptions.StripExternalReferences"/>
/// that removes references to outside resources which are NOT OPC relationships (fields, mail-merge and
/// schema-library settings, <c>w:altChunk</c>, VML location attributes). Same in-memory merge pattern as
/// <see cref="ExternalRelationshipStripperTests"/>; nothing here contacts the network — the assertions are on
/// the OOXML the converter would receive.
/// </summary>
public class ExternalReferenceNeutralizerTests
{
    private const string Namespaces =
        "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
        + "xmlns:v=\"urn:schemas-microsoft-com:vml\" "
        + "xmlns:o=\"urn:schemas-microsoft-com:office:office\" "
        + "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" "
        + "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" "
        + "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";

    private const string Pixel = "INCLUDEPICTURE \"http://attacker.example/pixel.png\" \\d";

    private const string Begin = "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>";
    private const string Separate = "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>";
    private const string End = "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>";

    private static string Instr(string text) => $"<w:r><w:instrText xml:space=\"preserve\">{text}</w:instrText></w:r>";

    private static string Text(string text) => $"<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r>";

    private static string Field(string instruction, string cachedResult) =>
        Begin + Instr(instruction) + Separate + Text(cachedResult) + End;

    private static Paragraph Para(string content) => new($"<w:p {Namespaces}>{content}</w:p>");

    /// <summary>
    /// Merges <paramref name="paragraphs"/> (prepended to a minimal synthetic layout's body) in memory with
    /// <see cref="MergeOptions.StripExternalReferences"/> on, then hands the open document to
    /// <paramref name="assert"/>.
    /// </summary>
    private static void MergeAndAssert(
        IEnumerable<Paragraph> paragraphs,
        Action<WordprocessingDocument, MergeResult> assert,
        Action<WordprocessingDocument>? arrange = null)
    {
        var body = SyntheticLayout.BoundField(
            "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:CompanyName[1]", SyntheticLayout.GoodItemId);
        var layoutPath = SyntheticLayout.Create(body);

        try
        {
            var dataset = SampleDataGenerator.Generate(
                SchemaProvider.FromLayout(layoutPath), new SampleDataOptions { Seed = 1, Rows = 1 });

            // A block, not a using declaration: the document must be closed before the finally deletes it, or
            // a failed assertion surfaces as a file-in-use error instead of itself.
            using (var doc = WordprocessingDocument.Open(layoutPath, true))
            {
                var documentBody = doc.MainDocumentPart!.Document!.Body!;
                foreach (var paragraph in paragraphs.Reverse())
                {
                    documentBody.PrependChild(paragraph);
                }

                arrange?.Invoke(doc);
                doc.MainDocumentPart.Document.Save();

                var result = MergeEngine.Merge(doc, dataset, new MergeOptions { StripExternalReferences = true });
                assert(doc, result);
            }
        }
        finally
        {
            File.Delete(layoutPath);
        }
    }

    private static string InstructionText(OpenXmlElement root) =>
        string.Concat(root.Descendants<FieldCode>().Select(c => c.Text));

    private static void AssertValid(WordprocessingDocument doc)
    {
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
        Assert.True(errors.Count == 0,
            "expected zero validation errors; found: "
            + string.Join(" | ", errors.Select(e => $"{e.Path?.XPath}: {e.Description}")));
    }

    // ---- complex fields ----

    [Fact]
    public void An_INCLUDEPICTURE_field_is_unlinked_to_its_cached_result_and_reported()
    {
        MergeAndAssert([Para(Text("before ") + Field(Pixel, "cached-picture") + Text(" after"))], (doc, result) =>
        {
            var body = doc.MainDocumentPart!.Document!.Body!;

            Assert.DoesNotContain("attacker.example", body.OuterXml, StringComparison.Ordinal);
            Assert.Empty(body.Descendants<FieldChar>());
            Assert.Contains("before cached-picture after", body.InnerText, StringComparison.Ordinal);

            var warning = Assert.Single(result.Warnings, w => w.Kind == "field-unlinked");
            Assert.Contains("1 INCLUDEPICTURE field(s)", warning.Message, StringComparison.Ordinal);
            Assert.Equal("document.xml", warning.Location);
            AssertValid(doc);
        });
    }

    [Fact]
    public void Self_contained_fields_stay_live_untouched_including_an_IF_around_PAGE_and_NUMPAGES()
    {
        var pageOfPages = Begin + Instr("IF ") + Field("PAGE", "1") + Instr(" &lt; ") + Field("NUMPAGES", "2")
            + Instr(" \"Continued\" \"\"") + Separate + Text("Continued") + End;
        var paragraphs = new[]
        {
            Para(Field("PAGE \\* MERGEFORMAT", "1") + Text(" / ") + Field(" numpages ", "3")),
            Para(Field("DATE \\@ \"d MMMM yyyy\"", "1 January 2026")),
            Para(pageOfPages),
            Para("<w:fldSimple w:instr=\" SECTIONPAGES \">" + Text("4") + "</w:fldSimple>"),
        };

        MergeAndAssert(paragraphs, (doc, result) =>
        {
            Assert.DoesNotContain(result.Warnings, w => w.Kind == "field-unlinked");

            var body = doc.MainDocumentPart!.Document!.Body!;
            Assert.Equal(18, body.Descendants<FieldChar>().Count()); // 6 complex fields x 3 markers, nothing removed
            Assert.Single(body.Descendants<SimpleField>());
            var instructions = InstructionText(body);
            Assert.Contains("IF ", instructions, StringComparison.Ordinal);
            Assert.Contains("NUMPAGES", instructions, StringComparison.Ordinal);
        });
    }

    [Fact]
    public void A_keyword_hidden_behind_a_nested_QUOTE_field_is_unlinked()
    {
        // The 2017 DDE-attack shape: the field TYPE comes from a nested field, so no literal keyword is
        // there to judge - the outer field is unlinked, nested QUOTE and all.
        var obfuscated = Begin + Field("QUOTE \"DDEAUTO\"", "DDEAUTO") + Instr(" c:\\\\windows\\\\system32\\\\cmd.exe \"/k calc\"")
            + Separate + Text("cached") + End;

        MergeAndAssert([Para(obfuscated)], (doc, result) =>
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            Assert.Empty(body.Descendants<FieldChar>());
            Assert.DoesNotContain("cmd.exe", body.OuterXml, StringComparison.Ordinal);
            Assert.Equal("cached", body.Descendants<Paragraph>().First().InnerText);

            var warning = Assert.Single(result.Warnings, w => w.Kind == "field-unlinked");
            Assert.Contains("computed rather than written literally", warning.Message, StringComparison.Ordinal);
        });
    }

    [Fact]
    public void A_keyword_that_runs_straight_into_a_nested_field_is_unlinked()
    {
        // "INCLUDE" + { QUOTE "PICTURE" } evaluates to INCLUDEPICTURE; the literal prefix alone decides nothing.
        var spliced = Begin + Instr("INCLUDE") + Field("QUOTE \"PICTURE\"", "PICTURE") + Instr(" \"http://attacker.example/p.png\"")
            + Separate + Text("cached") + End;

        MergeAndAssert([Para(spliced)], (doc, _) =>
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            Assert.Empty(body.Descendants<FieldChar>());
            Assert.DoesNotContain("attacker.example", body.OuterXml, StringComparison.Ordinal);
        });
    }

    [Fact]
    public void A_dangerous_field_nested_in_a_live_IF_is_unlinked_while_the_IF_stays()
    {
        var nested = Begin + Instr("IF 1 = 1 \"") + Field("INCLUDETEXT \"\\\\\\\\attacker-host\\\\share\\\\x.docx\"", "included")
            + Instr("\" \"\"") + Separate + Text("included") + End;

        MergeAndAssert([Para(nested)], (doc, result) =>
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            Assert.DoesNotContain("attacker-host", body.OuterXml, StringComparison.Ordinal);
            Assert.StartsWith("IF 1 = 1", InstructionText(body), StringComparison.Ordinal);
            Assert.Equal(3, body.Descendants<FieldChar>().Count()); // only the IF's own begin/separate/end

            var warning = Assert.Single(result.Warnings, w => w.Kind == "field-unlinked");
            Assert.Contains("INCLUDETEXT", warning.Message, StringComparison.Ordinal);
        });
    }

    [Fact]
    public void A_tracked_change_inside_an_instruction_unlinks_the_field()
    {
        // A deleted "PAGE " in front: read naively the keyword is PAGE, but the instruction Word keeps is
        // the INCLUDEPICTURE that follows. Any tracked change in an instruction makes the field unsafe.
        var tracked = Begin
            + "<w:del w:id=\"1\" w:author=\"x\"><w:r><w:delInstrText xml:space=\"preserve\">PAGE </w:delInstrText></w:r></w:del>"
            + Instr(Pixel) + Separate + Text("cached") + End;

        MergeAndAssert([Para(tracked)], (doc, _) =>
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            Assert.Empty(body.Descendants<FieldChar>());
            Assert.DoesNotContain("attacker.example", body.OuterXml, StringComparison.Ordinal);
        });
    }

    [Fact]
    public void Field_markup_inside_alternate_content_is_removed_from_every_branch()
    {
        // Which branch the converter reads is decided at open time, so the linear order of both branches is
        // not the instruction it will see: here a Fallback "PAGE " would make the field read as PAGE.
        var branched = Begin
            + "<w:r><mc:AlternateContent><mc:Choice Requires=\"w14\"><w:instrText xml:space=\"preserve\">PAGE </w:instrText></mc:Choice>"
            + "<mc:Fallback><w:instrText xml:space=\"preserve\">PAGE </w:instrText></mc:Fallback></mc:AlternateContent></w:r>"
            + Instr(Pixel) + Separate + Text("cached") + End;

        MergeAndAssert([Para(branched)], (doc, result) =>
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            Assert.Empty(body.Descendants<FieldChar>());
            Assert.Empty(body.Descendants<FieldCode>());
            Assert.DoesNotContain("attacker.example", body.OuterXml, StringComparison.Ordinal);
            Assert.Contains(result.Warnings, w => w.Kind == "field-unlinked" && w.Message.Contains("alternate content", StringComparison.Ordinal));
            Assert.Contains(result.Warnings, w => w.Kind == "field-unlinked" && w.Message.Contains("INCLUDEPICTURE", StringComparison.Ordinal));
        });
    }

    [Fact]
    public void An_unterminated_field_is_unlinked()
    {
        MergeAndAssert([Para(Begin + Instr(Pixel) + Separate + Text("cached"))], (doc, _) =>
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            Assert.Empty(body.Descendants<FieldChar>());
            Assert.Empty(body.Descendants<FieldCode>());
            Assert.Contains("cached", body.InnerText, StringComparison.Ordinal);
        });
    }

    [Fact]
    public void A_field_inside_a_text_box_is_judged_on_its_own()
    {
        // A text box is a separate story: its fields are complete inside it, so PAGE there stays live while
        // the INCLUDEPICTURE next to it is unlinked.
        var textBox = "<w:r><w:pict><v:shape id=\"tb1\" style=\"width:100pt;height:20pt\"><v:textbox><w:txbxContent>"
            + $"<w:p>{Field("PAGE", "1")}</w:p><w:p>{Field(Pixel, "cached")}</w:p>"
            + "</w:txbxContent></v:textbox></v:shape></w:pict></w:r>";

        MergeAndAssert([Para(textBox)], (doc, result) =>
        {
            var textBoxContent = Assert.Single(doc.MainDocumentPart!.Document!.Body!.Descendants<TextBoxContent>());
            Assert.Equal("PAGE", InstructionText(textBoxContent));
            Assert.DoesNotContain("attacker.example", textBoxContent.OuterXml, StringComparison.Ordinal);
            Assert.Single(result.Warnings, w => w.Kind == "field-unlinked");
        });
    }

    [Fact]
    public void Fields_nested_beyond_the_supported_depth_fail_the_merge_as_a_resource_limit()
    {
        var depth = ResourceLimits.MaxFieldNestingDepth + 1;
        var deep = string.Concat(Enumerable.Repeat(Begin + Instr("IF 1 = 1 \""), depth))
            + string.Concat(Enumerable.Repeat(Instr("\"") + End, depth));

        MergeAndAssertThrows(Para(deep));

        static void MergeAndAssertThrows(Paragraph paragraph) =>
            Assert.Throws<ResourceLimitExceededException>(() => MergeAndAssert([paragraph], (_, _) => { }));
    }

    // ---- simple fields ----

    [Fact]
    public void A_simple_INCLUDEPICTURE_field_is_unwrapped_to_its_result()
    {
        var simple = $"<w:fldSimple w:instr=\"{Pixel.Replace("\"", "&quot;", StringComparison.Ordinal)}\">"
            + Text("cached") + "</w:fldSimple>";

        MergeAndAssert([Para(simple)], (doc, result) =>
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            Assert.Empty(body.Descendants<SimpleField>());
            Assert.Equal("cached", body.Descendants<Paragraph>().First().InnerText);
            Assert.Contains(result.Warnings, w => w.Kind == "field-unlinked" && w.Message.Contains("INCLUDEPICTURE", StringComparison.Ordinal));
            AssertValid(doc);
        });
    }

    // ---- settings, imported content, VML ----

    [Fact]
    public void Mail_merge_update_on_open_and_schema_library_settings_are_removed()
    {
        // w:updateFields is the one observed live: with the Word COM converter it put Word's "fields that may
        // refer to other files" prompt on the desktop despite the hidden, alerts-off session.
        MergeAndAssert([], (doc, result) =>
        {
            var settings = doc.MainDocumentPart!.DocumentSettingsPart!.Settings!;
            Assert.Empty(settings.Elements<MailMerge>());
            Assert.Empty(settings.Elements<UpdateFieldsOnOpen>());
            Assert.DoesNotContain("attacker-host", settings.OuterXml, StringComparison.Ordinal);
            Assert.Equal(3, result.Warnings.Count(w => w.Kind == "external-reference-stripped" && w.Location == "settings.xml"));
            AssertValid(doc);
        },
        arrange: doc =>
        {
            var settingsPart = doc.MainDocumentPart!.DocumentSettingsPart ?? doc.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new Settings(
                $"<w:settings {Namespaces} xmlns:sl=\"http://schemas.openxmlformats.org/schemaLibrary/2006/main\">"
                + "<w:updateFields w:val=\"true\"/>"
                + "<w:mailMerge><w:mainDocumentType w:val=\"formLetters\"/><w:linkToQuery/><w:dataType w:val=\"native\"/>"
                + "<w:connectString w:val=\"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\\\attacker-host\\share\\db.accdb\"/>"
                + "<w:query w:val=\"SELECT * FROM t\"/></w:mailMerge>"
                + "<sl:schemaLibrary><sl:schema sl:uri=\"urn:x\" sl:manifestLocation=\"\\\\attacker-host\\share\\m.xml\"/></sl:schemaLibrary>"
                + "</w:settings>");
            settingsPart.Settings.Save();
        });
    }

    [Fact]
    public void An_altChunk_and_its_imported_part_are_removed()
    {
        string? chunkId = null;
        MergeAndAssert([], (doc, result) =>
        {
            var main = doc.MainDocumentPart!;
            Assert.Empty(main.Document!.Body!.Descendants<AltChunk>());
            Assert.Empty(main.AlternativeFormatImportParts);
            Assert.False(main.TryGetPartById(chunkId!, out _));
            Assert.Single(result.Warnings, w => w.Kind == "external-reference-stripped" && w.Message.Contains("altChunk", StringComparison.Ordinal));
            AssertValid(doc);
        },
        arrange: doc =>
        {
            var main = doc.MainDocumentPart!;
            var chunk = main.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html);
            using (var stream = chunk.GetStream(FileMode.Create))
            using (var writer = new StreamWriter(stream))
            {
                writer.Write("<html><body><img src=\"http://attacker.example/pixel.png\"></body></html>");
            }

            chunkId = main.GetIdOfPart(chunk);
            main.Document!.Body!.PrependChild(new AltChunk { Id = chunkId });
        });
    }

    [Fact]
    public void VML_location_attributes_are_removed()
    {
        var picture = "<w:r><w:pict><v:shape id=\"s1\" style=\"width:10pt;height:10pt\">"
            + "<v:imagedata src=\"http://attacker.example/a.png\" o:href=\"\\\\attacker-host\\share\\b.png\" o:title=\"logo\"/>"
            + "</v:shape></w:pict></w:r>";

        MergeAndAssert([Para(picture)], (doc, result) =>
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            Assert.DoesNotContain("attacker", body.OuterXml, StringComparison.Ordinal);
            Assert.Contains("o:title=\"logo\"", body.OuterXml, StringComparison.Ordinal);
            Assert.Single(result.Warnings, w => w.Kind == "external-reference-stripped" && w.Message.Contains("2 VML", StringComparison.Ordinal));
        });
    }

    // ---- persistence, scope, corpus ----

    [Fact]
    public void A_footnote_field_is_unlinked_in_the_saved_preview_copy()
    {
        // Footnotes have no explicit Save in MergeIntoOpenCopy: the unlink reaches the file through AutoSave.
        var layoutPath = SyntheticLayout.Create(SyntheticLayout.PlainParagraph("body"));
        var outputDir = Directory.CreateTempSubdirectory("bcwl-neutralize-footnote-").FullName;

        try
        {
            using (var doc = WordprocessingDocument.Open(layoutPath, true))
            {
                var footnotes = doc.MainDocumentPart!.AddNewPart<FootnotesPart>();
                footnotes.Footnotes = new Footnotes(new Footnote(Para(Field(Pixel, "cached"))) { Id = 1 });
                footnotes.Footnotes.Save();
            }

            var outputPath = Path.Combine(outputDir, "merged.docx");
            var result = MergeEngine.Merge(layoutPath, outputPath, new MergeOptions { StripExternalReferences = true });

            Assert.Contains(result.Warnings, w => w.Kind == "field-unlinked" && w.Location == "footnotes.xml");
            using var merged = WordprocessingDocument.Open(outputPath, false);
            var savedFootnotes = merged.MainDocumentPart!.FootnotesPart!.Footnotes!;
            Assert.DoesNotContain("attacker.example", savedFootnotes.OuterXml, StringComparison.Ordinal);
            Assert.Contains("cached", savedFootnotes.InnerText, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(layoutPath);
            Directory.Delete(outputDir, recursive: true);
        }
    }

    [Fact]
    public void Nothing_is_neutralized_when_the_option_is_off()
    {
        var layoutPath = SyntheticLayout.Create(SyntheticLayout.PlainParagraph("body"));

        try
        {
            var dataset = SampleDataGenerator.Generate(
                SchemaProvider.FromLayout(layoutPath), new SampleDataOptions { Seed = 1, Rows = 1 });
            using var doc = WordprocessingDocument.Open(layoutPath, true);
            doc.MainDocumentPart!.Document!.Body!.PrependChild(Para(Field(Pixel, "cached")));

            var result = MergeEngine.Merge(doc, dataset, new MergeOptions());

            Assert.DoesNotContain(result.Warnings, w => w.Kind is "field-unlinked" or "external-reference-stripped");
            Assert.Contains("attacker.example", InstructionText(doc.MainDocumentPart.Document), StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(layoutPath);
        }
    }

    [Fact]
    public void No_corpus_layout_loses_a_field_or_setting_to_the_hardening()
    {
        // Fidelity guard: the allowlist covers every field the real layouts use (PAGE, NUMPAGES, DATE).
        var outputDir = Directory.CreateTempSubdirectory("bcwl-neutralize-corpus-").FullName;

        try
        {
            foreach (var path in Directory.GetFiles(Path.Combine(AppContext.BaseDirectory, "corpus"), "*.docx"))
            {
                var result = MergeEngine.Merge(
                    path,
                    Path.Combine(outputDir, Path.GetFileName(path)),
                    new MergeOptions { StripExternalReferences = true, FlattenBindingsForRender = true });

                Assert.DoesNotContain(result.Warnings, w => w.Kind is "field-unlinked" or "external-reference-stripped");
            }
        }
        finally
        {
            Directory.Delete(outputDir, recursive: true);
        }
    }

    // ---- keyword parsing ----

    [Theory]
    [InlineData("PAGE", false, "PAGE")]
    [InlineData("  page \\* MERGEFORMAT", false, "PAGE")]
    [InlineData("INCLUDEPICTURE \"x\" \\d", false, "INCLUDEPICTURE")]
    [InlineData("HYPERLINK\"https://example.com\"", false, "HYPERLINK")]
    [InlineData("= 2 * 3", false, "=")]
    [InlineData("IF ", true, "IF")]
    [InlineData("INCLUDE", true, null)]          // runs straight into a nested field
    [InlineData("   ", true, null)]              // the keyword IS the nested field
    [InlineData("", false, null)]
    [InlineData("PAG\u0415", false, null)]       // Cyrillic look-alike E
    [InlineData("INCLUDE\u200BPICTURE", false, null)] // zero-width space inside the keyword
    public void LiteralKeyword_reads_only_a_literal_delimited_keyword(string instruction, bool nested, string? expected) =>
        Assert.Equal(expected, ExternalReferenceNeutralizer.LiteralKeyword(instruction, nested));
}
