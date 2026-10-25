using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using BcWordLayout.Merge;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// Covers GitHub issue #31: <c>wrap_in_repeater</c> (<see cref="LayoutEditor.WrapInRepeater"/>) wraps body blocks in
/// a block-level repeater in the corpus's two shapes - StandardStatement's section › group › item over a top-level
/// item, the stock quote's section › item over a nested one - and <c>remove_control keepText=true</c> unwraps it.
/// </summary>
public class WrapInRepeaterTests
{
    private const string WorkDescriptionLine = "/Header/WorkDescriptionLines/WorkDescriptionLine";

    private static string NewLayout(string corpusSchema)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-wrap-{Guid.NewGuid():N}.docx");
        LayoutBuilder.Create(Corpus.Path(corpusSchema), path, headingText: "DOCUMENT");
        return path;
    }

    private static int Field(WordprocessingDocument doc, string datasetPath) =>
        LayoutEditor.InsertField(doc, datasetPath, new Location { Type = LocationKind.DocumentEnd }).ControlId;

    private static Body BodyOf(WordprocessingDocument doc) => doc.MainDocumentPart!.Document!.Body!;

    private static SdtBlock Only(SdtBlock sdt) => Assert.IsType<SdtBlock>(Assert.Single(sdt.SdtContentBlock!.ChildElements));

    private static void AssertValid(string path)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc));
        var full = FullValidator.Full(path);
        Assert.True(full.Passed, string.Join(" | ", full.Findings.Select(f => $"{f.Check}: {f.Message}")));
    }

    // ---- the two corpus shapes ----

    [Fact]
    public void A_nested_item_holds_its_lone_field_as_a_block_control_in_section_and_item_with_no_group()
    {
        var path = NewLayout(Corpus.StandardSalesQuote);
        try
        {
            EditResult result;
            int field;
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                field = Field(doc, WorkDescriptionLine);
                result = LayoutEditor.WrapInRepeater(doc, "/Header/WorkDescriptionLines", field, null);
                doc.MainDocumentPart!.Document!.Save();
            }

            Assert.Equal("wrap_in_repeater", result.Operation);
            Assert.Equal("Repeater", result.Kind);
            Assert.Equal("/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:WorkDescriptionLines", result.XPath);
            Assert.Contains("1 paragraph(s) holding nothing but a field became that field as a block-level control", result.Summary);

            using (var doc = WordprocessingDocument.Open(path, false))
            {
                var repeater = BodyOf(doc).Elements<SdtBlock>().Single(s => SdtInspector.IsRepeater(s));
                Assert.Equal(result.ControlId, SdtInspector.ReadControlId(repeater));
                var item = Only(repeater);
                Assert.True(SdtInspector.IsRepeaterItem(item)); // section > item: no group for a nested item

                // The stock quote's shape: item > the field as a block control > its paragraph, same id and binding.
                var control = Only(item);
                Assert.Equal(field, SdtInspector.ReadControlId(control));
                Assert.Equal("#Nav: " + WorkDescriptionLine, SdtInspector.ReadAlias(control));
                var paragraph = Assert.IsType<Paragraph>(Assert.Single(control.SdtContentBlock!.ChildElements));
                Assert.Empty(paragraph.Descendants<SdtElement>());
                Assert.Equal("WorkDescriptionLine", paragraph.InnerText);
            }

            AssertValid(path);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void The_lifted_field_has_the_stock_quotes_shape_and_a_paragraph_holding_more_stays_as_it_is()
    {
        var path = NewLayout(Corpus.StandardSalesQuote);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var lone = Field(doc, WorkDescriptionLine);
            var mixed = Field(doc, WorkDescriptionLine);
            LayoutEditor.InsertText(doc, "- ", new Location { Type = LocationKind.AfterControl, ControlId = mixed });

            var result = LayoutEditor.WrapInRepeater(doc, "/Header/WorkDescriptionLines", lone, mixed);
            Assert.Contains("1 paragraph(s) holding nothing but a field", result.Summary);

            var item = Only(BodyOf(doc).Elements<SdtBlock>().Single(s => SdtInspector.IsRepeater(s)));
            var content = item.SdtContentBlock!.ChildElements.ToList();
            Assert.Equal(2, content.Count);
            Assert.IsType<SdtBlock>(content[0]);
            var kept = Assert.IsType<Paragraph>(content[1]); // text beside the field: no block form keeps it
            Assert.Equal(mixed, SdtInspector.ReadControlId(Assert.Single(kept.Elements<SdtRun>())));

            // The lifted control is the stock quote's: block sdt > sdtContent > paragraph, its sdtPr the field's.
            using var stock = WordprocessingDocument.Open(Corpus.Path(Corpus.StandardSalesQuote), false);
            var stockItem = Only(BodyOf(stock).Descendants<SdtBlock>().Single(s =>
                SdtInspector.IsRepeater(s) && SdtInspector.ReadAlias(s) == "#Nav: /Header/WorkDescriptionLines"));
            var stockField = Only(stockItem);
            static string Shape(SdtBlock sdt) => string.Join(">", sdt.Descendants().Select(e => e.LocalName)
                .Where(n => n is "sdt" or "sdtContent" or "p" or "r" or "t" or "dataBinding" or "text"));
            Assert.Equal(Shape(stockField), Shape((SdtBlock)content[0]));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_top_level_item_gets_StandardStatements_contentLocked_group_and_page_breaks()
    {
        var path = NewLayout(Corpus.StandardStatement);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                Field(doc, "/Customer/No_Cust");
                Field(doc, "/Customer/Integer/CompanyAddr1");
                var result = LayoutEditor.WrapInRepeater(doc, "/Customer", null, null, wholeBody: true, pageBreakBefore: true);
                Assert.Contains("contentLocked w:group", result.Summary);
                doc.MainDocumentPart!.Document!.Save();
            }

            using (var doc = WordprocessingDocument.Open(path, false))
            {
                var body = BodyOf(doc);
                Assert.IsType<SectionProperties>(body.LastChild); // the page setup stays outside the repeat
                var repeater = Assert.IsType<SdtBlock>(Assert.Single(body.ChildElements, c => c is not SectionProperties));
                Assert.True(SdtInspector.IsRepeater(repeater));

                var group = Only(repeater);
                Assert.NotNull(group.SdtProperties!.GetFirstChild<SdtContentGroup>());
                Assert.Equal(LockingValues.ContentLocked, group.SdtProperties.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Lock>()?.Val?.Value);

                var item = Only(group);
                Assert.True(SdtInspector.IsRepeaterItem(item));
                var paragraphs = item.SdtContentBlock!.Elements<Paragraph>().ToList();
                Assert.Equal(3, paragraphs.Count); // heading + two fields, moved in order
                Assert.NotNull(paragraphs[0].ParagraphProperties?.PageBreakBefore); // on the first, not a new one
                Assert.Equal("DOCUMENT", paragraphs[0].InnerText);
            }

            AssertValid(path);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void A_leading_table_gets_a_page_break_paragraph_in_front_and_a_paragraph_after()
    {
        var path = NewLayout(Corpus.StandardSalesQuote);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                LayoutEditor.InsertPlainTable(doc, 1, TableArgs.DomainColumns(1), new Location { Type = LocationKind.DocumentEnd });
                var tableIndex = BodyOf(doc).Descendants<Table>().Count() - 1;
                var field = LayoutEditor.InsertField(doc, WorkDescriptionLine,
                    new Location { Type = LocationKind.TableCell, TableIndex = tableIndex, Row = 0, Col = 0 }).ControlId;

                LayoutEditor.WrapInRepeater(doc, "/Header/WorkDescriptionLines", field, null, pageBreakBefore: true);
                doc.MainDocumentPart!.Document!.Save();
            }

            using (var doc = WordprocessingDocument.Open(path, false))
            {
                var item = Only(BodyOf(doc).Elements<SdtBlock>().Single(s => SdtInspector.IsRepeater(s)));
                var content = item.SdtContentBlock!.ChildElements.ToList();
                Assert.Equal(3, content.Count);
                Assert.NotNull(Assert.IsType<Paragraph>(content[0]).ParagraphProperties?.PageBreakBefore);
                Assert.IsType<Table>(content[1]);
                Assert.IsType<Paragraph>(content[2]);
            }

            AssertValid(path);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void The_range_runs_from_the_first_controls_block_to_the_last_controls_block_and_no_further()
    {
        var path = NewLayout(Corpus.StandardSalesQuote);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var before = Field(doc, "/Header/DocumentNo");
            var first = Field(doc, WorkDescriptionLine);
            var middle = Field(doc, "/Header/DocumentDate");
            var last = Field(doc, WorkDescriptionLine);
            var after = Field(doc, "/Header/YourReference");

            var result = LayoutEditor.WrapInRepeater(doc, "/Header/WorkDescriptionLines", first, last);

            // Two under the item; the header field between them is allowed and reported.
            Assert.Contains("2 bound control(s) inside follow each record", result.Summary);
            Assert.Contains("1 bound outside '/Header/WorkDescriptionLines'", result.Summary);

            var item = Only(BodyOf(doc).Elements<SdtBlock>().Single(s => SdtInspector.IsRepeater(s)));
            var ids = item.Descendants<SdtElement>().Select(SdtInspector.ReadControlId).ToList();
            Assert.Equal(new int?[] { first, middle, last }, ids);

            var outside = BodyOf(doc).ChildElements.Where(c => c is Paragraph).SelectMany(p => p.Descendants<SdtElement>())
                .Select(SdtInspector.ReadControlId).ToList();
            Assert.Contains((int?)before, outside);
            Assert.Contains((int?)after, outside);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- refusals ----

    [Fact]
    public void Refuses_a_range_without_ids_or_wholeBody_ids_together_with_wholeBody_and_controls_out_of_order()
    {
        var path = NewLayout(Corpus.StandardSalesQuote);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var first = Field(doc, WorkDescriptionLine);
            var second = Field(doc, WorkDescriptionLine);

            Assert.Equal("firstControlId", Assert.Throws<ArgumentException>(
                () => LayoutEditor.WrapInRepeater(doc, "/Header/WorkDescriptionLines", null, null)).ParamName);
            Assert.Equal("wholeBody", Assert.Throws<ArgumentException>(
                () => LayoutEditor.WrapInRepeater(doc, "/Header/WorkDescriptionLines", first, null, wholeBody: true)).ParamName);
            Assert.Equal("lastControlId", Assert.Throws<ArgumentException>(
                () => LayoutEditor.WrapInRepeater(doc, "/Header/WorkDescriptionLines", second, first)).ParamName);
            Assert.Throws<NotFoundException>(() => LayoutEditor.WrapInRepeater(doc, "/Header/WorkDescriptionLines", 12345, null));
            Assert.DoesNotContain(BodyOf(doc).Descendants<SdtElement>(), s => SdtInspector.IsRepeater(s));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Refuses_a_control_in_a_header_naming_the_part()
    {
        var path = NewLayout(Corpus.StandardSalesQuote);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var headerField = LayoutEditor.InsertField(doc, WorkDescriptionLine,
                new Location { Type = LocationKind.DocumentEnd, Part = LayoutPart.Header }).ControlId;

            var ex = Assert.Throws<ArgumentException>(
                () => LayoutEditor.WrapInRepeater(doc, "/Header/WorkDescriptionLines", headerField, null));
            Assert.Equal("firstControlId", ex.ParamName);
            Assert.Contains("header", ex.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Refuses_a_range_with_nothing_bound_under_the_item()
    {
        // A header field in a work-description repeater would print the same document number per line.
        var path = NewLayout(Corpus.StandardSalesQuote);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var field = Field(doc, "/Header/DocumentNo");

            var ex = Assert.Throws<ArgumentException>(
                () => LayoutEditor.WrapInRepeater(doc, "/Header/WorkDescriptionLines", field, null));
            Assert.Equal("dataItemPath", ex.ParamName);
            Assert.Contains("Nothing in the range is bound under", ex.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Refuses_a_repeater_in_the_range_over_data_outside_the_item_or_over_the_item_itself()
    {
        var path = NewLayout(Corpus.StandardSalesQuote);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var field = Field(doc, WorkDescriptionLine);
            var lines = LayoutEditor.InsertRepeaterTable(doc, "/Header/Line", TableArgs.DomainFieldsOf("/Header/Line", "ItemNo_Line"),
                new Location { Type = LocationKind.DocumentEnd }, new RepeaterTableOptions()).ControlId;

            var foreign = Assert.Throws<ArgumentException>(
                () => LayoutEditor.WrapInRepeater(doc, "/Header/WorkDescriptionLines", field, lines));
            Assert.Equal("dataItemPath", foreign.ParamName);
            Assert.Contains("'/Header/Line'", foreign.Message, StringComparison.Ordinal);

            var wrapped = LayoutEditor.WrapInRepeater(doc, "/Header/WorkDescriptionLines", field, null).ControlId;
            var same = Assert.Throws<ArgumentException>(
                () => LayoutEditor.WrapInRepeater(doc, "/Header/WorkDescriptionLines", field, null));
            Assert.Contains($"already holds a repeater over '/Header/WorkDescriptionLines' (id {wrapped})", same.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Refuses_a_range_holding_a_section_break()
    {
        // JobQuote's body has paragraph-level section breaks between its sections.
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-wrap-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(Corpus.JobQuote), path);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            Assert.Contains(BodyOf(doc).Elements<Paragraph>(), p => p.ParagraphProperties?.SectionProperties is not null);

            var ex = Assert.Throws<ArgumentException>(
                () => LayoutEditor.WrapInRepeater(doc, "/Job", null, null, wholeBody: true));
            Assert.Equal("lastControlId", ex.ParamName);
            Assert.Contains("section break", ex.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Refuses_a_leaf_column_as_the_data_item()
    {
        var path = NewLayout(Corpus.StandardSalesQuote);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var field = Field(doc, WorkDescriptionLine);

            Assert.Equal("dataItemPath", Assert.Throws<ArgumentException>(
                () => LayoutEditor.WrapInRepeater(doc, WorkDescriptionLine, field, null)).ParamName);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- undo: remove_control keepText=true ----

    [Fact]
    public void RemoveControl_keepText_unwraps_a_wrapped_top_level_repeater_back_to_its_blocks()
    {
        var path = NewLayout(Corpus.StandardStatement);
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                Field(doc, "/Customer/No_Cust");
                var before = BodyOf(doc).ChildElements.Select(c => c.OuterXml).ToList();

                var repeater = LayoutEditor.WrapInRepeater(doc, "/Customer", null, null, wholeBody: true).ControlId;
                var removed = LayoutEditor.RemoveControl(doc, repeater, keepText: true);

                Assert.Contains("kept its content", removed.Summary);
                Assert.Equal(before, BodyOf(doc).ChildElements.Select(c => c.OuterXml).ToList());
                doc.MainDocumentPart!.Document!.Save();
            }

            AssertValid(path);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveControl_keepText_unwraps_the_stock_quotes_work_description_repeater()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-wrap-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(Corpus.StandardSalesQuote), path);
        try
        {
            using var doc = WordprocessingDocument.Open(path, true);
            var body = BodyOf(doc);
            var repeater = body.Descendants<SdtBlock>().Single(s =>
                SdtInspector.IsRepeater(s) && SdtInspector.ReadAlias(s) == "#Nav: /Header/WorkDescriptionLines");
            var parent = repeater.Parent!;
            var field = repeater.Descendants<SdtElement>().Single(s => SdtInspector.ReadAlias(s) == "#Nav: " + WorkDescriptionLine);

            LayoutEditor.RemoveControl(doc, SdtInspector.ReadControlId(repeater)!.Value, keepText: true);

            Assert.DoesNotContain(body.Descendants<SdtBlock>(), s => SdtInspector.ReadAlias(s) == "#Nav: /Header/WorkDescriptionLines");
            Assert.Same(parent, field.Parent); // the item's own content, now where the repeater was
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- the mock merge expands what the tool wraps ----

    [Fact]
    public void The_mock_merge_prints_a_wrapped_paragraph_and_a_wrapped_body_once_per_record()
    {
        var nested = NewLayout(Corpus.StandardSalesQuote);
        var statement = NewLayout(Corpus.StandardStatement);
        var nestedOut = Path.Combine(Path.GetTempPath(), $"bcwl-wrap-out-{Guid.NewGuid():N}.docx");
        var statementOut = Path.Combine(Path.GetTempPath(), $"bcwl-wrap-out-{Guid.NewGuid():N}.docx");
        try
        {
            using (var doc = WordprocessingDocument.Open(nested, true))
            {
                LayoutEditor.WrapInRepeater(doc, "/Header/WorkDescriptionLines", Field(doc, WorkDescriptionLine), null);
                doc.MainDocumentPart!.Document!.Save();
            }

            using (var doc = WordprocessingDocument.Open(statement, true))
            {
                Field(doc, "/Customer/No_Cust");
                LayoutEditor.WrapInRepeater(doc, "/Customer", null, null, wholeBody: true, pageBreakBefore: true);
                doc.MainDocumentPart!.Document!.Save();
            }

            var nestedResult = MergeEngine.Merge(nested, nestedOut, new MergeOptions { Rows = 3 });
            Assert.Equal(1, nestedResult.Stats.RepeatersExpanded);
            Assert.Equal(3, nestedResult.Stats.RowsGenerated);

            var statementResult = MergeEngine.Merge(statement, statementOut, new MergeOptions { Rows = 3 });
            Assert.Equal(3, statementResult.Stats.RowsGenerated);
            using var merged = WordprocessingDocument.Open(statementOut, false);
            Assert.Equal(3, BodyOf(merged).Descendants<PageBreakBefore>().Count()); // one page per customer
        }
        finally
        {
            foreach (var p in new[] { nested, statement, nestedOut, statementOut })
            {
                File.Delete(p);
            }
        }
    }

    // ---- the tool ----

    [Fact]
    public void The_tool_wraps_validates_and_saves()
    {
        var path = NewLayout(Corpus.StandardSalesQuote);
        try
        {
            int field;
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                field = Field(doc, WorkDescriptionLine);
                doc.MainDocumentPart!.Document!.Save();
            }

            var response = EditTools.WrapInRepeater(path, "/Header/WorkDescriptionLines", firstControlId: field);

            Assert.True(response.Ok, response.Error?.Message);
            var data = Assert.IsType<EditResultDto>(response.Data);
            Assert.Equal("Repeater", data.Kind);
            Assert.True(data.QuickValidation.Passed);
            using var reopened = WordprocessingDocument.Open(path, false);
            Assert.Single(BodyOf(reopened).Elements<SdtBlock>(), s => SdtInspector.IsRepeater(s));
        }
        finally
        {
            File.Delete(path);
        }
    }
}
