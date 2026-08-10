using System.Xml.Linq;
using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// Covers <see cref="LayoutBuilder.Create"/> directly (no MCP tool layer): a layout it creates must pass
/// <see cref="OpenXmlValidator"/> and <see cref="LayoutValidator.Quick"/>, be readable by
/// <see cref="SchemaProvider"/>/<see cref="LayoutReader"/>, carry exactly one BC custom XML part plus the
/// glossary part <see cref="SdtFactory"/>'s placeholders reference, and — the key proof — be immediately
/// editable by <see cref="LayoutEditor"/>.
/// </summary>
public class LayoutBuilderTests
{
    // Word's own built-in "Click here to enter text." glossary docPart id. SdtFactory.DefaultPlaceholderDocPart
    // is internal but IS visible here (Domain grants InternalsVisibleTo to Tests as of B12(b)); the literal is
    // still deliberately duplicated so this test FAILS if the factory's constant ever drifts from Word's
    // well-known id - referencing the constant would make the assertion tautological.
    private const string DefaultPlaceholderDocPart = "DefaultPlaceholder_-1854013440";

    private static string TempOutputPath() =>
        Path.Combine(Path.GetTempPath(), $"bcwl-created-{Guid.NewGuid():N}.docx");

    /// <summary>
    /// Finds every custom XML part whose root namespace starts with the BC prefix, via PUBLIC API only.
    /// <c>SchemaProvider.FindBcParts</c> is internal and HAS been visible here since Domain granted
    /// InternalsVisibleTo to Tests (B12(b)) - this helper deliberately sticks to the public surface anyway,
    /// so it keeps proving the part is discoverable the way a real consumer would discover it.
    /// </summary>
    private static List<CustomXmlPart> FindBcCustomXmlParts(MainDocumentPart main)
    {
        var result = new List<CustomXmlPart>();
        foreach (var part in main.CustomXmlParts)
        {
            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            XElement? root;
            try
            {
                root = XDocument.Load(stream).Root;
            }
            catch
            {
                continue;
            }

            if (root is not null && root.Name.NamespaceName.StartsWith(OoxmlNames.BcNamespacePrefix, StringComparison.Ordinal))
            {
                result.Add(part);
            }
        }

        return result;
    }

    private static void AssertNoOpenXmlErrors(WordprocessingDocument doc)
    {
        var errors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc).ToList();
        Assert.True(errors.Count == 0, "expected 0 OpenXmlValidator errors; got: "
            + string.Join(" | ", errors.Take(10).Select(e => $"{e.Path?.XPath}: {e.Description}")));
    }

    private static void AssertQuickPasses(WordprocessingDocument doc)
    {
        var quick = LayoutValidator.Quick(doc);
        Assert.True(quick.Passed, "expected pass; errors: " + string.Join(" | ",
            quick.Findings.Where(f => f.Severity == FindingSeverity.Error).Select(f => f.Message)));
        Assert.Equal(0, quick.ErrorCount);
    }

    [Fact]
    public void Created_blank_layout_carries_the_BC_standard_A4_page_setup()
    {
        // Every corpus layout is A4 with 567-twip margins and a 1134-twip left margin - an exact
        // 10206-twip content width. Without this a blank layout gets Word's defaults and every
        // full-width BC-style block authored into it (insert_table's default width, the corpus-shaped
        // widths agents copy) overhangs the right margin.
        var outputPath = TempOutputPath();
        try
        {
            LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), outputPath);

            using var doc = WordprocessingDocument.Open(outputPath, false);
            var sectPr = doc.MainDocumentPart!.Document!.Body!.Elements<SectionProperties>().Single();
            var pageSize = sectPr.GetFirstChild<PageSize>()!;
            Assert.Equal(11907u, pageSize.Width!.Value);
            Assert.Equal(16839u, pageSize.Height!.Value);
            var margins = sectPr.GetFirstChild<PageMargin>()!;
            Assert.Equal(10206, (int)(pageSize.Width.Value - margins.Left!.Value - margins.Right!.Value));
        }
        finally
        {
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void Created_blank_layout_ships_empty_header_and_footer_parts_wired_into_the_page_setup()
    {
        // Every corpus layout has at least one header and one footer part; a blank build used to have
        // neither, so insert_field/insert_label layoutPart='footer' could not resolve at all and a
        // from-scratch layout had no way to author the per-page legal/contact block real BC documents put
        // there. A part that exists but is referenced by no section renders on no page, so
        // the w:headerReference/w:footerReference wiring is as much the deliverable as the parts are.
        var outputPath = TempOutputPath();
        try
        {
            LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), outputPath);

            using var doc = WordprocessingDocument.Open(outputPath, false);
            var main = doc.MainDocumentPart!;
            var headerPart = Assert.Single(main.HeaderParts);
            var footerPart = Assert.Single(main.FooterParts);

            // Scaffolded, not populated: one empty paragraph each (a w:hdr/w:ftr needs one block child).
            Assert.Equal(string.Empty, headerPart.Header!.InnerText);
            Assert.Equal(string.Empty, footerPart.Footer!.InnerText);

            var sectPr = main.Document!.Body!.Elements<SectionProperties>().Single();
            Assert.Equal(main.GetIdOfPart(headerPart), sectPr.Elements<HeaderReference>().Single().Id!.Value);
            Assert.Equal(main.GetIdOfPart(footerPart), sectPr.Elements<FooterReference>().Single().Id!.Value);

            AssertNoOpenXmlErrors(doc);
            AssertQuickPasses(doc);
        }
        finally
        {
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void Created_blank_layout_ships_a_default_styles_part_pinning_Calibri_and_defining_TableGrid()
    {
        // Without a styles part nothing in a from-scratch layout names a typeface, so Word and BC each
        // render their own application default and disagree (observed against a real BC sandbox,
        // 2026-08-01 - GitHub issue #3). The scaffold pins the typography every stock corpus layout
        // resolves to (Calibri 11pt) explicitly in docDefaults, and defines TableGrid so
        // insert_repeater_table's documented tableStyle example resolves instead of referencing nothing.
        var outputPath = TempOutputPath();
        try
        {
            LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), outputPath);

            using var doc = WordprocessingDocument.Open(outputPath, false);
            var stylesPart = doc.MainDocumentPart!.StyleDefinitionsPart;
            Assert.NotNull(stylesPart);

            var runDefaults = stylesPart!.Styles!.DocDefaults!.RunPropertiesDefault!.RunPropertiesBaseStyle!;
            Assert.Equal("Calibri", runDefaults.GetFirstChild<RunFonts>()!.Ascii!.Value);
            Assert.Equal("Calibri", runDefaults.GetFirstChild<RunFonts>()!.HighAnsi!.Value);
            Assert.Equal("22", runDefaults.GetFirstChild<FontSize>()!.Val!.Value);

            var styles = stylesPart.Styles.Elements<Style>().ToList();
            var normal = Assert.Single(styles, s => s.StyleId?.Value == "Normal");
            Assert.Equal(StyleValues.Paragraph, normal.Type!.Value);
            Assert.True(normal.Default!.Value);
            var tableGrid = Assert.Single(styles, s => s.StyleId?.Value == "TableGrid");
            Assert.Equal(StyleValues.Table, tableGrid.Type!.Value);
            Assert.NotNull(tableGrid.StyleTableProperties!.GetFirstChild<TableBorders>());

            AssertNoOpenXmlErrors(doc);
            AssertQuickPasses(doc);
        }
        finally
        {
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void Created_blank_layout_declares_compatibility_mode_15_so_Word_preserves_its_repeaters()
    {
        // A layout that declares no compatibility mode is mode 12 to Word - Word 2007, where
        // repeating-section content controls do not exist - so an interactive Word save converted every
        // repeater to a plain rich-text control and dropped its w15:dataBinding (GitHub issue #51). The
        // scaffolded part declares what every stock corpus layout declares.
        var outputPath = TempOutputPath();
        try
        {
            LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), outputPath);

            using var doc = WordprocessingDocument.Open(outputPath, false);
            var main = doc.MainDocumentPart!;
            var settingsPart = main.DocumentSettingsPart;
            Assert.NotNull(settingsPart);

            var compat = settingsPart!.Settings!.GetFirstChild<Compatibility>();
            Assert.NotNull(compat);
            var setting = Assert.Single(compat!.Elements<CompatibilitySetting>());
            Assert.Equal(CompatSettingNameValues.CompatibilityMode, setting.Name!.Value);
            Assert.Equal("http://schemas.microsoft.com/office/word", setting.Uri!.Value);
            Assert.Equal("15", setting.Val!.Value);

            // The value the rest of the tool reasons about, read back through the public helper.
            Assert.Equal(15, DocumentSettingsScaffold.ReadCompatibilityMode(main));

            AssertNoOpenXmlErrors(doc);
            AssertQuickPasses(doc);
        }
        finally
        {
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void Blank_build_with_a_repeater_is_free_of_the_compatibility_mode_warning_end_to_end()
    {
        // The end-to-end point of the scaffold: author the construct that the missing mode used to put at
        // risk, and the layout comes out clean - no compatibility-mode finding anywhere in its own
        // post-edit validation.
        var outputPath = TempOutputPath();
        try
        {
            LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), outputPath);

            using (var doc = WordprocessingDocument.Open(outputPath, true))
            {
                LayoutEditor.InsertRepeaterTable(
                    doc,
                    "/Header/Line",
                    ["ItemNo_Line", "Description_Line"],
                    new Location { Type = LocationKind.DocumentEnd },
                    new RepeaterTableOptions());
                doc.MainDocumentPart!.Document!.Save();
            }

            var result = LayoutValidator.Quick(outputPath);
            Assert.DoesNotContain(result.Findings, f => f.Check == "compatibility-mode");
            Assert.True(result.Passed);
        }
        finally
        {
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void Created_blank_layout_can_take_a_footer_insert_immediately()
    {
        // The end-to-end point of the scaffolding above, through the real edit path.
        var outputPath = TempOutputPath();
        try
        {
            LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), outputPath);

            EditResult result;
            using (var doc = WordprocessingDocument.Open(outputPath, true))
            {
                result = LayoutEditor.InsertField(
                    doc, "/Header/CustomerAddress1",
                    new Location { Type = LocationKind.DocumentEnd, Part = LayoutPart.Footer });
                doc.MainDocumentPart!.Document!.Save();
                foreach (var footer in doc.MainDocumentPart.FooterParts)
                {
                    footer.Footer?.Save();
                }
            }

            using var reopened = WordprocessingDocument.Open(outputPath, false);
            var inventory = LayoutReader.Read(reopened);
            Assert.Contains(inventory.Controls, c => c.SdtId == result.ControlId && c.Part == result.Part);
            Assert.Contains("footer", result.Part, StringComparison.OrdinalIgnoreCase);
            AssertNoOpenXmlErrors(reopened);
        }
        finally
        {
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void Create_headingText_overrides_the_default_heading_and_empty_means_none()
    {
        var outputPath = TempOutputPath();
        try
        {
            LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), outputPath, headingText: "Sales Order Confirmation");
            using (var doc = WordprocessingDocument.Open(outputPath, false))
            {
                var heading = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();
                Assert.Equal("Sales Order Confirmation", heading.InnerText);
                Assert.NotNull(heading.GetFirstChild<Run>()!.RunProperties!.Bold);
            }

            LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), outputPath, headingText: "");
            using (var doc = WordprocessingDocument.Open(outputPath, false))
            {
                var body = doc.MainDocumentPart!.Document!.Body!;
                Assert.Equal(string.Empty, body.Elements<Paragraph>().Single().InnerText);
                AssertNoOpenXmlErrors(doc);
            }
        }
        finally
        {
            File.Delete(outputPath);
        }
    }

    // ---- create from a corpus .docx layout as schemaSource ----

    [Fact]
    public void Create_from_a_corpus_docx_passes_validation_and_yields_an_equivalent_schema_with_one_BC_part()
    {
        var outputPath = TempOutputPath();
        try
        {
            var sourceTree = SchemaProvider.FromLayout(Corpus.Path(Corpus.SalesInvoice));

            var result = LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), outputPath);

            Assert.True(File.Exists(outputPath));
            Assert.Equal(Path.GetFullPath(outputPath), result.OutputPath);
            Assert.False(result.UsedTemplate);
            Assert.False(result.ReplacedExistingBcPart);
            Assert.False(string.IsNullOrWhiteSpace(result.StoreItemId));
            Assert.Equal(sourceTree.Report.ReportName, result.ReportName);
            Assert.Equal(sourceTree.Report.ReportId, result.ReportId);
            Assert.Equal(sourceTree.Report.Namespace, result.Namespace);

            // A freshly created (non-template) layout's own post-build quick validation is always clean, and
            // there is nothing to warn about (no template BC part was ever replaced).
            Assert.Equal("quick", result.QuickValidation.Level);
            Assert.True(result.QuickValidation.Passed);
            Assert.Equal(0, result.QuickValidation.ErrorCount);

            using var doc = WordprocessingDocument.Open(outputPath, false);
            AssertNoOpenXmlErrors(doc);
            AssertQuickPasses(doc);

            // Exactly one BC custom XML part.
            var main = doc.MainDocumentPart!;
            Assert.Single(FindBcCustomXmlParts(main));

            // Same report identity + an equivalent schema tree (same data-item/column counts) as the source;
            // a FRESH storeItemID, not the source's own.
            var createdTree = SchemaProvider.FromLayout(doc);
            Assert.Equal(sourceTree.Report.ReportName, createdTree.Report.ReportName);
            Assert.Equal(sourceTree.Report.ReportId, createdTree.Report.ReportId);
            Assert.Equal(sourceTree.Report.Namespace, createdTree.Report.Namespace);
            Assert.Equal(result.StoreItemId, createdTree.Report.StoreItemId);
            Assert.NotEqual(sourceTree.Report.StoreItemId, createdTree.Report.StoreItemId, StringComparer.OrdinalIgnoreCase);

            Assert.Equal(sourceTree.AllDataItems(includeSystem: true).Count(), createdTree.AllDataItems(includeSystem: true).Count());
            Assert.Equal(sourceTree.AllColumns(includeSystem: true).Count(), createdTree.AllColumns(includeSystem: true).Count());
        }
        finally
        {
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }

    [Fact]
    public void Create_from_a_corpus_docx_storeItemId_is_a_fresh_well_formed_guid_token()
    {
        var outputPath = TempOutputPath();
        try
        {
            var result = LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), outputPath);

            Assert.Matches(@"^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$", result.StoreItemId);
        }
        finally
        {
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }

    // ---- editability proof: the created layout can immediately take insert_field / insert_repeater_table ----

    [Fact]
    public void Created_layout_is_immediately_editable_a_freshly_inserted_field_reads_back_bound_to_the_new_part()
    {
        var outputPath = TempOutputPath();
        try
        {
            var result = LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), outputPath);

            EditResult fieldResult;
            using (var doc = WordprocessingDocument.Open(outputPath, true))
            {
                fieldResult = LayoutEditor.InsertField(
                    doc, "/Header/CustomerAddress1", new Location { Type = LocationKind.DocumentEnd });
                doc.MainDocumentPart!.Document!.Save();
            }

            Assert.Equal("insert_field", fieldResult.Operation);
            Assert.Equal("Field", fieldResult.Kind);

            using var reopened = WordprocessingDocument.Open(outputPath, false);
            AssertNoOpenXmlErrors(reopened);
            AssertQuickPasses(reopened);

            var inventory = LayoutReader.Read(reopened);
            var control = Assert.Single(inventory.Controls, c => c.SdtId == fieldResult.ControlId);
            Assert.Equal(ControlKind.Field, control.Kind);
            Assert.Equal(result.StoreItemId, control.StoreItemId, StringComparer.OrdinalIgnoreCase);
            Assert.Equal(fieldResult.XPath, control.XPath);
        }
        finally
        {
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }

    [Fact]
    public void Created_layout_is_immediately_editable_a_freshly_inserted_repeater_table_passes_quick_validation()
    {
        var outputPath = TempOutputPath();
        try
        {
            var result = LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), outputPath);

            EditResult tableResult;
            using (var doc = WordprocessingDocument.Open(outputPath, true))
            {
                tableResult = LayoutEditor.InsertRepeaterTable(
                    doc,
                    "/Header/Line",
                    new[] { "ItemNo_Line", "Description_Line", "TransHeaderAmount" },
                    new Location { Type = LocationKind.DocumentEnd },
                    new RepeaterTableOptions());
                doc.MainDocumentPart!.Document!.Save();
            }

            Assert.Equal("insert_repeater_table", tableResult.Operation);
            Assert.Equal("Repeater", tableResult.Kind);
            Assert.Equal(3, tableResult.ColumnCount);

            using var reopened = WordprocessingDocument.Open(outputPath, false);
            AssertNoOpenXmlErrors(reopened);
            AssertQuickPasses(reopened);

            var inventory = LayoutReader.Read(reopened);
            var control = Assert.Single(inventory.Controls, c => c.SdtId == tableResult.ControlId);
            Assert.Equal(ControlKind.Repeater, control.Kind);
            Assert.Equal(result.StoreItemId, control.StoreItemId, StringComparer.OrdinalIgnoreCase);
        }
        finally
        {
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }

    // ---- create from a standalone schema .xml ----

    [Fact]
    public void Create_from_a_standalone_schema_xml_passes_validation_and_yields_an_equivalent_schema()
    {
        var xmlPath = Path.Combine(Path.GetTempPath(), $"bcwl-schema-{Guid.NewGuid():N}.xml");
        var outputPath = TempOutputPath();
        try
        {
            // Extract the corpus layout's BC part bytes verbatim to a standalone .xml - exactly how a real
            // exported schema file would look, and the same convention LayoutBuilder itself expects here.
            using (var sourceDoc = WordprocessingDocument.Open(Corpus.Path(Corpus.SalesInvoice), false))
            {
                var bcPart = FindBcCustomXmlParts(sourceDoc.MainDocumentPart!).Single();
                using var partStream = bcPart.GetStream(FileMode.Open, FileAccess.Read);
                using var buffer = new MemoryStream();
                partStream.CopyTo(buffer);
                File.WriteAllBytes(xmlPath, buffer.ToArray());
            }

            var sourceTree = SchemaProvider.FromLayout(Corpus.Path(Corpus.SalesInvoice));

            var result = LayoutBuilder.Create(xmlPath, outputPath);

            Assert.False(result.UsedTemplate);
            Assert.False(result.ReplacedExistingBcPart);
            Assert.Equal(sourceTree.Report.ReportName, result.ReportName);
            Assert.Equal(sourceTree.Report.ReportId, result.ReportId);
            Assert.Equal(sourceTree.Report.Namespace, result.Namespace);
            Assert.False(string.IsNullOrWhiteSpace(result.StoreItemId));

            using var createdDoc = WordprocessingDocument.Open(outputPath, false);
            AssertNoOpenXmlErrors(createdDoc);
            AssertQuickPasses(createdDoc);
            Assert.Single(FindBcCustomXmlParts(createdDoc.MainDocumentPart!));

            var createdTree = SchemaProvider.FromLayout(createdDoc);
            Assert.Equal(sourceTree.Report.ReportName, createdTree.Report.ReportName);
            Assert.Equal(result.StoreItemId, createdTree.Report.StoreItemId);
            Assert.Equal(sourceTree.AllDataItems(includeSystem: true).Count(), createdTree.AllDataItems(includeSystem: true).Count());
            Assert.Equal(sourceTree.AllColumns(includeSystem: true).Count(), createdTree.AllColumns(includeSystem: true).Count());
        }
        finally
        {
            if (File.Exists(xmlPath))
            {
                File.Delete(xmlPath);
            }

            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }

    [Fact]
    public void Create_from_a_standalone_schema_xml_is_also_immediately_editable()
    {
        var xmlPath = Path.Combine(Path.GetTempPath(), $"bcwl-schema-{Guid.NewGuid():N}.xml");
        var outputPath = TempOutputPath();
        try
        {
            using (var sourceDoc = WordprocessingDocument.Open(Corpus.Path(Corpus.SalesInvoice), false))
            {
                var bcPart = FindBcCustomXmlParts(sourceDoc.MainDocumentPart!).Single();
                using var partStream = bcPart.GetStream(FileMode.Open, FileAccess.Read);
                using var buffer = new MemoryStream();
                partStream.CopyTo(buffer);
                File.WriteAllBytes(xmlPath, buffer.ToArray());
            }

            var result = LayoutBuilder.Create(xmlPath, outputPath);

            EditResult fieldResult;
            using (var doc = WordprocessingDocument.Open(outputPath, true))
            {
                fieldResult = LayoutEditor.InsertField(
                    doc, "/Header/CustomerAddress1", new Location { Type = LocationKind.DocumentEnd });
                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(outputPath, false);
            AssertNoOpenXmlErrors(reopened);
            AssertQuickPasses(reopened);

            var inventory = LayoutReader.Read(reopened);
            var control = Assert.Single(inventory.Controls, c => c.SdtId == fieldResult.ControlId);
            Assert.Equal(result.StoreItemId, control.StoreItemId, StringComparer.OrdinalIgnoreCase);
        }
        finally
        {
            if (File.Exists(xmlPath))
            {
                File.Delete(xmlPath);
            }

            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }

    // ---- templatePath: starting from a second corpus docx that already has its own BC part ----

    [Fact]
    public void Create_with_templatePath_that_is_a_full_BC_layout_throws_TemplateNotUnboundException_and_writes_nothing()
    {
        // A PRIVATE per-test output directory, so the no-leftovers scan below cannot race a concurrent
        // test class's own in-flight Create staging a .bcwl-build-* file in the SHARED temp dir (the
        // staged build file lives next to outputPath, so scanning outputPath's own private dir is exact).
        var outputDir = Path.Combine(Path.GetTempPath(), $"bcwl-b22-refusal-{Guid.NewGuid():N}");
        Directory.CreateDirectory(outputDir);
        var outputPath = Path.Combine(outputDir, "out.docx");
        try
        {
            // StandardStatement (the "template" here) is a fully-populated real layout with its own bound
            // content controls, built against ITS OWN schema/storeItemID - deliberately a mismatched pairing
            // (SalesInvoice's schema replaces it), so this test can assert the REFUSAL actually fires
            // rather than only exercising the clean-template path.
            var ex = Assert.Throws<TemplateNotUnboundException>(() => LayoutBuilder.Create(
                Corpus.Path(Corpus.SalesInvoice), outputPath, Corpus.Path(Corpus.StandardStatement)));

            Assert.True(ex.ErrorCount > 0);
            Assert.Contains(ex.ErrorCount.ToString(), ex.Message);
            Assert.Contains("full BC layout", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(Corpus.Path(Corpus.StandardStatement), ex.Message, StringComparison.Ordinal);

            // The whole point of an atomic, staged build: refusing must leave NOTHING behind - not a
            // half-written outputPath, and not a stray .bcwl-build-* temp file in its directory either.
            Assert.False(File.Exists(outputPath));
            Assert.Empty(Directory.GetFiles(outputDir));
        }
        finally
        {
            if (Directory.Exists(outputDir))
            {
                Directory.Delete(outputDir, recursive: true);
            }
        }
    }

    [Fact]
    public void Create_with_templatePath_that_has_an_existing_BC_part_but_zero_bound_controls_succeeds_and_stays_editable()
    {
        // The edge case the refusal must NOT trip on: a template whose BC part is present but has nothing
        // bound to it. Built synthetically ('s own suggestion) via LayoutBuilder.Create itself
        // with NO templatePath - a freshly created layout's body carries only a heading paragraph, so it is
        // guaranteed to have a real BC part and exactly zero content controls, without touching any corpus
        // file or exercising remove_control repeatedly.
        var templatePath = TempOutputPath();
        var outputPath = TempOutputPath();
        try
        {
            var templateResult = LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), templatePath);
            Assert.False(templateResult.UsedTemplate);

            using (var templateDoc = WordprocessingDocument.Open(templatePath, false))
            {
                Assert.Empty(LayoutReader.Read(templateDoc).Controls);
            }

            var sourceTree = SchemaProvider.FromLayout(Corpus.Path(Corpus.SalesInvoice));

            // Same schemaSource as the template happened to be built from - a fresh storeItemID is generated
            // every call regardless, so this still proves the part is genuinely REPLACED (not merely reused),
            // even though nothing was stale to refuse.
            var result = LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), outputPath, templatePath);

            Assert.True(result.UsedTemplate);
            Assert.True(result.ReplacedExistingBcPart);
            Assert.NotEqual(templateResult.StoreItemId, result.StoreItemId, StringComparer.OrdinalIgnoreCase);
            Assert.Equal(sourceTree.Report.ReportName, result.ReportName);

            Assert.Equal("quick", result.QuickValidation.Level);
            Assert.True(result.QuickValidation.Passed);
            Assert.Equal(0, result.QuickValidation.ErrorCount);

            using (var doc = WordprocessingDocument.Open(outputPath, false))
            {
                AssertNoOpenXmlErrors(doc);
                AssertQuickPasses(doc);
                Assert.Single(FindBcCustomXmlParts(doc.MainDocumentPart!));
                Assert.Empty(LayoutReader.Read(doc).Controls);
            }

            // Still editable: a fresh control binds to the new storeItemID and reads back correctly.
            EditResult fieldResult;
            using (var doc = WordprocessingDocument.Open(outputPath, true))
            {
                fieldResult = LayoutEditor.InsertField(
                    doc, "/Header/CustomerAddress1", new Location { Type = LocationKind.DocumentEnd });
                doc.MainDocumentPart!.Document!.Save();
            }

            using (var reopened = WordprocessingDocument.Open(outputPath, false))
            {
                var inventory = LayoutReader.Read(reopened);
                var control = Assert.Single(inventory.Controls, c => c.SdtId == fieldResult.ControlId);
                Assert.Equal(ControlKind.Field, control.Kind);
                Assert.Equal(result.StoreItemId, control.StoreItemId, StringComparer.OrdinalIgnoreCase);
            }
        }
        finally
        {
            if (File.Exists(templatePath))
            {
                File.Delete(templatePath);
            }

            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }

    [Fact]
    public void Create_with_templatePath_when_template_has_no_existing_BC_part_reports_not_replaced()
    {
        // A standalone schema .xml re-saved as a ".docx"-like template would be unusual; instead, build a
        // template from a corpus file but strip its BC part first so "no existing BC part" is exercised too.
        var templatePath = Path.Combine(Path.GetTempPath(), $"bcwl-template-noBC-{Guid.NewGuid():N}.docx");
        var outputPath = TempOutputPath();
        try
        {
            File.Copy(Corpus.Path(Corpus.InventoryOrderDetails), templatePath, overwrite: true);
            using (var doc = WordprocessingDocument.Open(templatePath, true))
            {
                var main = doc.MainDocumentPart!;
                foreach (var part in FindBcCustomXmlParts(main))
                {
                    if (part.CustomXmlPropertiesPart is { } props)
                    {
                        part.DeletePart(props);
                    }

                    main.DeletePart(part);
                }
            }

            var result = LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), outputPath, templatePath);

            Assert.True(result.UsedTemplate);
            // No BC part was replaced (it was already gone before Create ever ran) - the stale-controls
            // refusal is gated on ReplacedExistingBcPart specifically, so it does not fire
            // here regardless of whether the template's own now-orphaned body content happens to
            // quick-validate cleanly or not; this call must still succeed.
            Assert.False(result.ReplacedExistingBcPart);

            using var createdDoc = WordprocessingDocument.Open(outputPath, false);
            Assert.Single(FindBcCustomXmlParts(createdDoc.MainDocumentPart!));
        }
        finally
        {
            if (File.Exists(templatePath))
            {
                File.Delete(templatePath);
            }

            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }

    [Fact]
    public void Create_with_templatePath_never_injects_header_footer_parts_into_the_templates_own_shell()
    {
        // Header/footer scaffolding is a BLANK-build affordance only: a template brings its
        // own header/footer story - branded letterhead, a distinct first page, or deliberately neither -
        // and adding parts to it would be authoring, not scaffolding.
        var templatePath = SyntheticLayout.Create(SyntheticLayout.PlainParagraph("template body"));
        var outputPath = TempOutputPath();
        try
        {
            var result = LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), outputPath, templatePath);
            Assert.True(result.UsedTemplate);

            using var doc = WordprocessingDocument.Open(outputPath, false);
            Assert.Empty(doc.MainDocumentPart!.HeaderParts);
            Assert.Empty(doc.MainDocumentPart.FooterParts);
        }
        finally
        {
            if (File.Exists(templatePath))
            {
                File.Delete(templatePath);
            }

            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }

    [Fact]
    public void Create_with_templatePath_never_injects_a_settings_part_into_the_templates_own_shell()
    {
        // Same BLANK-build-only contract as the two scaffolds around this test, and for a sharper reason:
        // compatibility mode also selects Word's layout metrics, so retrofitting a mode onto a shell that
        // already renders somehow could move its pagination. The risk a low-mode template introduces is
        // REPORTED instead, by LayoutValidator's compatibility-mode check (and only once the layout actually
        // contains a repeater) - see DocumentSettingsScaffold's remarks.
        var templatePath = SyntheticLayout.Create(SyntheticLayout.PlainParagraph("template body"));
        var outputPath = TempOutputPath();
        try
        {
            var result = LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), outputPath, templatePath);
            Assert.True(result.UsedTemplate);

            using var doc = WordprocessingDocument.Open(outputPath, false);
            Assert.Null(doc.MainDocumentPart!.DocumentSettingsPart);

            // No repeater in the template shell, so the mode it lacks is not yet a risk and must not be
            // reported - the create-time envelope stays clean.
            Assert.DoesNotContain(
                result.QuickValidation.Findings, f => f.Check == "compatibility-mode");
        }
        finally
        {
            if (File.Exists(templatePath))
            {
                File.Delete(templatePath);
            }

            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }

    [Fact]
    public void Create_with_templatePath_never_injects_a_styles_part_into_the_templates_own_shell()
    {
        // Same contract as the header/footer scaffold above: the default styles part is a BLANK-build
        // affordance only. A template brings its own look - styles, theme, or deliberately neither -
        // and injecting docDefaults into it would restyle content the caller authored elsewhere.
        var templatePath = SyntheticLayout.Create(SyntheticLayout.PlainParagraph("template body"));
        var outputPath = TempOutputPath();
        try
        {
            var result = LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), outputPath, templatePath);
            Assert.True(result.UsedTemplate);

            using var doc = WordprocessingDocument.Open(outputPath, false);
            Assert.Null(doc.MainDocumentPart!.StyleDefinitionsPart);
        }
        finally
        {
            if (File.Exists(templatePath))
            {
                File.Delete(templatePath);
            }

            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }

    [Fact]
    public void Create_with_templatePath_preserves_the_templates_own_styles_part_rather_than_rescaffolding_it()
    {
        // A template WITH a styles part must keep it exactly as authored - the scaffold's own defaults must
        // never win over a template's deliberate typography. Built by creating a blank layout (which ships
        // the scaffolded styles part) and then re-pinning its docDefaults font to Arial, so the assertion
        // below can tell "preserved the template's styles" apart from "re-ran the scaffold".
        var templatePath = TempOutputPath();
        var outputPath = TempOutputPath();
        try
        {
            LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), templatePath);
            using (var templateDoc = WordprocessingDocument.Open(templatePath, true))
            {
                var runDefaults = templateDoc.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                    .DocDefaults!.RunPropertiesDefault!.RunPropertiesBaseStyle!;
                var fonts = runDefaults.GetFirstChild<RunFonts>()!;
                fonts.Ascii = "Arial";
                fonts.HighAnsi = "Arial";
                templateDoc.MainDocumentPart.StyleDefinitionsPart.Styles.Save();
            }

            var result = LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), outputPath, templatePath);
            Assert.True(result.UsedTemplate);

            using var doc = WordprocessingDocument.Open(outputPath, false);
            var createdDefaults = doc.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .DocDefaults!.RunPropertiesDefault!.RunPropertiesBaseStyle!;
            Assert.Equal("Arial", createdDefaults.GetFirstChild<RunFonts>()!.Ascii!.Value);
        }
        finally
        {
            if (File.Exists(templatePath))
            {
                File.Delete(templatePath);
            }

            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }

    // ---- glossary part ----

    [Fact]
    public void Created_layout_has_a_GlossaryDocumentPart_containing_the_DefaultPlaceholder_docPart()
    {
        var outputPath = TempOutputPath();
        try
        {
            LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), outputPath);

            using var doc = WordprocessingDocument.Open(outputPath, false);
            var glossaryPart = doc.MainDocumentPart!.GlossaryDocumentPart;
            Assert.NotNull(glossaryPart);

            var hasDefaultPlaceholder = glossaryPart!.GlossaryDocument?.Descendants<DocPartName>()
                .Any(n => n.Val?.Value == DefaultPlaceholderDocPart) ?? false;
            Assert.True(hasDefaultPlaceholder);
        }
        finally
        {
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }

    // ---- error paths ----

    [Fact]
    public void Create_missing_schemaSource_throws_FileNotFoundException()
    {
        var ex = Assert.Throws<FileNotFoundException>(() =>
            LayoutBuilder.Create("Z:\\does-not-exist.docx", TempOutputPath()));
        Assert.Equal("Z:\\does-not-exist.docx", ex.FileName);
    }

    [Fact]
    public void Create_missing_templatePath_throws_FileNotFoundException()
    {
        var ex = Assert.Throws<FileNotFoundException>(() =>
            LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), TempOutputPath(), "Z:\\does-not-exist.docx"));
        Assert.Equal("Z:\\does-not-exist.docx", ex.FileName);
    }

    [Fact]
    public void Create_from_a_non_BC_xml_throws_InvalidDataException()
    {
        var badXmlPath = Path.Combine(Path.GetTempPath(), $"bcwl-badschema-{Guid.NewGuid():N}.xml");
        try
        {
            File.WriteAllText(badXmlPath, "<SomeOtherRoot xmlns=\"urn:not-bc\"><Foo/></SomeOtherRoot>");

            Assert.Throws<InvalidDataException>(() => LayoutBuilder.Create(badXmlPath, TempOutputPath()));
        }
        finally
        {
            if (File.Exists(badXmlPath))
            {
                File.Delete(badXmlPath);
            }
        }
    }

    [Fact]
    public void Create_from_a_docx_with_no_BC_part_throws_InvalidDataException()
    {
        // Synthetic docx with a main document part but no custom XML parts at all.
        var path = SyntheticLayout.Create(SyntheticLayout.PlainParagraph("no BC part here"));
        try
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                // SyntheticLayout.Create always attaches a BC part; strip it to exercise "no BC part".
                var main = doc.MainDocumentPart!;
                foreach (var part in FindBcCustomXmlParts(main))
                {
                    if (part.CustomXmlPropertiesPart is { } props)
                    {
                        part.DeletePart(props);
                    }

                    main.DeletePart(part);
                }
            }

            Assert.Throws<InvalidDataException>(() => LayoutBuilder.Create(path, TempOutputPath()));
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    // ---- outputPath handling ----

    [Fact]
    public void Create_creates_the_parent_directory_of_outputPath_when_missing()
    {
        var root = Path.Combine(Path.GetTempPath(), $"bcwl-createdir-{Guid.NewGuid():N}");
        var outputPath = Path.Combine(root, "nested", "layout.docx");
        try
        {
            Assert.False(Directory.Exists(root));

            var result = LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), outputPath);

            Assert.True(File.Exists(outputPath));
            Assert.Equal(Path.GetFullPath(outputPath), result.OutputPath);
        }
        finally
        {
            if (Directory.Exists(root))
            {
                Directory.Delete(root, recursive: true);
            }
        }
    }

    [Fact]
    public void Create_overwrites_an_existing_file_at_outputPath()
    {
        var outputPath = TempOutputPath();
        try
        {
            File.WriteAllText(outputPath, "not a real docx");

            var result = LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), outputPath);

            using var doc = WordprocessingDocument.Open(outputPath, false);
            AssertNoOpenXmlErrors(doc);
            Assert.Equal(Path.GetFullPath(outputPath), result.OutputPath);
        }
        finally
        {
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }

    // ---- atomic write: build-in-temp-then-move, never touch outputPath on failure ----

    private static bool HasStrayBuildTempFiles(string directory) =>
        Directory.Exists(directory) && Directory.EnumerateFiles(directory, ".bcwl-build-*").Any();

    [Fact]
    public void Create_leaves_no_stray_build_temp_file_behind_after_a_successful_create()
    {
        var outputPath = TempOutputPath();
        var dir = Path.GetDirectoryName(Path.GetFullPath(outputPath))!;
        try
        {
            LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), outputPath);

            Assert.True(File.Exists(outputPath));
            Assert.False(HasStrayBuildTempFiles(dir), "expected no leftover .bcwl-build-* temp file");
        }
        finally
        {
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }

    [Fact]
    public void Create_with_a_template_lacking_a_main_document_part_throws_and_leaves_a_preexisting_outputPath_byte_identical()
    {
        // A .docx-shaped OPC package with no main document part at all (never AddMainDocumentPart'd) - the
        // "template lacking a MainDocumentPart" nit the atomic-write rework specifically fixes: previously
        // the template would already have been copied directly onto outputPath before this was discovered,
        // corrupting/replacing whatever was there.
        var brokenTemplatePath = Path.Combine(Path.GetTempPath(), $"bcwl-template-nomainpart-{Guid.NewGuid():N}.docx");
        var outputPath = TempOutputPath();
        try
        {
            using (WordprocessingDocument.Create(brokenTemplatePath, WordprocessingDocumentType.Document))
            {
                // Deliberately never call AddMainDocumentPart().
            }

            var before = "pre-existing content that must survive"u8.ToArray();
            File.WriteAllBytes(outputPath, before);

            Assert.Throws<InvalidDataException>(
                () => LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), outputPath, brokenTemplatePath));

            Assert.Equal(before, File.ReadAllBytes(outputPath));

            var dir = Path.GetDirectoryName(Path.GetFullPath(outputPath))!;
            Assert.False(HasStrayBuildTempFiles(dir), "expected no leftover .bcwl-build-* temp file after a failure");
        }
        finally
        {
            if (File.Exists(brokenTemplatePath))
            {
                File.Delete(brokenTemplatePath);
            }

            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }

    [Fact]
    public void Create_when_outputPath_equals_templatePath_refreshes_the_file_without_corruption()
    {
        // The other nit the atomic-write rework fixes: outputPath and templatePath aliasing the same file
        // used to mean opening that file editable while ALSO being the thing File.Copy had just overwritten
        // it with a moment earlier - now the template is only ever read (via a copy into a private temp file)
        // before outputPath is touched at all. Uses a template with a BC part but ZERO bound controls (built
        // synthetically, same fixture shape as the "BC part but no bound controls" test above) so this
        // exercises ONLY the aliasing safety property - a template with real pre-existing bound controls
        // would now trip the stale-controls refusal, covered together with aliasing by
        // Create_when_outputPath_equals_templatePath_and_the_aliased_file_is_a_full_BC_layout_throws_and_leaves_it_untouched below.
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-template-self-{Guid.NewGuid():N}.docx");
        try
        {
            LayoutBuilder.Create(Corpus.Path(Corpus.StandardStatement), path);

            var result = LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), path, path);

            Assert.True(result.UsedTemplate);
            Assert.True(result.ReplacedExistingBcPart);

            using var doc = WordprocessingDocument.Open(path, false);
            AssertNoOpenXmlErrors(doc);
            Assert.Single(FindBcCustomXmlParts(doc.MainDocumentPart!));

            var createdTree = SchemaProvider.FromLayout(doc);
            Assert.Equal("Standard_Sales_Invoice", createdTree.Report.ReportName);
            Assert.Equal(result.StoreItemId, createdTree.Report.StoreItemId);
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void Create_when_outputPath_equals_templatePath_and_the_aliased_file_is_a_full_BC_layout_throws_and_leaves_it_untouched()
    {
        // Combines the aliasing edge case above with the stale-controls refusal: outputPath
        // and templatePath name the SAME real, fully-populated layout. LayoutBuilder.Create reads
        // schemaSource/templatePath and builds entirely into a private temp file BEFORE ever touching
        // outputPath - so even though outputPath IS the template here, refusing must leave it byte-for-byte
        // as it was, not merely "not deleted".
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-template-self-full-{Guid.NewGuid():N}.docx");
        try
        {
            File.Copy(Corpus.Path(Corpus.StandardStatement), path, overwrite: true);
            var before = File.ReadAllBytes(path);

            Assert.Throws<TemplateNotUnboundException>(
                () => LayoutBuilder.Create(Corpus.Path(Corpus.SalesInvoice), path, path));

            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }
}
