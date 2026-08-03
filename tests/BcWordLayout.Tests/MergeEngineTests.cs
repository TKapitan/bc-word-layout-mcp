using System.Xml.Linq;
using BcWordLayout.Domain;
using BcWordLayout.Merge;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Blip = DocumentFormat.OpenXml.Drawing.Blip;

namespace BcWordLayout.Tests;

public class MergeEngineTests
{
    private static string TempDocxPath() =>
        Path.Combine(Path.GetTempPath(), $"bcwl-merge-out-{Guid.NewGuid():N}.docx");

    private static bool IsRepeaterItemSdt(SdtElement sdt) =>
        sdt.GetFirstChild<SdtProperties>()?.Elements()
            .Any(e => e.LocalName == "repeatingSectionItem" && e.NamespaceUri == OoxmlNames.W15) == true;

    /// <summary>Yields every (hosting part, root element) pair the merge engine itself walks: the main
    /// document plus every header and footer — lets picture-fill tests search all of them uniformly.</summary>
    private static IEnumerable<(OpenXmlPart Part, OpenXmlElement? Root)> DocumentParts(WordprocessingDocument doc)
    {
        var main = doc.MainDocumentPart!;
        yield return (main, main.Document);

        foreach (var header in main.HeaderParts)
        {
            yield return (header, header.Header);
        }

        foreach (var footer in main.FooterParts)
        {
            yield return (footer, footer.Footer);
        }
    }

    private static readonly byte[] PngSignature = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };

    private static bool ContainsAsciiToken(byte[] data, string token)
    {
        var pattern = System.Text.Encoding.ASCII.GetBytes(token);
        for (var i = 0; i <= data.Length - pattern.Length; i++)
        {
            var match = true;
            for (var j = 0; j < pattern.Length; j++)
            {
                if (data[i + j] != pattern[j])
                {
                    match = false;
                    break;
                }
            }

            if (match)
            {
                return true;
            }
        }

        return false;
    }

    [Fact]
    public void Field_fill_matches_the_generators_exact_value()
    {
        const string FieldXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:CompanyName[1]";

        var body = SyntheticLayout.BoundField(FieldXPath, SyntheticLayout.GoodItemId);
        var layoutPath = SyntheticLayout.Create(body);
        var outputPath = TempDocxPath();

        try
        {
            var options = new MergeOptions { Seed = 12345, Rows = 1 };
            var result = MergeEngine.Merge(layoutPath, outputPath, options);

            Assert.Equal(1, result.Stats.FieldsFilled);
            Assert.Equal(0, result.Stats.Unresolved);
            Assert.Empty(result.Warnings);

            // Compute the exact expected value via the same generator pipeline, independently.
            var schema = SchemaProvider.FromLayout(layoutPath);
            var expected = SampleDataGenerator.Generate(
                schema, new SampleDataOptions { Seed = options.Seed, Rows = options.Rows });
            XNamespace ns = expected.Namespace;
            var expectedValue = expected.Xml.Root!.Element(ns + "Header")!.Element(ns + "CompanyName")!.Value;

            using var doc = WordprocessingDocument.Open(outputPath, false);
            var sdt = doc.MainDocumentPart!.Document!.Descendants<SdtElement>().Single();
            var actualValue = sdt.Descendants<Text>().First().Text;

            Assert.False(string.IsNullOrEmpty(actualValue));
            Assert.Equal(expectedValue, actualValue);
        }
        finally
        {
            File.Delete(layoutPath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void Repeater_expands_configured_row_count_with_reanchored_per_row_values()
    {
        const string RepeaterXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line";
        const string FieldXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:ItemNo_Line[1]";
        var schemaXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<NavWordReportXmlPart xmlns=\"urn:microsoft-dynamics-nav/reports/TestReport/50000/\">"
            + "<BCReportInformation><CreationDateTime>2026-01-01</CreationDateTime></BCReportInformation>"
            + "<Header><CompanyName>Contoso</CompanyName><Line><ItemNo_Line>X</ItemNo_Line></Line></Header>"
            + "</NavWordReportXmlPart>";

        var body = SyntheticLayout.RepeaterWithField(RepeaterXPath, FieldXPath, SyntheticLayout.GoodItemId);
        var layoutPath = SyntheticLayout.Create(body, datasetXml: schemaXml);
        var outputPath = TempDocxPath();

        try
        {
            var options = new MergeOptions { Seed = 555, Rows = 3 };
            var result = MergeEngine.Merge(layoutPath, outputPath, options);

            Assert.Equal(1, result.Stats.RepeatersExpanded);
            Assert.Equal(3, result.Stats.RowsGenerated);
            Assert.Equal(0, result.Stats.Unresolved);

            // Control for the re-anchoring guard (see the "divergent binding" tests below): a normal
            // in-item binding — the field's XPath is a genuine descendant of the repeater's own row path —
            // must keep re-anchoring row-relatively with ZERO warnings. The guard only ever suppresses
            // re-anchoring on a STRUCTURAL mismatch; a matching prefix must never trip it.
            Assert.Empty(result.Warnings);

            var schema = SchemaProvider.FromLayout(layoutPath);
            var expected = SampleDataGenerator.Generate(
                schema, new SampleDataOptions { Seed = options.Seed, Rows = options.Rows });
            XNamespace ns = expected.Namespace;
            var expectedValues = expected.Xml.Root!.Element(ns + "Header")!.Elements(ns + "Line")
                .Select(l => l.Element(ns + "ItemNo_Line")!.Value)
                .ToList();

            using var doc = WordprocessingDocument.Open(outputPath, false);
            var rows = doc.MainDocumentPart!.Document!.Descendants<SdtElement>()
                .Where(IsRepeaterItemSdt)
                .ToList();

            Assert.Equal(3, rows.Count);
            var actualValues = rows.Select(r => r.Descendants<Text>().First().Text).ToList();

            // Exact, order-sensitive match proves re-anchoring selected a DIFFERENT Line node per clone
            // (a broken re-anchor would repeat the template's own Line[1] value in every row instead).
            Assert.Equal(expectedValues, actualValues);
            Assert.Equal(3, actualValues.Distinct().Count());
        }
        finally
        {
            File.Delete(layoutPath);
            File.Delete(outputPath);
        }
    }

    // ---- Pre-scan repeater bindings so generation only multiplies rows the document actually
    // repeats, instead of every repeating item in the schema regardless of whether anything reads it. ----

    [Fact]
    public void ScanRepeaterConsumedPaths_finds_only_the_data_item_a_real_repeater_targets()
    {
        const string RepeaterXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:LineA";
        const string FieldXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:LineA[1]/ns0:Value_LineA[1]";
        var schemaXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<NavWordReportXmlPart xmlns=\"urn:microsoft-dynamics-nav/reports/TestReport/50000/\">"
            + "<BCReportInformation><CreationDateTime>2026-01-01</CreationDateTime></BCReportInformation>"
            + "<Header><CompanyName>Contoso</CompanyName>"
            + "<LineA><Value_LineA>A</Value_LineA></LineA>"
            + "<LineB><Value_LineB>B</Value_LineB></LineB>"
            + "</Header></NavWordReportXmlPart>";

        var body = SyntheticLayout.RepeaterWithField(RepeaterXPath, FieldXPath, SyntheticLayout.GoodItemId);
        var layoutPath = SyntheticLayout.Create(body, datasetXml: schemaXml);

        try
        {
            var schema = SchemaProvider.FromLayout(layoutPath);
            using var doc = WordprocessingDocument.Open(layoutPath, false);

            var consumed = MergeEngine.ScanRepeaterConsumedPaths(doc.MainDocumentPart!, schema);

            Assert.Contains("/Header/LineA", consumed);
            Assert.DoesNotContain("/Header/LineB", consumed);
            Assert.DoesNotContain("/Header", consumed);
        }
        finally
        {
            File.Delete(layoutPath);
        }
    }

    [Fact]
    public void Merge_multiplies_rows_only_for_the_sibling_item_a_real_repeater_targets_rest_get_one_instance()
    {
        // Two sibling repeating items (LineA/LineB) under Header; only LineA has a real repeater in the
        // document. At Rows=10, LineA (repeater-consumed) must get all 10 rows while LineB (nothing reads
        // it) gets exactly one generated instance.
        const string RepeaterXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:LineA";
        const string FieldXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:LineA[1]/ns0:Value_LineA[1]";
        var schemaXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<NavWordReportXmlPart xmlns=\"urn:microsoft-dynamics-nav/reports/TestReport/50000/\">"
            + "<BCReportInformation><CreationDateTime>2026-01-01</CreationDateTime></BCReportInformation>"
            + "<Header><CompanyName>Contoso</CompanyName>"
            + "<LineA><Value_LineA>A</Value_LineA></LineA>"
            + "<LineB><Value_LineB>B</Value_LineB></LineB>"
            + "</Header></NavWordReportXmlPart>";

        var body = SyntheticLayout.RepeaterWithField(RepeaterXPath, FieldXPath, SyntheticLayout.GoodItemId);
        var layoutPath = SyntheticLayout.Create(body, datasetXml: schemaXml);
        var outputPath = TempDocxPath();

        try
        {
            var options = new MergeOptions { Seed = 1, Rows = 10 };
            var result = MergeEngine.Merge(layoutPath, outputPath, options);

            Assert.Equal(1, result.Stats.RepeatersExpanded);
            Assert.Equal(10, result.Stats.RowsGenerated); // LineA (consumed) got the full 10 rows.
            Assert.Equal(0, result.Stats.Unresolved);

            // Directly confirm LineB (the unconsumed sibling) got exactly ONE generated instance, by
            // reproducing the same scan-then-generate steps MergeEngine.Merge runs internally.
            var schema = SchemaProvider.FromLayout(layoutPath);
            using var doc = WordprocessingDocument.Open(layoutPath, false);
            var consumedPaths = MergeEngine.ScanRepeaterConsumedPaths(doc.MainDocumentPart!, schema);
            var dataset = SampleDataGenerator.Generate(
                schema,
                new SampleDataOptions { Seed = options.Seed, Rows = options.Rows, RepeaterConsumedPaths = consumedPaths });

            XNamespace ns = schema.Report.Namespace;
            var header = dataset.Xml.Root!.Element(ns + "Header")!;
            Assert.Equal(10, header.Elements(ns + "LineA").Count());
            Assert.Single(header.Elements(ns + "LineB"));
        }
        finally
        {
            File.Delete(layoutPath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void Merge_does_not_raise_sample_data_capped_when_only_an_unconsumed_deep_sibling_would_have_exhausted_the_budget()
    {
        // Header has a deep chain (J1 -> J2 -> J3) nothing in the document ever reads, alongside a real
        // repeater bound to Line. Earlier, generation would multiply the J-chain too (rows^depth),
        // exhausting MaxTotalInstances before Line - the item the document actually repeats - got its rows;
        // this test proves the fixed merge no longer caps here at all.
        const string RepeaterXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line";
        const string FieldXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:ItemNo_Line[1]";
        var schemaXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<NavWordReportXmlPart xmlns=\"urn:microsoft-dynamics-nav/reports/TestReport/50000/\">"
            + "<BCReportInformation><CreationDateTime>2026-01-01</CreationDateTime></BCReportInformation>"
            + "<Header><CompanyName>Contoso</CompanyName>"
            + "<J1><J2><J3><Leaf>x</Leaf></J3></J2></J1>"
            + "<Line><ItemNo_Line>X</ItemNo_Line></Line>"
            + "</Header></NavWordReportXmlPart>";

        var body = SyntheticLayout.RepeaterWithField(RepeaterXPath, FieldXPath, SyntheticLayout.GoodItemId);
        var layoutPath = SyntheticLayout.Create(body, datasetXml: schemaXml);
        var outputPath = TempDocxPath();

        try
        {
            var options = new MergeOptions { Seed = 1, Rows = 50, MaxTotalInstances = 100 };
            var result = MergeEngine.Merge(layoutPath, outputPath, options);

            Assert.Equal(1, result.Stats.RepeatersExpanded);
            Assert.Equal(50, result.Stats.RowsGenerated); // Line got its full Rows, not starved by the J-chain.
            Assert.Equal(0, result.Stats.Unresolved);
            Assert.DoesNotContain(result.Warnings, w => w.Kind == "sample-data-capped");
        }
        finally
        {
            File.Delete(layoutPath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void Field_bound_under_a_non_repeated_item_resolves_with_exactly_one_generated_instance()
    {
        // Header/Totals/Amount is bound by a plain FIELD - no repeater anywhere in the document - so Totals
        // is "case (c)": a data item with a bound field under it but no repeater. It must get exactly ONE
        // generated instance regardless of a large Rows value, and the field must still resolve cleanly.
        const string FieldXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Totals[1]/ns0:Amount[1]";
        var schemaXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<NavWordReportXmlPart xmlns=\"urn:microsoft-dynamics-nav/reports/TestReport/50000/\">"
            + "<BCReportInformation><CreationDateTime>2026-01-01</CreationDateTime></BCReportInformation>"
            + "<Header><CompanyName>Contoso</CompanyName><Totals><Amount>0</Amount></Totals></Header>"
            + "</NavWordReportXmlPart>";

        var body = SyntheticLayout.BoundField(FieldXPath, SyntheticLayout.GoodItemId);
        var layoutPath = SyntheticLayout.Create(body, datasetXml: schemaXml);
        var outputPath = TempDocxPath();

        try
        {
            var options = new MergeOptions { Seed = 1, Rows = 25 };
            var result = MergeEngine.Merge(layoutPath, outputPath, options);

            Assert.Equal(1, result.Stats.FieldsFilled);
            Assert.Equal(0, result.Stats.Unresolved);
            Assert.Empty(result.Warnings);

            var schema = SchemaProvider.FromLayout(layoutPath);
            using var doc = WordprocessingDocument.Open(layoutPath, false);
            var consumedPaths = MergeEngine.ScanRepeaterConsumedPaths(doc.MainDocumentPart!, schema);
            Assert.Empty(consumedPaths); // no repeater anywhere in this document.

            var dataset = SampleDataGenerator.Generate(
                schema,
                new SampleDataOptions { Seed = options.Seed, Rows = options.Rows, RepeaterConsumedPaths = consumedPaths });
            XNamespace ns = schema.Report.Namespace;
            Assert.Single(dataset.Xml.Root!.Element(ns + "Header")!.Elements(ns + "Totals"));
        }
        finally
        {
            File.Delete(layoutPath);
            File.Delete(outputPath);
        }
    }

    // ---- Re-anchoring must not drop a prefix that merely has the right STEP COUNT ----
    //
    // When a repeater row is cloned, an inner binding is re-anchored by dropping the repeater's own step
    // count from the front of its XPath. Dropping by COUNT alone, without checking that the dropped prefix
    // actually matches the repeater's path, silently mis-resolves three real shapes:
    //   (a) an equal-depth SIBLING - e.g. /Header/DocumentNo inside a /Header/Line repeater, a legitimate
    //       BC shape (a document no shown on every line). Three steps are dropped as if they were
    //       /Header/Line, the remainder is empty, and the control resolves to the ROW NODE - rendering
    //       that row's whole concatenated text instead of the document number, counted as filled, with
    //       zero warnings. The saved OOXML is fine and BC renders it correctly; it is the PREVIEW that
    //       lies, which is the worse failure for a tool whose promise is that a clean preview is trustworthy.
    //   (b) a DEEPER divergent path - evaluates its tail against the wrong node and reports a FALSE
    //       unresolved-binding error at validate_layout level=full for a layout BC renders fine.
    //   (c) a divergent NESTED repeater - gets expression "." and clones one bogus row.
    // Reachable through the tool surface, not just hand-authored layouts: insert_column mode=field
    // explicitly allows binding a parent/Header field inside a repeater row.
    //
    // All three tests below use the internal
    // MergeEngine.Merge(doc, dataset) overload with a hand-built SampleDataset, so the expected string
    // values are asserted directly rather than re-derived through SampleDataGenerator.

    /// <summary>Finds every sdt in <paramref name="root"/> whose own w:dataBinding carries the literal
    /// <paramref name="xpath"/> attribute value — MergeEngine never rewrites that attribute, so this finds
    /// a field control regardless of whether its value was resolved via row-relative re-anchoring or a
    /// document-root fallback.</summary>
    private static List<SdtElement> FindFieldSdts(OpenXmlElement root, string xpath) =>
        root.Descendants<SdtElement>()
            .Where(s => s.GetFirstChild<SdtProperties>()?.Elements()
                .Any(e => (e.LocalName == "dataBinding")
                    && e.GetAttributes().Any(a => a.LocalName == "xpath" && a.Value == xpath)) == true)
            .ToList();

    [Fact]
    public void Field_bound_to_an_equal_depth_header_sibling_inside_a_repeater_row_resolves_the_sibling_not_the_row_text()
    {
        // Shape (a): a field sits INSIDE a /Header/Line repeater row but is bound to an EQUAL-DEPTH
        // sibling under Header — 3 raw steps, exactly the repeater's own 3 consumed steps. Blindly dropping
        // the first 3 steps because the COUNT matches would resolve the control to the Line ROW NODE
        // itself, rendering its concatenated descendant text ("ITEM-4670Some description") instead of the
        // real Header-level value ("INV-001") — reproduced verbatim in the code review. The fix must
        // recognize the dropped prefix names a DIFFERENT element (DocumentNo, not Line) and fall back to
        // document-root evaluation instead of guessing.
        const string Ns = "urn:microsoft-dynamics-nav/reports/TestReport/50000/";
        const string RepeaterXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line";
        const string FieldXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:DocumentNo[1]";

        var body = SyntheticLayout.RepeaterWithField(RepeaterXPath, FieldXPath, SyntheticLayout.GoodItemId);
        var layoutPath = SyntheticLayout.Create(body);

        XNamespace ns = Ns;
        var xml = new XDocument(
            new XElement(
                ns + "NavWordReportXmlPart",
                new XElement(
                    ns + "Header",
                    new XElement(ns + "DocumentNo", "INV-001"),
                    new XElement(
                        ns + "Line",
                        new XElement(ns + "ItemNo_Line", "ITEM-4670"),
                        new XElement(ns + "Description_Line", "Some description")))));
        var dataset = new SampleDataset { Xml = xml, Namespace = Ns };

        try
        {
            using (var doc = WordprocessingDocument.Open(layoutPath, true))
            {
                var result = MergeEngine.Merge(doc, dataset);

                // Counted as filled (a real value WAS resolved and written) with zero unresolved bindings —
                // the defect's whole danger is that it looks clean by these counters alone.
                Assert.Equal(1, result.Stats.FieldsFilled);
                Assert.Equal(0, result.Stats.Unresolved);

                var warning = Assert.Single(result.Warnings, w => w.Kind == "xpath-fallback");
                Assert.Contains(FieldXPath, warning.Location);

                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(layoutPath, false);
            var fieldSdt = Assert.Single(FindFieldSdts(reopened.MainDocumentPart!.Document!, FieldXPath));
            var actualValue = fieldSdt.Descendants<Text>().First().Text;

            Assert.Equal("INV-001", actualValue);
            Assert.NotEqual("ITEM-4670Some description", actualValue);
        }
        finally
        {
            File.Delete(layoutPath);
        }
    }

    [Fact]
    public void Field_bound_deeper_than_the_repeater_but_structurally_divergent_resolves_from_the_document_root()
    {
        // Shape (b): a binding DEEPER than the repeater (4 raw steps vs. the repeater's 3) but whose
        // extra depth branches off a DIFFERENT element (Header/Totals/Amount, not Header/Line/...). The old
        // code would still drop 3 steps (count-only check: 4 >= 3) and evaluate the 1-step remainder
        // ("Amount") against the wrong context node (the Line row), typically finding nothing and raising a
        // FALSE unresolved-binding error for a layout BC renders fine. The fix must recognize the divergence
        // at step 3 (Totals != Line) and fall back to evaluating the full path from the document root.
        const string RepeaterXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line";
        const string FieldXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Totals[1]/ns0:Amount[1]";
        const string Ns = "urn:microsoft-dynamics-nav/reports/TestReport/50000/";

        var body = SyntheticLayout.RepeaterWithField(RepeaterXPath, FieldXPath, SyntheticLayout.GoodItemId);
        var layoutPath = SyntheticLayout.Create(body);

        XNamespace ns = Ns;
        var xml = new XDocument(
            new XElement(
                ns + "NavWordReportXmlPart",
                new XElement(
                    ns + "Header",
                    new XElement(ns + "Totals", new XElement(ns + "Amount", "999.99")),
                    new XElement(ns + "Line", new XElement(ns + "ItemNo_Line", "ITEM-4670")))));
        var dataset = new SampleDataset { Xml = xml, Namespace = Ns };

        try
        {
            using (var doc = WordprocessingDocument.Open(layoutPath, true))
            {
                var result = MergeEngine.Merge(doc, dataset);

                Assert.Equal(1, result.Stats.FieldsFilled);
                Assert.Equal(0, result.Stats.Unresolved);

                var warning = Assert.Single(result.Warnings, w => w.Kind == "xpath-fallback");
                Assert.Contains(FieldXPath, warning.Location);

                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(layoutPath, false);
            var fieldSdt = Assert.Single(FindFieldSdts(reopened.MainDocumentPart!.Document!, FieldXPath));
            Assert.Equal("999.99", fieldSdt.Descendants<Text>().First().Text);
        }
        finally
        {
            File.Delete(layoutPath);
        }
    }

    [Fact]
    public void Structurally_divergent_nested_repeater_resolves_the_real_matched_rows_instead_of_a_bogus_dot_clone()
    {
        // Shape (c): the INNER repeater itself (not just a field) is nested inside the outer repeater's
        // row, but its own binding is an equal-depth sibling of the outer row's path (Header/Extra, not
        // Header/Line) rather than a genuine descendant. The old code would drop all 3 steps (count-only
        // match), leave an EMPTY remainder, and evaluate expression "." against the Line row context —
        // matching that single row node and cloning exactly ONE bogus row regardless of how many real Extra
        // elements exist. The fix must recognize the divergence and fall back to evaluating the inner
        // repeater's absolute XPath from the document root, producing clones for the ACTUAL matched nodes.
        const string OuterXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line";
        const string InnerXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Extra";
        const string FieldXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Extra[1]/ns0:Value[1]";
        const string Ns = "urn:microsoft-dynamics-nav/reports/TestReport/50000/";

        var body = SyntheticLayout.NestedRepeater(OuterXPath, InnerXPath, FieldXPath, SyntheticLayout.GoodItemId);
        var layoutPath = SyntheticLayout.Create(body);

        XNamespace ns = Ns;
        var xml = new XDocument(
            new XElement(
                ns + "NavWordReportXmlPart",
                new XElement(
                    ns + "Header",
                    new XElement(ns + "Line", new XElement(ns + "ItemNo_Line", "ITEM-1")),
                    new XElement(ns + "Extra", new XElement(ns + "Value", "V1")),
                    new XElement(ns + "Extra", new XElement(ns + "Value", "V2")))));
        var dataset = new SampleDataset { Xml = xml, Namespace = Ns };

        try
        {
            using (var doc = WordprocessingDocument.Open(layoutPath, true))
            {
                var result = MergeEngine.Merge(doc, dataset);

                // 1 outer expansion (1 Line row) + 1 inner expansion (that row's own Extra clone attempt).
                Assert.Equal(2, result.Stats.RepeatersExpanded);

                // The bug would produce exactly 1 bogus "." clone here; the fix must produce the 2 REAL
                // matched Extra nodes instead — 1 outer row + 2 inner rows = 3.
                Assert.Equal(3, result.Stats.RowsGenerated);
                Assert.Equal(0, result.Stats.Unresolved);

                var warning = Assert.Single(result.Warnings, w => w.Kind == "xpath-fallback");
                Assert.Contains(InnerXPath, warning.Location);

                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(layoutPath, false);
            var fieldSdts = FindFieldSdts(reopened.MainDocumentPart!.Document!, FieldXPath);

            // Exactly 2 clones (the real Extra nodes), never 1 (the bogus "." clone) or 0.
            Assert.Equal(2, fieldSdts.Count);
            var actualValues = fieldSdts.Select(s => s.Descendants<Text>().First().Text).ToList();
            Assert.Equal(new[] { "V1", "V2" }, actualValues);
        }
        finally
        {
            File.Delete(layoutPath);
        }
    }

    [Fact]
    public void FlattenBindingsForRender_severs_live_bindings_but_keeps_merged_data_and_rows()
    {
        // Regression: BC layout controls are live data-bound, so a renderer (Word) re-syncs each control from
        // the un-populated custom XML part and re-evaluates repeating sections against it — showing field-name
        // placeholders and one template row instead of the merged data/rows. FlattenBindingsForRender severs
        // those links so the merged sample data and every cloned row actually render.
        const string RepeaterXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line";
        const string FieldXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:ItemNo_Line[1]";
        var schemaXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<NavWordReportXmlPart xmlns=\"urn:microsoft-dynamics-nav/reports/TestReport/50000/\">"
            + "<BCReportInformation><CreationDateTime>2026-01-01</CreationDateTime></BCReportInformation>"
            + "<Header><CompanyName>Contoso</CompanyName><Line><ItemNo_Line>X</ItemNo_Line></Line></Header>"
            + "</NavWordReportXmlPart>";

        var body = SyntheticLayout.RepeaterWithField(RepeaterXPath, FieldXPath, SyntheticLayout.GoodItemId);
        var layoutPath = SyntheticLayout.Create(body, datasetXml: schemaXml);
        var boundPath = TempDocxPath();
        var flattenedPath = TempDocxPath();

        try
        {
            var schema = SchemaProvider.FromLayout(layoutPath);
            XNamespace ns = schema.Report.Namespace;
            var expectedRowValues = SampleDataGenerator
                .Generate(schema, new SampleDataOptions { Seed = 7, Rows = 3 })
                .Xml.Root!.Element(ns + "Header")!.Elements(ns + "Line")
                .Select(l => l.Element(ns + "ItemNo_Line")!.Value)
                .ToList();

            // Default merge (no flatten) keeps its live bindings — the logical output validation relies on.
            MergeEngine.Merge(layoutPath, boundPath, new MergeOptions { Seed = 7, Rows = 3 });
            using (var boundDoc = WordprocessingDocument.Open(boundPath, false))
            {
                Assert.Contains(boundDoc.MainDocumentPart!.Document!.Descendants(),
                    e => e.LocalName == "dataBinding");
            }

            // Flatten severs every live binding + repeating-section marker.
            var result = MergeEngine.Merge(layoutPath, flattenedPath,
                new MergeOptions { Seed = 7, Rows = 3, FlattenBindingsForRender = true });
            Assert.Equal(3, result.Stats.RowsGenerated);

            using var doc = WordprocessingDocument.Open(flattenedPath, false);
            var descendants = doc.MainDocumentPart!.Document!.Descendants().ToList();
            Assert.DoesNotContain(descendants, e => e.LocalName == "dataBinding");
            Assert.DoesNotContain(descendants, e => e.LocalName == "repeatingSection");
            Assert.DoesNotContain(descendants, e => e.LocalName == "repeatingSectionItem");

            // The merged data and all three cloned rows survive as static content.
            var allText = string.Concat(descendants.OfType<Text>().Select(t => t.Text));
            Assert.All(expectedRowValues, v => Assert.Contains(v, allText));
            Assert.Equal(3, expectedRowValues.Distinct().Count());

            // The flattened snapshot is still structurally valid OOXML.
            var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc).ToList();
            Assert.Empty(openXmlErrors);
        }
        finally
        {
            File.Delete(layoutPath);
            File.Delete(boundPath);
            File.Delete(flattenedPath);
        }
    }

    [Fact]
    public void FlattenBindingsForRender_unwraps_row_sdts_so_cloned_rows_are_plain_siblings_of_the_tblHeader_row()
    {
        // Regression for GitHub issue #19: stripping the row-level sdts' binding/repeating properties but
        // KEEPING their shells makes Word fragment the repeater table at every shell — the w:tblHeader
        // header row ends up alone in a one-row table fragment, which can never break across a page, so the
        // header repetition BC renders on every page silently never triggers in the mock. The flatten step
        // must therefore unwrap row-level sdts entirely: every cloned data row a plain w:tr, DIRECT sibling
        // of the header row under one contiguous w:tbl.
        const string RepeaterXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line";
        const string FieldXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:ItemNo_Line[1]";
        var schemaXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<NavWordReportXmlPart xmlns=\"urn:microsoft-dynamics-nav/reports/TestReport/50000/\">"
            + "<BCReportInformation><CreationDateTime>2026-01-01</CreationDateTime></BCReportInformation>"
            + "<Header><CompanyName>Contoso</CompanyName><Line><ItemNo_Line>X</ItemNo_Line></Line></Header>"
            + "</NavWordReportXmlPart>";

        var body = SyntheticLayout.RepeaterTable(RepeaterXPath, FieldXPath, SyntheticLayout.GoodItemId)
            + SyntheticLayout.PlainParagraph("after");
        var layoutPath = SyntheticLayout.Create(body, datasetXml: schemaXml);
        var boundPath = TempDocxPath();
        var flattenedPath = TempDocxPath();

        try
        {
            // The default (non-flattened) logical merge must keep its row-level sdt structure untouched —
            // the unwrap is strictly part of the opt-in render flatten, never of the user-facing merge.
            MergeEngine.Merge(layoutPath, boundPath, new MergeOptions { Seed = 7, Rows = 3 });
            using (var boundDoc = WordprocessingDocument.Open(boundPath, false))
            {
                Assert.NotEmpty(boundDoc.MainDocumentPart!.Document!.Descendants<SdtRow>());
            }

            var result = MergeEngine.Merge(layoutPath, flattenedPath,
                new MergeOptions { Seed = 7, Rows = 3, FlattenBindingsForRender = true });
            Assert.Equal(3, result.Stats.RowsGenerated);

            using var doc = WordprocessingDocument.Open(flattenedPath, false);
            var document = doc.MainDocumentPart!.Document!;
            Assert.Empty(document.Descendants<SdtRow>());

            // One contiguous table: header w:tr first (w:tblHeader intact), then the 3 cloned data rows as
            // DIRECT w:tbl children — nothing row-shaped left for Word to fragment the table at.
            var table = Assert.Single(document.Descendants<Table>());
            Assert.DoesNotContain(table.ChildElements, e => e is SdtElement);
            var directRows = table.Elements<TableRow>().ToList();
            Assert.Equal(4, directRows.Count);
            Assert.NotNull(directRows[0].TableRowProperties?.GetFirstChild<TableHeader>());

            // The unwrap is row-level-only: each cloned cell's own (binding-stripped) field control shell
            // survives, exactly as before.
            Assert.Equal(3, table.Descendants<SdtBlock>().Count());

            var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc).ToList();
            Assert.Empty(openXmlErrors);
        }
        finally
        {
            File.Delete(layoutPath);
            File.Delete(boundPath);
            File.Delete(flattenedPath);
        }
    }

    [Fact]
    public void FlattenBindingsForRender_prunes_a_table_left_rowless_by_a_zero_row_repeater()
    {
        // Degenerate corner of the issue-#19 unwrap: a table whose ONLY content is the repeater (no static
        // header row — a real corpus shape, BC line tables often carry no w:tblHeader), merged against a
        // real override dataset with ZERO matching rows. Unwrapping then leaves a w:tbl with no rows at
        // all — something Word itself never writes — so the flatten must remove the table outright and the
        // result must still validate clean.
        const string RepeaterXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line";
        const string FieldXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:ItemNo_Line[1]";

        var body = SyntheticLayout.RepeaterTable(
                RepeaterXPath, FieldXPath, SyntheticLayout.GoodItemId, headerRow: false)
            + SyntheticLayout.PlainParagraph("after");
        var layoutPath = SyntheticLayout.Create(body);
        var outputPath = TempDocxPath();
        var overridesPath = BuildRowCapOverrideXml(lineCount: 0);

        try
        {
            var result = MergeEngine.Merge(layoutPath, outputPath, new MergeOptions
            {
                DataOverridesPath = overridesPath,
                FlattenBindingsForRender = true,
            });
            Assert.Equal(0, result.Stats.RowsGenerated);

            using var doc = WordprocessingDocument.Open(outputPath, false);
            var document = doc.MainDocumentPart!.Document!;
            Assert.Empty(document.Descendants<SdtRow>());
            Assert.Empty(document.Descendants<Table>());

            var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc).ToList();
            Assert.Empty(openXmlErrors);
        }
        finally
        {
            File.Delete(layoutPath);
            File.Delete(outputPath);
            File.Delete(overridesPath);
        }
    }

    [Fact]
    public void Nested_layout_with_large_rows_is_bounded_by_MaxTotalInstances_and_warns()
    {
        // A 2-level nested repeater at Rows=30 would generate 30 + 30^2 + 30^3 = 27,930 business instances
        // (the rows^depth blow-up). MaxTotalInstances must bound generation AND surface a sample-data-capped
        // warning so the (deliberately partial) preview is never silent.
        const string OuterXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line";
        const string InnerXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:SubLine";
        const string FieldXPath =
            "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:SubLine[1]/ns0:Value[1]";
        var schemaXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<NavWordReportXmlPart xmlns=\"urn:microsoft-dynamics-nav/reports/TestReport/50000/\">"
            + "<BCReportInformation><CreationDateTime>2026-01-01</CreationDateTime></BCReportInformation>"
            + "<Header><CompanyName>Contoso</CompanyName>"
            + "<Line><SubLine><Value>X</Value></SubLine></Line></Header>"
            + "</NavWordReportXmlPart>";

        var body = SyntheticLayout.NestedRepeater(OuterXPath, InnerXPath, FieldXPath, SyntheticLayout.GoodItemId);
        var layoutPath = SyntheticLayout.Create(body, datasetXml: schemaXml);
        var outputPath = TempDocxPath();

        try
        {
            var options = new MergeOptions { Seed = 42, Rows = 30, MaxTotalInstances = 20 };
            var result = MergeEngine.Merge(layoutPath, outputPath, options);

            Assert.Single(result.Warnings, w => w.Kind == "sample-data-capped");
            // Bounded: total rendered rows can never exceed the instance budget, no matter how big Rows is.
            Assert.True(result.Stats.RowsGenerated <= options.MaxTotalInstances,
                $"rows generated ({result.Stats.RowsGenerated}) should be bounded by the budget ({options.MaxTotalInstances})");
        }
        finally
        {
            File.Delete(layoutPath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void Nested_repeater_reanchors_across_two_levels_and_counts_multiply()
    {
        const string OuterXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line";
        const string InnerXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:SubLine";
        const string FieldXPath =
            "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:SubLine[1]/ns0:Value[1]";
        var schemaXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<NavWordReportXmlPart xmlns=\"urn:microsoft-dynamics-nav/reports/TestReport/50000/\">"
            + "<BCReportInformation><CreationDateTime>2026-01-01</CreationDateTime></BCReportInformation>"
            + "<Header><CompanyName>Contoso</CompanyName>"
            + "<Line><SubLine><Value>X</Value></SubLine></Line></Header>"
            + "</NavWordReportXmlPart>";

        var body = SyntheticLayout.NestedRepeater(OuterXPath, InnerXPath, FieldXPath, SyntheticLayout.GoodItemId);
        var layoutPath = SyntheticLayout.Create(body, datasetXml: schemaXml);
        var outputPath = TempDocxPath();

        try
        {
            var options = new MergeOptions { Seed = 909, Rows = 3 };
            var result = MergeEngine.Merge(layoutPath, outputPath, options);

            // 1 outer expansion + 3 inner expansions (one per outer row's own clone of the inner repeater).
            Assert.Equal(4, result.Stats.RepeatersExpanded);
            // 3 outer rows + 9 inner rows (3 per outer row): counts multiply as expected.
            Assert.Equal(12, result.Stats.RowsGenerated);
            Assert.Equal(0, result.Stats.Unresolved);

            var schema = SchemaProvider.FromLayout(layoutPath);
            var expected = SampleDataGenerator.Generate(
                schema, new SampleDataOptions { Seed = options.Seed, Rows = options.Rows });
            XNamespace ns = expected.Namespace;
            var expectedFlat = expected.Xml.Root!.Element(ns + "Header")!.Elements(ns + "Line")
                .SelectMany(line => line.Elements(ns + "SubLine"))
                .Select(sub => sub.Element(ns + "Value")!.Value)
                .ToList();

            using var doc = WordprocessingDocument.Open(outputPath, false);
            var allItems = doc.MainDocumentPart!.Document!.Descendants<SdtElement>()
                .Where(IsRepeaterItemSdt)
                .ToList();

            // Outer-row clones still contain a nested repeatingSectionItem (their own inner rows); the
            // innermost (leaf) clones do not — that split recovers the 3-outer / 9-inner shape.
            var outerRows = allItems.Where(item => item.Descendants<SdtElement>().Any(IsRepeaterItemSdt)).ToList();
            var innerRows = allItems.Where(item => !item.Descendants<SdtElement>().Any(IsRepeaterItemSdt)).ToList();

            Assert.Equal(3, outerRows.Count);
            Assert.Equal(9, innerRows.Count);

            // Document order over the whole body already interleaves correctly (all of outer row 0's inner
            // rows precede outer row 1's), so the leaf items alone give the flat nested-order sequence.
            var actualFlat = innerRows.Select(r => r.Descendants<Text>().First().Text).ToList();
            Assert.Equal(expectedFlat, actualFlat);
        }
        finally
        {
            File.Delete(layoutPath);
            File.Delete(outputPath);
        }
    }

    // ---- MaxRowsPerRepeater (large-repeater robustness cap) ----

    [Fact]
    public void Repeater_with_a_large_rows_value_under_the_default_cap_still_merges_every_row()
    {
        const string RepeaterXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line";
        const string FieldXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:ItemNo_Line[1]";
        var schemaXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<NavWordReportXmlPart xmlns=\"urn:microsoft-dynamics-nav/reports/TestReport/50000/\">"
            + "<BCReportInformation><CreationDateTime>2026-01-01</CreationDateTime></BCReportInformation>"
            + "<Header><CompanyName>Contoso</CompanyName><Line><ItemNo_Line>X</ItemNo_Line></Line></Header>"
            + "</NavWordReportXmlPart>";

        var body = SyntheticLayout.RepeaterWithField(RepeaterXPath, FieldXPath, SyntheticLayout.GoodItemId);
        var layoutPath = SyntheticLayout.Create(body, datasetXml: schemaXml);
        var outputPath = TempDocxPath();

        try
        {
            // The task's own "still works" example: a large-but-comfortably-under-the-default-cap-of-100
            // Rows value must merge every row exactly as an uncapped merge always has - no truncation, no
            // row-cap warning, at the DEFAULT MaxRowsPerRepeater (not set here).
            var options = new MergeOptions { Seed = 111, Rows = 25 };
            var result = MergeEngine.Merge(layoutPath, outputPath, options);

            Assert.Equal(1, result.Stats.RepeatersExpanded);
            Assert.Equal(25, result.Stats.RowsGenerated);
            Assert.Equal(0, result.Stats.Unresolved);
            Assert.DoesNotContain(result.Warnings, w => w.Kind == "row-cap");

            using var doc = WordprocessingDocument.Open(outputPath, false);
            var rows = doc.MainDocumentPart!.Document!.Descendants<SdtElement>().Where(IsRepeaterItemSdt).ToList();
            Assert.Equal(25, rows.Count);
        }
        finally
        {
            File.Delete(layoutPath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void Repeater_exceeding_MaxRowsPerRepeater_is_capped_and_raises_a_row_cap_warning_naming_the_repeater_and_the_cap()
    {
        // Uses DataOverridesPath (a hand-crafted real dataset), NOT generated data: since generation itself
        // now also bounds every repeating item to MaxRowsPerRepeater (SampleDataOptions.MaxRowsPerItem, wired
        // from MergeOptions.MaxRowsPerRepeater - see Repeater_with_generated_data_exceeding_MaxRowsPerRepeater_is_bounded_at_generation_and_still_warns
        // below), a GENERATED dataset can never itself contain more rows than the cap, so this specific
        // per-repeater, clone-time cap/warning can only be exercised end-to-end with a dataset that isn't
        // shaped by generation at all - exactly the real-world case (a real BC export) it remains essential
        // for.
        const string RepeaterXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line";
        const string FieldXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:ItemNo_Line[1]";

        var body = SyntheticLayout.RepeaterWithField(RepeaterXPath, FieldXPath, SyntheticLayout.GoodItemId);
        var layoutPath = SyntheticLayout.Create(body);
        var outputPath = TempDocxPath();
        var overridesPath = BuildRowCapOverrideXml(lineCount: 10);

        try
        {
            const int cap = 5;
            var options = new MergeOptions { MaxRowsPerRepeater = cap, DataOverridesPath = overridesPath };
            var result = MergeEngine.Merge(layoutPath, outputPath, options);

            Assert.Equal(1, result.Stats.RepeatersExpanded);

            // Never silent: RowsGenerated reflects what was ACTUALLY cloned (the cap), not the 10 rows that
            // actually matched the binding.
            Assert.Equal(cap, result.Stats.RowsGenerated);

            var warning = Assert.Single(result.Warnings, w => w.Kind == "row-cap");
            Assert.Contains(RepeaterXPath, warning.Message);
            Assert.Contains(cap.ToString(), warning.Message);
            Assert.Contains("10", warning.Message); // the matched (pre-cap) row count is named too.
            Assert.NotNull(warning.Location);
            Assert.Contains(RepeaterXPath, warning.Location);

            using var doc = WordprocessingDocument.Open(outputPath, false);
            var rows = doc.MainDocumentPart!.Document!.Descendants<SdtElement>().Where(IsRepeaterItemSdt).ToList();
            Assert.Equal(cap, rows.Count);

            var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc).ToList();
            Assert.Empty(openXmlErrors);
        }
        finally
        {
            File.Delete(layoutPath);
            File.Delete(outputPath);
            File.Delete(overridesPath);
        }
    }

    [Fact]
    public void MaxRowsPerRepeater_applies_independently_per_repeating_section_not_as_one_document_wide_total()
    {
        // Uses DataOverridesPath for the same reason as the test above (generation itself now bounds every
        // repeating item to the cap, so a GENERATED dataset could never itself exceed it at either nesting
        // level). Mirrors Nested_repeater_reanchors_across_two_levels_and_counts_multiply's exact fixture
        // shape, but with a hand-built dataset carrying enough rows at both levels to exceed a small cap -
        // proving the cap is re-applied FRESH for EACH repeater expansion (the outer repeater's one
        // expansion, plus each surviving outer row's own separate inner-repeater expansion), never shared as
        // a single document-wide budget across all of them - exactly "a SINGLE repeating section" per the
        // cap's own contract.
        const string OuterXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line";
        const string InnerXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:SubLine";
        const string FieldXPath =
            "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:SubLine[1]/ns0:Value[1]";

        var body = SyntheticLayout.NestedRepeater(OuterXPath, InnerXPath, FieldXPath, SyntheticLayout.GoodItemId);
        var layoutPath = SyntheticLayout.Create(body);
        var outputPath = TempDocxPath();
        const int rowsPerLevel = 6;
        var overridesPath = BuildNestedRowCapOverrideXml(rowsPerLevel);

        try
        {
            const int cap = 3;
            var options = new MergeOptions { MaxRowsPerRepeater = cap, DataOverridesPath = overridesPath };
            var result = MergeEngine.Merge(layoutPath, outputPath, options);

            // 1 outer expansion (6 Line rows matched, capped to 3) + one inner expansion PER SURVIVING outer
            // row (3), each independently matching 6 SubLine rows of its own and independently capping to 3.
            Assert.Equal(1 + cap, result.Stats.RepeatersExpanded);
            Assert.Equal(cap + (cap * cap), result.Stats.RowsGenerated); // 3 outer + 3*3 inner = 12.
            Assert.Equal(0, result.Stats.Unresolved);

            var rowCapWarnings = result.Warnings.Where(w => w.Kind == "row-cap").ToList();
            Assert.Equal(1 + cap, rowCapWarnings.Count);
            Assert.All(rowCapWarnings, w => Assert.Contains(cap.ToString(), w.Message));

            using var doc = WordprocessingDocument.Open(outputPath, false);
            var allItems = doc.MainDocumentPart!.Document!.Descendants<SdtElement>().Where(IsRepeaterItemSdt).ToList();
            var outerRows = allItems.Where(item => item.Descendants<SdtElement>().Any(IsRepeaterItemSdt)).ToList();
            var innerRows = allItems.Where(item => !item.Descendants<SdtElement>().Any(IsRepeaterItemSdt)).ToList();

            Assert.Equal(cap, outerRows.Count);
            Assert.Equal(cap * cap, innerRows.Count);

            var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc).ToList();
            Assert.Empty(openXmlErrors);
        }
        finally
        {
            File.Delete(layoutPath);
            File.Delete(outputPath);
            File.Delete(overridesPath);
        }
    }

    [Fact]
    public void Repeater_with_generated_data_exceeding_MaxRowsPerRepeater_is_bounded_at_generation_and_still_warns()
    {
        // Follow-up hardening: SampleDataGenerator itself now never generates more than MaxRowsPerRepeater
        // instances of any repeating item (SampleDataOptions.MaxRowsPerItem, wired from
        // MergeOptions.MaxRowsPerRepeater) - avoiding the rows^depth generation-time blow-up a merge-only
        // cap could not prevent. Because generation itself never exceeds the cap, the PER-REPEATER,
        // clone-time check (rows.Count > cap) can no longer observe an excess for GENERATED data
        // specifically - so MergeEngine.Merge raises ONE document-level "row-cap" warning instead whenever
        // the REQUESTED Rows exceeds the cap, preserving "no silent caps" end to end.
        const string RepeaterXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line";
        const string FieldXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:ItemNo_Line[1]";
        var schemaXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<NavWordReportXmlPart xmlns=\"urn:microsoft-dynamics-nav/reports/TestReport/50000/\">"
            + "<BCReportInformation><CreationDateTime>2026-01-01</CreationDateTime></BCReportInformation>"
            + "<Header><CompanyName>Contoso</CompanyName><Line><ItemNo_Line>X</ItemNo_Line></Line></Header>"
            + "</NavWordReportXmlPart>";

        var body = SyntheticLayout.RepeaterWithField(RepeaterXPath, FieldXPath, SyntheticLayout.GoodItemId);
        var layoutPath = SyntheticLayout.Create(body, datasetXml: schemaXml);
        var outputPath = TempDocxPath();

        try
        {
            const int cap = 5;
            var options = new MergeOptions { Seed = 444, Rows = 10, MaxRowsPerRepeater = cap };

            // Waste-elimination proof: generation itself (independent of any merge) never produces more
            // than `cap` instances, even though Rows=10 was requested.
            var schema = SchemaProvider.FromLayout(layoutPath);
            var generated = SampleDataGenerator.Generate(
                schema,
                new SampleDataOptions { Seed = options.Seed, Rows = options.Rows, MaxRowsPerItem = options.MaxRowsPerRepeater });
            XNamespace ns = generated.Namespace;
            var generatedLineCount = generated.Xml.Root!.Element(ns + "Header")!.Elements(ns + "Line").Count();
            Assert.Equal(cap, generatedLineCount);

            var result = MergeEngine.Merge(layoutPath, outputPath, options);

            // Final rendered row count is still exactly the cap - "changes no output" holds for the document.
            Assert.Equal(1, result.Stats.RepeatersExpanded);
            Assert.Equal(cap, result.Stats.RowsGenerated);

            // But "no silent caps" still holds too: a document-level warning fires because the REQUESTED
            // Rows (10) exceeded the cap (5), even though no single repeater's own clone-time count ever
            // appeared to exceed it (generation already limited it).
            var warning = Assert.Single(result.Warnings, w => w.Kind == "row-cap");
            Assert.Contains("10", warning.Message);
            Assert.Contains(cap.ToString(), warning.Message);

            using var doc = WordprocessingDocument.Open(outputPath, false);
            var rows = doc.MainDocumentPart!.Document!.Descendants<SdtElement>().Where(IsRepeaterItemSdt).ToList();
            Assert.Equal(cap, rows.Count);
        }
        finally
        {
            File.Delete(layoutPath);
            File.Delete(outputPath);
        }
    }

    /// <summary>Hand-built override dataset (bypasses generation entirely) with <paramref name="lineCount"/> Line rows under one Header.</summary>
    [Fact]
    public void DataOverrides_in_the_report_uis_export_shape_merge_end_to_end()
    {
        // The full path GitHub issue #4 promises: a Send to → XML export (ReportDataSet shape, id matching
        // the layout's report) handed straight to MergeOptions.DataOverridesPath — converted internally to
        // the data-store part shape, repeater expanded one row per DataItem SIBLING, every binding resolved.
        const string RepeaterXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line";
        const string FieldXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:ItemNo_Line[1]";

        var body = SyntheticLayout.RepeaterWithField(RepeaterXPath, FieldXPath, SyntheticLayout.GoodItemId);
        var layoutPath = SyntheticLayout.Create(body);
        var outputPath = TempDocxPath();

        // SyntheticLayout's dataset namespace is urn:microsoft-dynamics-nav/reports/TestReport/50000/.
        var overridesPath = Path.Combine(Path.GetTempPath(), $"bcwl-merge-export-shape-{Guid.NewGuid():N}.xml");
        File.WriteAllText(
            overridesPath,
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
            + "<ReportDataSet name=\"TestReport\" id=\"50000\" language=\"en-US\" formatRegion=\"en-US\" wordMergeDataItem=\"Header\">"
            + "<DataItems><DataItem name=\"Header\">"
            + "<Columns><Column name=\"CompanyName\">Contoso</Column></Columns>"
            + "<DataItems>"
            + "<DataItem name=\"Line\"><Columns><Column name=\"ItemNo_Line\">ITEM-A</Column></Columns></DataItem>"
            + "<DataItem name=\"Line\"><Columns><Column name=\"ItemNo_Line\">ITEM-B</Column></Columns></DataItem>"
            + "</DataItems>"
            + "</DataItem></DataItems>"
            + "</ReportDataSet>");

        try
        {
            var result = MergeEngine.Merge(
                layoutPath, outputPath, new MergeOptions { DataOverridesPath = overridesPath });

            Assert.Equal(0, result.Stats.Unresolved);
            Assert.Equal(1, result.Stats.RepeatersExpanded);
            Assert.Equal(2, result.Stats.RowsGenerated);

            using var doc = WordprocessingDocument.Open(outputPath, false);
            var text = doc.MainDocumentPart!.Document!.InnerText;
            Assert.Contains("ITEM-A", text);
            Assert.Contains("ITEM-B", text);

            var openXmlErrors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc).ToList();
            Assert.Empty(openXmlErrors);
        }
        finally
        {
            File.Delete(layoutPath);
            File.Delete(outputPath);
            File.Delete(overridesPath);
        }
    }

    private static string BuildRowCapOverrideXml(int lineCount)
    {
        XNamespace ns = "urn:microsoft-dynamics-nav/reports/TestReport/50000/";
        var xdoc = new XDocument(
            new XElement(
                ns + "NavWordReportXmlPart",
                new XElement(
                    ns + "Header",
                    new XElement(ns + "CompanyName", "Contoso"),
                    Enumerable.Range(1, lineCount).Select(i => new XElement(ns + "Line", new XElement(ns + "ItemNo_Line", $"ITEM-{i:D2}"))))));

        var path = Path.Combine(Path.GetTempPath(), $"bcwl-merge-rowcap-override-{Guid.NewGuid():N}.xml");
        xdoc.Save(path);
        return path;
    }

    /// <summary>
    /// Hand-built override dataset (bypasses generation entirely): one Header with <paramref name="rowsPerLevel"/>
    /// Line rows, EACH with its own <paramref name="rowsPerLevel"/> SubLine rows - for proving the merge-time
    /// cap re-applies independently at both nesting levels.
    /// </summary>
    private static string BuildNestedRowCapOverrideXml(int rowsPerLevel)
    {
        XNamespace ns = "urn:microsoft-dynamics-nav/reports/TestReport/50000/";
        var xdoc = new XDocument(
            new XElement(
                ns + "NavWordReportXmlPart",
                new XElement(
                    ns + "Header",
                    new XElement(ns + "CompanyName", "Contoso"),
                    Enumerable.Range(1, rowsPerLevel).Select(i =>
                        new XElement(
                            ns + "Line",
                            Enumerable.Range(1, rowsPerLevel).Select(j =>
                                new XElement(ns + "SubLine", new XElement(ns + "Value", $"L{i}S{j}"))))))));

        var path = Path.Combine(Path.GetTempPath(), $"bcwl-merge-rowcap-nested-override-{Guid.NewGuid():N}.xml");
        xdoc.Save(path);
        return path;
    }

    [Fact]
    public void Unresolved_binding_shows_placeholder_and_raises_warning()
    {
        const string BogusXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:BogusField[1]";

        var body = SyntheticLayout.BoundField(BogusXPath, SyntheticLayout.GoodItemId);
        var layoutPath = SyntheticLayout.Create(body);
        var outputPath = TempDocxPath();

        try
        {
            var result = MergeEngine.Merge(layoutPath, outputPath);

            Assert.Equal(0, result.Stats.FieldsFilled);
            Assert.Equal(1, result.Stats.Unresolved);

            var warning = Assert.Single(result.Warnings, w => w.Kind == "unresolved-binding");
            Assert.Contains("BogusField", warning.Message);
            Assert.NotNull(warning.Location);

            using var doc = WordprocessingDocument.Open(outputPath, false);
            var sdt = doc.MainDocumentPart!.Document!.Descendants<SdtElement>().Single();
            var actualValue = sdt.Descendants<Text>().First().Text;

            Assert.Equal("«BogusField?»", actualValue);
        }
        finally
        {
            File.Delete(layoutPath);
            File.Delete(outputPath);
        }
    }

    [Theory]
    [InlineData(Corpus.SalesInvoice)]
    [InlineData(Corpus.InventoryOrderDetails)]
    [InlineData(Corpus.StandardStatement)]
    public void Corpus_layouts_merge_and_validate_with_zero_errors(string fileName)
    {
        var outputPath = TempDocxPath();
        try
        {
            // Rows=3 (not 1) - deliberately clones every repeater more than once, including SalesInvoice's
            // PaymentServiceLogo picture and StandardStatement's CompanyPicture, each living inside a body
            // repeater: MergeEngine now regenerates each cloned row's wp:docPr/bookmark ids,
            // so this is a genuine regression guard for that fix, not just a smoke test that happens to
            // avoid it.
            var result = MergeEngine.Merge(
                Corpus.Path(fileName), outputPath, new MergeOptions { Seed = 12345, Rows = 3 });

            Assert.True(result.Stats.RepeatersExpanded > 0, "expected at least one repeater to expand");
            Assert.True(result.Stats.RowsGenerated > 0, "expected at least one generated row");

            // Harden against a silent re-anchoring regression: if re-anchoring broke, bindings would
            // still "resolve" against the WRONG (e.g. always-first) node in many cases rather than
            // raising anything, so the strongest guard available here is that nothing needed to fall
            // back or go unresolved at all on these known-good corpus files.
            Assert.Equal(0, result.Stats.Unresolved);
            Assert.True(result.Stats.FieldsFilled > 0, "expected at least one field to fill");
            Assert.DoesNotContain(result.Warnings, w => w.Kind == "unresolved-binding" || w.Kind == "xpath-fallback");

            using var doc = WordprocessingDocument.Open(outputPath, false);
            var errors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc).ToList();

            Assert.True(errors.Count == 0,
                "expected zero validation errors; found: "
                + string.Join(" | ", errors.Select(e => $"{e.Path?.XPath}: {e.Description}")));
        }
        finally
        {
            File.Delete(outputPath);
        }
    }

    [Theory]
    // SalesInvoiceForSubscriptionBilling's one repeater-nested picture (PaymentServiceLogo, inside the
    // PaymentReportingArgument repeater) expands once per generated row (Rows=2) on top of the one
    // non-repeated header CompanyPicture: 2 + 1 = 3.
    [InlineData(Corpus.SalesInvoice, 3)]
    // InventoryOrderDetails has no picture controls at all, so it cannot exercise this theory (which
    // asserts a placeholder blip was actually resolved) - StandardPurchaseOrder.docx has exactly one
    // (header-only, non-repeated) picture control and merges clean instead.
    [InlineData("StandardPurchaseOrder.docx", 1)]
    [InlineData(Corpus.StandardStatement, 2)]
    public void Corpus_picture_controls_are_filled_with_a_valid_placeholder_png(string fileName, int expectedPicturesFilled)
    {
        var outputPath = TempDocxPath();
        try
        {
            var result = MergeEngine.Merge(
                Corpus.Path(fileName), outputPath, new MergeOptions { Seed = 12345, Rows = 2 });

            Assert.Equal(expectedPicturesFilled, result.Stats.PicturesFilled);
            Assert.DoesNotContain(result.Warnings, w => w.Kind == "picture-no-blip");

            using var doc = WordprocessingDocument.Open(outputPath, false);

            ImagePart? imagePart = null;
            foreach (var (part, root) in DocumentParts(doc))
            {
                var blip = root?.Descendants<Blip>().FirstOrDefault(b => !string.IsNullOrEmpty(b.Embed?.Value));
                if (blip is not null)
                {
                    imagePart = Assert.IsType<ImagePart>(part.GetPartById(blip.Embed!.Value!));
                    break;
                }
            }

            // At least one picture blip must resolve to the new placeholder — never the original 10-byte
            // stub, which is neither PNG-signed nor even this long.
            Assert.NotNull(imagePart);
            Assert.Equal("image/png", imagePart!.ContentType);

            using var stream = imagePart.GetStream(FileMode.Open, FileAccess.Read);
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            var bytes = buffer.ToArray();

            Assert.True(bytes.Length > 10,
                $"expected placeholder image bytes to exceed the original 10-byte stub, was {bytes.Length}");
            Assert.Equal(PngSignature, bytes.Take(PngSignature.Length).ToArray());
        }
        finally
        {
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void Repeater_row_with_a_picture_and_a_bookmark_gets_distinct_ids_per_clone_and_validates_clean()
    {
        // Regression: a repeater row containing a wp:docPr (picture) and a bookmarkStart/End pair used to
        // leave every cloned row's ids identical - a real Sem_UniqueAttributeValue OOXML validation error
        // once Rows > 1. This is the focused regression test for the fix (see also the corpus-level
        // Corpus_layouts_merge_and_validate_with_zero_errors and Corpus_picture_controls_are_filled_with_a_
        // valid_placeholder_png theories above, which now run at Rows=3 for the same reason).
        // Exactly ONE <Line> in the schema XML (not three) - matching every other synthetic-repeater test's
        // convention (e.g. Repeater_expands_configured_row_count_with_reanchored_per_row_values above):
        // SchemaProvider.BuildNode adds one DataItem per XElement occurrence rather than deduplicating by
        // tag name, so 3 sibling <Line> elements here would build 3 SEPARATE (duplicate) "Line" DataItem
        // nodes, each independently generating Rows instances - 9 total, not 3. The generated ROW COUNT is
        // controlled entirely by MergeOptions.Rows, not by how many elements the schema sample happens to
        // contain.
        const string RepeaterXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line";
        var schemaXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<NavWordReportXmlPart xmlns=\"urn:microsoft-dynamics-nav/reports/TestReport/50000/\">"
            + "<BCReportInformation><CreationDateTime>2026-01-01</CreationDateTime></BCReportInformation>"
            + "<Header><CompanyName>Contoso</CompanyName><Line><ItemNo_Line>A</ItemNo_Line></Line></Header>"
            + "</NavWordReportXmlPart>";

        var body = SyntheticLayout.RepeaterWithPictureAndBookmark(RepeaterXPath, SyntheticLayout.GoodItemId);
        var layoutPath = SyntheticLayout.Create(body, datasetXml: schemaXml);
        var outputPath = TempDocxPath();

        try
        {
            var options = new MergeOptions { Seed = 1, Rows = 3 };
            var result = MergeEngine.Merge(layoutPath, outputPath, options);

            Assert.Equal(1, result.Stats.RepeatersExpanded);
            Assert.Equal(3, result.Stats.RowsGenerated);
            Assert.Equal(3, result.Stats.PicturesFilled);

            using var doc = WordprocessingDocument.Open(outputPath, false);

            // ---- zero OpenXmlValidator errors: the actual acceptance criterion ----
            var errors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc).ToList();
            Assert.True(errors.Count == 0,
                "expected zero validation errors; found: "
                + string.Join(" | ", errors.Select(e => $"{e.Path?.XPath}: {e.Description}")));

            var body2 = doc.MainDocumentPart!.Document!.Body!;

            // ---- direct assertion: every cloned row's wp:docPr id is distinct (not just "validator-clean") ----
            var docPrIds = body2.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>()
                .Select(d => d.Id?.Value)
                .ToList();
            Assert.Equal(3, docPrIds.Count);
            Assert.Equal(3, docPrIds.Distinct().Count());

            // ---- bookmarkStart/End ids are distinct, and each pair still matches (start id == end id) ----
            var bookmarkStarts = body2.Descendants<BookmarkStart>().ToList();
            var bookmarkEnds = body2.Descendants<BookmarkEnd>().ToList();
            Assert.Equal(3, bookmarkStarts.Count);
            Assert.Equal(3, bookmarkEnds.Count);

            var startIds = bookmarkStarts.Select(b => b.Id?.Value).ToList();
            Assert.Equal(3, startIds.Distinct().Count());
            Assert.Equal(startIds.OrderBy(x => x), bookmarkEnds.Select(b => b.Id?.Value).OrderBy(x => x));

            // ---- bookmark names: first clone keeps the original name; later clones get a distinct suffix ----
            var names = bookmarkStarts.Select(b => b.Name?.Value).ToList();
            Assert.Contains("TestBookmark", names);
            Assert.Equal(3, names.Distinct().Count());
            Assert.All(names.Where(n => n != "TestBookmark"), n => Assert.StartsWith("TestBookmark_r", n));
        }
        finally
        {
            File.Delete(layoutPath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void Bookmark_range_spanning_the_repeater_boundary_gets_distinct_orphaned_end_ids_per_clone_and_validates_clean()
    {
        // B32 follow-up (Opus review): a bookmarkStart BEFORE the repeater paired (by id) with a
        // bookmarkEnd INSIDE the row template used to leave every clone's bookmarkEnd with the SAME
        // (unmatched) id - RegenerateClonedIds only ever walked the clone's OWN BookmarkStart elements to
        // find a partner, so an end whose partner lives outside the clone was never reached at all.
        // Rows=3 with the fix disabled reproduced exactly the reviewer's repro: bookmarkEnd ids [0,0,0],
        // one Sem_UniqueAttributeValue error.
        const string RepeaterXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line";
        const string BookmarkId = "0";
        var schemaXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<NavWordReportXmlPart xmlns=\"urn:microsoft-dynamics-nav/reports/TestReport/50000/\">"
            + "<BCReportInformation><CreationDateTime>2026-01-01</CreationDateTime></BCReportInformation>"
            + "<Header><CompanyName>Contoso</CompanyName><Line><ItemNo_Line>A</ItemNo_Line></Line></Header>"
            + "</NavWordReportXmlPart>";

        var body =
            SyntheticLayout.BookmarkStartOnly("Outside", BookmarkId)
            + SyntheticLayout.RepeaterWithBookmarkEndOnly(RepeaterXPath, SyntheticLayout.GoodItemId, BookmarkId);
        var layoutPath = SyntheticLayout.Create(body, datasetXml: schemaXml);
        var outputPath = TempDocxPath();

        try
        {
            var options = new MergeOptions { Seed = 1, Rows = 3 };
            var result = MergeEngine.Merge(layoutPath, outputPath, options);

            Assert.Equal(1, result.Stats.RepeatersExpanded);
            Assert.Equal(3, result.Stats.RowsGenerated);

            using var doc = WordprocessingDocument.Open(outputPath, false);

            // ---- zero OpenXmlValidator errors: the actual acceptance criterion ----
            var errors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc).ToList();
            Assert.True(errors.Count == 0,
                "expected zero validation errors; found: "
                + string.Join(" | ", errors.Select(e => $"{e.Path?.XPath}: {e.Description}")));

            var body2 = doc.MainDocumentPart!.Document!.Body!;

            // ---- the outside bookmarkStart is untouched: exactly one, still carrying its original id ----
            var starts = body2.Descendants<BookmarkStart>().ToList();
            Assert.Single(starts);
            Assert.Equal("Outside", starts[0].Name?.Value);
            Assert.Equal(BookmarkId, starts[0].Id?.Value);

            // ---- direct assertion: all 3 cloned bookmarkEnd ids are distinct (not just "validator-clean") ----
            var endIds = body2.Descendants<BookmarkEnd>().Select(e => e.Id?.Value).ToList();
            Assert.Equal(3, endIds.Count);
            Assert.Equal(3, endIds.Distinct().Count());

            // Deliberately orphaned: none of the 3 (renumbered) ends still matches the outside start's id.
            Assert.DoesNotContain(BookmarkId, endIds);
        }
        finally
        {
            File.Delete(layoutPath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void PlaceholderImage_bytes_are_a_valid_png()
    {
        var bytes = PlaceholderImage.PngBytes;

        Assert.True(bytes.Length is > 16 and < 10_000,
            $"expected a plausible, non-trivial PNG byte length, was {bytes.Length}");
        Assert.Equal(PngSignature, bytes.Take(PngSignature.Length).ToArray());
        Assert.True(ContainsAsciiToken(bytes, "IHDR"), "expected an IHDR chunk");
        Assert.True(ContainsAsciiToken(bytes, "IDAT"), "expected an IDAT chunk");
        Assert.True(ContainsAsciiToken(bytes, "IEND"), "expected an IEND chunk");
    }

    [Fact]
    public void Merging_the_same_corpus_file_twice_is_deterministic()
    {
        var output1 = TempDocxPath();
        var output2 = TempDocxPath();

        try
        {
            var options = new MergeOptions { Seed = 777, Rows = 2 };
            MergeEngine.Merge(Corpus.Path(Corpus.SalesInvoice), output1, options);
            MergeEngine.Merge(Corpus.Path(Corpus.SalesInvoice), output2, options);

            static string DocumentXmlText(string path)
            {
                using var doc = WordprocessingDocument.Open(path, false);
                return doc.MainDocumentPart!.Document!.OuterXml;
            }

            // The OpenXml SDK's AddNewPart<ImagePart>() (MergeEngine.GetOrCreatePlaceholderImagePart, used
            // the first time a picture control in a given part needs the shared placeholder image part)
            // mints a fresh, non-deterministic relationship id on every call - confirmed directly against
            // the SDK - even for byte-identical input across separate calls. That is real SDK behavior, not
            // a MergeEngine defect (the merged document is correct either way: the blip still resolves to a
            // valid embedded placeholder PNG), so normalize r:embed values away before comparing; this
            // corpus layout's PaymentServiceLogo picture control lives in document.xml itself (nested in a
            // repeater), unlike a header/footer-only picture, so it would otherwise make this assertion flaky.
            static string NormalizeVolatileRelationshipIds(string xml) =>
                System.Text.RegularExpressions.Regex.Replace(xml, "r:embed=\"[^\"]*\"", "r:embed=\"R_NORMALIZED\"");

            Assert.Equal(
                NormalizeVolatileRelationshipIds(DocumentXmlText(output1)),
                NormalizeVolatileRelationshipIds(DocumentXmlText(output2)));
        }
        finally
        {
            File.Delete(output1);
            File.Delete(output2);
        }
    }

    [Fact]
    public void Merge_throws_when_outputPath_is_the_same_file_as_layoutPath()
    {
        var body = SyntheticLayout.BoundField(
            "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:CompanyName[1]", SyntheticLayout.GoodItemId);
        var layoutPath = SyntheticLayout.Create(body);

        try
        {
            var directory = Path.GetDirectoryName(layoutPath)!;
            var fileName = Path.GetFileName(layoutPath);

            // Same file as layoutPath once Path.GetFullPath collapses the redundant "." segment — a
            // plain string comparison would miss this, which is exactly why the guard normalizes first.
            var sameFileDifferentSpelling = Path.Combine(directory, ".", fileName);
            Assert.NotEqual(layoutPath, sameFileDifferentSpelling);

            var ex = Assert.Throws<ArgumentException>(
                () => MergeEngine.Merge(layoutPath, sameFileDifferentSpelling));
            Assert.Equal("outputPath", ex.ParamName);
        }
        finally
        {
            File.Delete(layoutPath);
        }
    }

    [Fact]
    public void Merge_success_overwrites_a_preexisting_output_file_and_leaves_no_staging_leftovers()
    {
        // the merge now stages into a throwaway copy in outputPath's own directory and only replaces
        // outputPath (via File.Move(overwrite:true)) once the whole merge succeeds. Prove the happy path
        // still overwrites an existing outputPath correctly and leaves no `.bcwl-merge-stage-*` behind.
        var body = SyntheticLayout.BoundField(
            "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:CompanyName[1]", SyntheticLayout.GoodItemId);
        var layoutPath = SyntheticLayout.Create(body);

        // An isolated directory (not the shared temp root) so the "no staging leftovers" check below can't
        // be confused by another test's concurrent merge into the same shared temp directory.
        var outputDir = Path.Combine(Path.GetTempPath(), $"bcwl-merge-overwrite-{Guid.NewGuid():N}");
        Directory.CreateDirectory(outputDir);
        var outputPath = Path.Combine(outputDir, "output.docx");
        File.WriteAllBytes(outputPath, "stale output from a previous run"u8.ToArray());

        try
        {
            var result = MergeEngine.Merge(layoutPath, outputPath, new MergeOptions { Seed = 1, Rows = 1 });

            Assert.Equal(1, result.Stats.FieldsFilled);
            using (var doc = WordprocessingDocument.Open(outputPath, false))
            {
                Assert.NotNull(doc.MainDocumentPart?.Document);
            }

            Assert.Empty(Directory.GetFiles(outputDir, ".bcwl-merge-stage-*"));
        }
        finally
        {
            File.Delete(layoutPath);
            Directory.Delete(outputDir, recursive: true);
        }
    }

    [Fact]
    public void Merge_sweeps_a_stale_orphaned_merge_stage_file_in_outputPaths_directory()
    {
        // Opus review of B21 (interaction): a hard kill between the merge's own stage-copy and
        // its commit rename leaves a `.bcwl-merge-stage-*.docx` behind in outputPath's directory - the
        // exact same orphan-on-crash failure mode already handled for ToolGuards' edit-commit staging file,
        // just for the merge pipeline's own staging file instead. Prove a LATER merge call sweeps a stale
        // one away.
        var body = SyntheticLayout.BoundField(
            "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:CompanyName[1]", SyntheticLayout.GoodItemId);
        var layoutPath = SyntheticLayout.Create(body);

        var outputDir = Path.Combine(Path.GetTempPath(), $"bcwl-merge-stage-sweep-{Guid.NewGuid():N}");
        Directory.CreateDirectory(outputDir);
        var outputPath = Path.Combine(outputDir, "output.docx");
        // CreationTimeUtc, not LastWriteTimeUtc, is the sweep's age signal: File.Copy preserves the
        // SOURCE's last-write time onto a freshly-copied file, so a genuinely brand-new staging file
        // copied from an old layout would already look stale by last-write time alone - only creation
        // time reliably reflects how long the STAGING FILE ITSELF has existed.
        var staleMergeStage = Path.Combine(outputDir, $".bcwl-merge-stage-{Guid.NewGuid():N}.docx");
        File.WriteAllText(staleMergeStage, "orphaned mid-merge artifact from a hard-killed process");
        File.SetCreationTimeUtc(staleMergeStage, DateTime.UtcNow - TimeSpan.FromDays(2));

        try
        {
            var result = MergeEngine.Merge(layoutPath, outputPath, new MergeOptions { Seed = 1, Rows = 1 });

            Assert.Equal(1, result.Stats.FieldsFilled);
            Assert.False(
                File.Exists(staleMergeStage), "a stale .bcwl-merge-stage-* file must be swept by a later merge call");
        }
        finally
        {
            File.Delete(layoutPath);
            Directory.Delete(outputDir, recursive: true);
        }
    }

    [Fact]
    public void Merge_never_sweeps_a_fresh_merge_stage_file_in_outputPaths_directory()
    {
        // Symmetric guard: a FRESH `.bcwl-merge-stage-*.docx` (well inside the retention window - the
        // shape a genuinely concurrent in-flight merge into a DIFFERENT output in the same directory would
        // have) must never be swept, so a live merge can never be mistaken for an orphan.
        var body = SyntheticLayout.BoundField(
            "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:CompanyName[1]", SyntheticLayout.GoodItemId);
        var layoutPath = SyntheticLayout.Create(body);

        var outputDir = Path.Combine(Path.GetTempPath(), $"bcwl-merge-stage-sweep-fresh-{Guid.NewGuid():N}");
        Directory.CreateDirectory(outputDir);
        var outputPath = Path.Combine(outputDir, "output.docx");
        var freshMergeStage = Path.Combine(outputDir, $".bcwl-merge-stage-{Guid.NewGuid():N}.docx");
        File.WriteAllText(freshMergeStage, "a live in-flight merge's staged file");
        // No SetCreationTimeUtc: just-created, well inside the retention window.

        try
        {
            var result = MergeEngine.Merge(layoutPath, outputPath, new MergeOptions { Seed = 1, Rows = 1 });

            Assert.Equal(1, result.Stats.FieldsFilled);
            Assert.True(File.Exists(freshMergeStage), "a fresh .bcwl-merge-stage-* file must survive a merge call");
        }
        finally
        {
            File.Delete(layoutPath);
            Directory.Delete(outputDir, recursive: true);
        }
    }

    [Fact]
    public void Internal_overload_merges_a_hand_built_dataset_against_an_open_document()
    {
        const string Ns = "urn:microsoft-dynamics-nav/reports/TestReport/50000/";
        const string FieldXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:CompanyName[1]";
        const string RepeaterXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line";
        const string InnerFieldXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line[1]/ns0:ItemNo_Line[1]";

        var body =
            SyntheticLayout.BoundField(FieldXPath, SyntheticLayout.GoodItemId)
            + SyntheticLayout.RepeaterWithField(RepeaterXPath, InnerFieldXPath, SyntheticLayout.GoodItemId);
        var layoutPath = SyntheticLayout.Create(body);

        // Hand-built dataset — deliberately bypasses SchemaProvider/SampleDataGenerator entirely, so this
        // test exercises the internal overload's own merge logic in isolation from the generator.
        XNamespace ns = Ns;
        var xml = new XDocument(
            new XElement(
                ns + "NavWordReportXmlPart",
                new XElement(
                    ns + "Header",
                    new XElement(ns + "CompanyName", "Hand Built Co"),
                    new XElement(ns + "Line", new XElement(ns + "ItemNo_Line", "ITEM-0001")),
                    new XElement(ns + "Line", new XElement(ns + "ItemNo_Line", "ITEM-0002")))));
        var dataset = new SampleDataset { Xml = xml, Namespace = Ns };

        try
        {
            using (var doc = WordprocessingDocument.Open(layoutPath, true))
            {
                var result = MergeEngine.Merge(doc, dataset);

                // 1 standalone field + 1 inner field per cloned row (2 rows) = 3 fields filled in total.
                Assert.Equal(3, result.Stats.FieldsFilled);
                Assert.Equal(1, result.Stats.RepeatersExpanded);
                Assert.Equal(2, result.Stats.RowsGenerated);
                Assert.Equal(0, result.Stats.Unresolved);

                doc.MainDocumentPart!.Document!.Save();
            }

            using var reopened = WordprocessingDocument.Open(layoutPath, false);
            var texts = reopened.MainDocumentPart!.Document!.Descendants<Text>().Select(t => t.Text).ToList();

            Assert.Contains("Hand Built Co", texts);
            Assert.Contains("ITEM-0001", texts);
            Assert.Contains("ITEM-0002", texts);
        }
        finally
        {
            File.Delete(layoutPath);
        }
    }
}
