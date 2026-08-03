using System.Globalization;
using System.Text;
using System.Xml.Linq;
using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using BcWordLayout.Merge;

namespace BcWordLayout.Tests;

public class SampleDataGeneratorTests
{
    private static DatasetTree MinimalSchema()
    {
        var root = new DataItem { Name = "NavWordReportXmlPart", Path = "/" };
        return new DatasetTree
        {
            Report = new ReportIdentity
            {
                ReportName = "Test",
                ReportId = "1",
                Namespace = "urn:microsoft-dynamics-nav/reports/Test/1/",
            },
            Root = root,
        };
    }

    private static DatasetTree SchemaWithHeaderColumns(params string[] columnNames)
    {
        var schema = MinimalSchema();
        var header = new DataItem { Name = "Header", Path = "/Header" };
        foreach (var name in columnNames)
        {
            header.Columns.Add(new DatasetColumn { Name = name, Path = $"/Header/{name}" });
        }

        schema.Root.Children.Add(header);
        return schema;
    }

    [Fact]
    public void Same_seed_produces_byte_identical_xml_and_different_seed_differs()
    {
        var schema = SchemaProvider.FromLayout(Corpus.Path(Corpus.SalesInvoice));

        var a = SampleDataGenerator.Generate(schema, new SampleDataOptions { Seed = 42, Rows = 2 });
        var b = SampleDataGenerator.Generate(schema, new SampleDataOptions { Seed = 42, Rows = 2 });
        var c = SampleDataGenerator.Generate(schema, new SampleDataOptions { Seed = 43, Rows = 2 });

        Assert.Equal(a.Xml.ToString(), b.Xml.ToString());
        Assert.NotEqual(a.Xml.ToString(), c.Xml.ToString());
    }

    [Fact]
    public void Rows_option_controls_business_item_instance_counts_recursively()
    {
        var schema = SchemaProvider.FromLayout(Corpus.Path(Corpus.SalesInvoice));
        const int rows = 3;
        var dataset = SampleDataGenerator.Generate(schema, new SampleDataOptions { Seed = 7, Rows = rows });

        XNamespace ns = schema.Report.Namespace;
        var root = dataset.Xml.Root!;

        Assert.Equal("NavWordReportXmlPart", root.Name.LocalName);
        Assert.Equal(schema.Report.Namespace, root.Name.NamespaceName);

        // Exactly 1 BCReportInformation (system subtree), regardless of Rows.
        Assert.Single(root.Elements(ns + "BCReportInformation"));

        // Header is a business data item directly under the root: exactly `rows` instances.
        var headers = root.Elements(ns + "Header").ToList();
        Assert.Equal(rows, headers.Count);

        // Each Header instance gets its OWN independently generated set of Line children (rows per Header,
        // not rows total) — proves the recursive "own instance set" rule, not just top-level counting.
        foreach (var header in headers)
        {
            Assert.Equal(rows, header.Elements(ns + "Line").Count());
        }
    }

    [Fact]
    public void Label_column_renders_humanized_caption_without_lbl_suffix()
    {
        var schema = SchemaWithHeaderColumns("TotalAmountLbl");
        var dataset = SampleDataGenerator.Generate(schema, new SampleDataOptions { Seed = 1, Rows = 1 });

        XNamespace ns = schema.Report.Namespace;
        var value = dataset.Xml.Root!.Element(ns + "Header")!.Element(ns + "TotalAmountLbl")!.Value;

        Assert.Equal("Total Amount", value);
        Assert.DoesNotContain("Lbl", value);
    }

    [Fact]
    public void Amount_and_date_columns_are_type_aware()
    {
        var schema = SchemaWithHeaderColumns("TotalAmount", "DocumentDate");
        var dataset = SampleDataGenerator.Generate(schema, new SampleDataOptions { Seed = 5, Rows = 1 });

        XNamespace ns = schema.Report.Namespace;
        var header = dataset.Xml.Root!.Element(ns + "Header")!;

        var amountText = header.Element(ns + "TotalAmount")!.Value;
        Assert.True(
            decimal.TryParse(amountText, NumberStyles.Number, CultureInfo.InvariantCulture, out _),
            $"'{amountText}' should parse as an invariant-culture decimal");
        Assert.Contains('.', amountText);

        var dateText = header.Element(ns + "DocumentDate")!.Value;
        Assert.True(
            DateTime.TryParseExact(dateText, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out _),
            $"'{dateText}' should match dd/MM/yyyy");
    }

    [Theory]
    // Abbreviated/alternate spellings from the 2026-07-31 corpus preview sweep: before these tokens were
    // recognized, every such column fell through to name-echo fallback text.
    [InlineData("OriginalAmt_CustLedgEntry2", "decimal")]
    [InlineData("AgingBandBufCol1Amt", "decimal")]
    [InlineData("CustBalance_CustLedgEntryHdr", "decimal")]
    [InlineData("LineDisc_PurchLine", "decimal")]
    [InlineData("ExptRcptDt_PurchHeader", "date")]
    [InlineData("TodayFormatted", "date")]
    [InlineData("ContractBillingDetailsDays", "integer")]
    public void Abbreviated_type_spellings_get_typed_samples(string columnName, string expectedKind)
    {
        var schema = SchemaWithHeaderColumns(columnName);
        var dataset = SampleDataGenerator.Generate(schema, new SampleDataOptions { Seed = 11, Rows = 1 });

        XNamespace ns = schema.Report.Namespace;
        var value = dataset.Xml.Root!.Element(ns + "Header")!.Element(ns + columnName)!.Value;

        switch (expectedKind)
        {
            case "decimal":
                Assert.True(
                    decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out _),
                    $"'{value}' should parse as an invariant-culture decimal");
                Assert.Contains('.', value);
                break;
            case "date":
                Assert.True(
                    DateTime.TryParseExact(value, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out _),
                    $"'{value}' should match dd/MM/yyyy");
                break;
            case "integer":
                Assert.True(int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out _),
                    $"'{value}' should parse as a plain integer");
                break;
            default:
                Assert.Fail($"unknown expectation '{expectedKind}'");
                break;
        }
    }

    [Theory]
    // Caption-indicator words must trump type tokens (2026-07-31 corpus preview sweep): these are real
    // column names from BC's standard layouts whose header/totals cells previously rendered as dates,
    // DOCN- codes, and bare decimals.
    [InlineData("DueDateCaption", "Due Date")]
    [InlineData("DocNoCaption", "Doc No")]
    [InlineData("DescCustLedgEntry2Caption", "Desc Cust Ledg Entry 2")]
    [InlineData("TotalText", "Total")]
    [InlineData("VATAmountLbl_Header", "VAT Amount Header")]
    public void Caption_indicator_columns_get_constant_caption_text_not_typed_samples(
        string columnName, string expected)
    {
        var schema = SchemaWithHeaderColumns(columnName);
        var dataset = SampleDataGenerator.Generate(schema, new SampleDataOptions { Seed = 3, Rows = 1 });

        XNamespace ns = schema.Report.Namespace;
        var value = dataset.Xml.Root!.Element(ns + "Header")!.Element(ns + columnName)!.Value;

        Assert.Equal(expected, value);
    }

    [Theory]
    // Precedence fixes from the 2026-07-31 corpus preview sweep: identifiers, flags, and percentages
    // whose names ALSO contain an amount token must not be sampled as free-range decimals.
    [InlineData("CompanyVATRegistrationNo")]
    [InlineData("CompanyGiroNo")]
    [InlineData("VATIdentifier_VATCounter")]
    public void Identifier_columns_with_amount_tokens_still_get_code_samples(string columnName)
    {
        var schema = SchemaWithHeaderColumns(columnName);
        var dataset = SampleDataGenerator.Generate(schema, new SampleDataOptions { Seed = 8, Rows = 1 });

        XNamespace ns = schema.Report.Namespace;
        var value = dataset.Xml.Root!.Element(ns + "Header")!.Element(ns + columnName)!.Value;

        Assert.Matches(@"^[A-Z]{1,4}-\d{4,}$", value);
    }

    [Fact]
    public void Phone_number_columns_keep_phone_samples_despite_the_No_token()
    {
        var schema = SchemaWithHeaderColumns("CompanyPhoneNo");
        var dataset = SampleDataGenerator.Generate(schema, new SampleDataOptions { Seed = 8, Rows = 1 });

        XNamespace ns = schema.Report.Namespace;
        var value = dataset.Xml.Root!.Element(ns + "Header")!.Element(ns + "CompanyPhoneNo")!.Value;

        Assert.StartsWith("+1-555-", value);
    }

    [Theory]
    [InlineData("PricesInclVAT_SalesHeader")]
    [InlineData("ShowShippingAddr")]
    [InlineData("IsServiceContract")]
    public void Boolean_flag_columns_get_yes_no_samples(string columnName)
    {
        var schema = SchemaWithHeaderColumns(columnName);
        var dataset = SampleDataGenerator.Generate(schema, new SampleDataOptions { Seed = 8, Rows = 1 });

        XNamespace ns = schema.Report.Namespace;
        var value = dataset.Xml.Root!.Element(ns + "Header")!.Element(ns + columnName)!.Value;

        Assert.Contains(value, new[] { "Yes", "No" });
    }

    [Theory]
    [InlineData("VATPct_Line")]
    [InlineData("LineDiscountPercent")]
    public void Percent_columns_sample_within_0_to_100(string columnName)
    {
        var schema = SchemaWithHeaderColumns(columnName);

        // Several seeds/rows so an out-of-range generator cannot pass by luck.
        foreach (var seed in new[] { 1, 7, 42, 99 })
        {
            var dataset = SampleDataGenerator.Generate(schema, new SampleDataOptions { Seed = seed, Rows = 3 });
            XNamespace ns = schema.Report.Namespace;
            foreach (var header in dataset.Xml.Root!.Elements(ns + "Header"))
            {
                var value = decimal.Parse(header.Element(ns + columnName)!.Value, CultureInfo.InvariantCulture);
                Assert.InRange(value, 0m, 100m);
            }
        }
    }

    [Theory]
    [InlineData("UnitOfMeasure_Line")]
    [InlineData("UOMCode_Line")]
    public void Unit_of_measure_columns_get_short_codes(string columnName)
    {
        var schema = SchemaWithHeaderColumns(columnName);
        var dataset = SampleDataGenerator.Generate(schema, new SampleDataOptions { Seed = 4, Rows = 1 });

        XNamespace ns = schema.Report.Namespace;
        var value = dataset.Xml.Root!.Element(ns + "Header")!.Element(ns + columnName)!.Value;

        // Short realistic code, never the ~25-char fallback that wrapped letter-by-letter in previews.
        Assert.Matches("^[A-Z]{2,4}$", value);
    }

    [Fact]
    public void Unit_price_stays_an_amount_despite_the_Unit_token()
    {
        var schema = SchemaWithHeaderColumns("UnitPrice_Line");
        var dataset = SampleDataGenerator.Generate(schema, new SampleDataOptions { Seed = 4, Rows = 1 });

        XNamespace ns = schema.Report.Namespace;
        var value = dataset.Xml.Root!.Element(ns + "Header")!.Element(ns + "UnitPrice_Line")!.Value;

        Assert.True(
            decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out _),
            $"'{value}' should stay a decimal - 'Unit' alone must not trigger the UOM rule");
    }

    [Fact]
    public void Caption_indicator_values_stay_constant_across_rows()
    {
        var schema = MinimalSchema();
        var header = new DataItem { Name = "Header", Path = "/Header" };
        var line = new DataItem { Name = "Line", Path = "/Header/Line" };
        line.Columns.Add(new DatasetColumn { Name = "UnitPriceCaption", Path = "/Header/Line/UnitPriceCaption" });
        header.Children.Add(line);
        schema.Root.Children.Add(header);

        var dataset = SampleDataGenerator.Generate(schema, new SampleDataOptions { Seed = 21, Rows = 3 });

        XNamespace ns = schema.Report.Namespace;
        var values = dataset.Xml.Root!.Element(ns + "Header")!.Elements(ns + "Line")
            .Select(l => l.Element(ns + "UnitPriceCaption")!.Value)
            .Distinct()
            .ToList();

        // A real caption never varies per row; the sample must not either.
        Assert.Equal(["Unit Price"], values);
    }

    [Fact]
    public void Rows_produce_varied_values_across_instances()
    {
        // A repeating business item: 3 Line rows under 1 Header, each with a Quantity column that must
        // look distinct (driven by the seeded Random + instance index).
        var schema = MinimalSchema();
        var header = new DataItem { Name = "Header", Path = "/Header" };
        var line = new DataItem { Name = "Line", Path = "/Header/Line" };
        line.Columns.Add(new DatasetColumn { Name = "Quantity_Line", Path = "/Header/Line/Quantity_Line" });
        header.Children.Add(line);
        schema.Root.Children.Add(header);

        var dataset = SampleDataGenerator.Generate(schema, new SampleDataOptions { Seed = 99, Rows = 3 });

        XNamespace ns = schema.Report.Namespace;
        var values = dataset.Xml.Root!.Element(ns + "Header")!.Elements(ns + "Line")
            .Select(l => l.Element(ns + "Quantity_Line")!.Value)
            .ToList();

        Assert.Equal(3, values.Count);
        Assert.True(values.Distinct().Count() > 1, "expected varied Quantity values across the 3 Line rows");
    }

    // ---- Follow-up review (post-Phase-4.3): MaxRowsPerItem bounds GENERATION itself, not just merge-time
    // cloning - waste elimination for a large Rows value on a deeply-nested schema. ----

    [Fact]
    public void MaxRowsPerItem_caps_generated_instances_at_every_level_even_when_Rows_is_larger()
    {
        var schema = MinimalSchema();
        var header = new DataItem { Name = "Header", Path = "/Header" };
        var line = new DataItem { Name = "Line", Path = "/Header/Line" };
        line.Columns.Add(new DatasetColumn { Name = "Quantity_Line", Path = "/Header/Line/Quantity_Line" });
        header.Children.Add(line);
        schema.Root.Children.Add(header);

        var dataset = SampleDataGenerator.Generate(
            schema, new SampleDataOptions { Seed = 1, Rows = 10, MaxRowsPerItem = 4 });

        XNamespace ns = schema.Report.Namespace;

        // Header itself is a non-system child of the root: also capped to MaxRowsPerItem (4), not Rows (10).
        var headers = dataset.Xml.Root!.Elements(ns + "Header").ToList();
        Assert.Equal(4, headers.Count);

        // Each Header's own Line children are independently capped too - proves the cap applies at every
        // level of the recursive walk, not just the top.
        Assert.All(headers, h => Assert.Equal(4, h.Elements(ns + "Line").Count()));
    }

    /// <summary>
    /// A chain of <paramref name="depth"/> nested business data items (L1 ▸ L2 ▸ … each with one leaf column),
    /// under the root — the shape that makes generation multiply <c>count^depth</c> across nesting.
    /// </summary>
    private static DatasetTree DeepNestedSchema(int depth, string leafColumn)
    {
        var schema = MinimalSchema();
        var parent = schema.Root;
        for (var i = 1; i <= depth; i++)
        {
            var path = parent.Path.TrimEnd('/') + $"/L{i}";
            var item = new DataItem { Name = $"L{i}", Path = path };
            item.Columns.Add(new DatasetColumn { Name = leafColumn, Path = $"{path}/{leafColumn}" });
            parent.Children.Add(item);
            parent = item;
        }

        return schema;
    }

    private static int BusinessInstanceCount(SampleDataset dataset, int depth) =>
        dataset.Xml.Descendants().Count(e =>
            Enumerable.Range(1, depth).Any(i => e.Name.LocalName == $"L{i}"));

    [Fact]
    public void MaxTotalInstances_bounds_a_deeply_nested_schema_and_flags_truncation()
    {
        // Without a global budget, depth=3 at Rows=50 would generate 50 + 50^2 + 50^3 = 127,550 business
        // instances (the rows^depth blow-up that hangs a real preview). The global cap must stop it well short.
        const int depth = 3;
        var schema = DeepNestedSchema(depth, "Quantity_Line");

        var dataset = SampleDataGenerator.Generate(
            schema,
            new SampleDataOptions { Seed = 5, Rows = 50, MaxRowsPerItem = 100, MaxTotalInstances = 1000 });

        Assert.True(dataset.Truncated, "generation should report it was truncated by the global budget");
        var count = BusinessInstanceCount(dataset, depth);
        Assert.InRange(count, 1, 1000); // never exceeds the budget, regardless of Rows^depth
    }

    [Fact]
    public void MaxTotalInstances_is_a_no_op_when_generation_stays_within_budget()
    {
        // depth=3 at Rows=2 => 2 + 2^2 + 2^3 = 14 business instances, far under any real budget: output must be
        // byte-identical with the default budget vs. an explicitly huge one, and Truncated must stay false.
        const int depth = 3;
        var schema = DeepNestedSchema(depth, "Quantity_Line");

        var withDefault = SampleDataGenerator.Generate(schema, new SampleDataOptions { Seed = 9, Rows = 2 });
        var withHugeBudget = SampleDataGenerator.Generate(
            schema, new SampleDataOptions { Seed = 9, Rows = 2, MaxTotalInstances = 1_000_000 });

        Assert.False(withDefault.Truncated);
        Assert.False(withHugeBudget.Truncated);
        Assert.Equal(14, BusinessInstanceCount(withDefault, depth));
        Assert.Equal(withHugeBudget.Xml.ToString(), withDefault.Xml.ToString());
    }

    // ---- RepeaterConsumedPaths - only multiply rows for data items the DOCUMENT actually repeats ----

    [Fact]
    public void RepeaterConsumedPaths_multiplies_only_the_consumed_sibling_item()
    {
        // Two sibling repeating items under Header: only LineA's path is in RepeaterConsumedPaths (as if a
        // real w15:repeatingSection in the document were bound to it). LineB - a sibling nothing repeats -
        // must get exactly ONE instance instead of Rows, even though it is shaped identically to LineA.
        var schema = MinimalSchema();
        var header = new DataItem { Name = "Header", Path = "/Header" };
        var lineA = new DataItem { Name = "LineA", Path = "/Header/LineA" };
        lineA.Columns.Add(new DatasetColumn { Name = "Value_LineA", Path = "/Header/LineA/Value_LineA" });
        var lineB = new DataItem { Name = "LineB", Path = "/Header/LineB" };
        lineB.Columns.Add(new DatasetColumn { Name = "Value_LineB", Path = "/Header/LineB/Value_LineB" });
        header.Children.Add(lineA);
        header.Children.Add(lineB);
        schema.Root.Children.Add(header);

        var dataset = SampleDataGenerator.Generate(
            schema,
            new SampleDataOptions
            {
                Seed = 1,
                Rows = 10,
                RepeaterConsumedPaths = new HashSet<string> { "/Header/LineA" },
            });

        XNamespace ns = schema.Report.Namespace;
        var headerElement = dataset.Xml.Root!.Element(ns + "Header")!;

        Assert.Equal(10, headerElement.Elements(ns + "LineA").Count());
        Assert.Single(headerElement.Elements(ns + "LineB"));
    }

    [Fact]
    public void RepeaterConsumedPaths_null_treats_every_business_item_as_consumed()
    {
        // The default (null - no live document was scanned) must reproduce the pre-B23 behavior exactly:
        // every business item multiplies to Rows, regardless of nesting.
        var schema = MinimalSchema();
        var header = new DataItem { Name = "Header", Path = "/Header" };
        var line = new DataItem { Name = "Line", Path = "/Header/Line" };
        line.Columns.Add(new DatasetColumn { Name = "Value_Line", Path = "/Header/Line/Value_Line" });
        header.Children.Add(line);
        schema.Root.Children.Add(header);

        var dataset = SampleDataGenerator.Generate(
            schema, new SampleDataOptions { Seed = 1, Rows = 4, RepeaterConsumedPaths = null });

        XNamespace ns = schema.Report.Namespace;
        var headers = dataset.Xml.Root!.Elements(ns + "Header").ToList();
        Assert.Equal(4, headers.Count);
        Assert.All(headers, h => Assert.Equal(4, h.Elements(ns + "Line").Count()));
    }

    [Fact]
    public void RepeaterConsumedPaths_prevents_an_unconsumed_deep_sibling_from_starving_the_repeater_of_budget()
    {
        // Header has an unconsumed deep chain (J1 -> J2 -> J3, nothing in any document ever binds to it)
        // alongside the item an ACTUAL repeater targets (Line). Without RepeaterConsumedPaths (legacy
        // behavior - every business item multiplies), the J-chain's rows^depth blow-up exhausts
        // MaxTotalInstances before Line - the item the document actually repeats, and which comes AFTER the
        // J-chain in schema/document order - gets any rows at all: exactly the residual bug this rule
        // describes ("unused items consumed the budget while items actually repeated got truncated"). With
        // RepeaterConsumedPaths naming only "/Header/Line", the unconsumed chain costs almost nothing and
        // Line gets its full Rows.
        var schema = MinimalSchema();
        var header = new DataItem { Name = "Header", Path = "/Header" };
        var j1 = new DataItem { Name = "J1", Path = "/Header/J1" };
        var j2 = new DataItem { Name = "J2", Path = "/Header/J1/J2" };
        var j3 = new DataItem { Name = "J3", Path = "/Header/J1/J2/J3" };
        j3.Columns.Add(new DatasetColumn { Name = "Leaf", Path = "/Header/J1/J2/J3/Leaf" });
        j2.Children.Add(j3);
        j1.Children.Add(j2);
        var line = new DataItem { Name = "Line", Path = "/Header/Line" };
        line.Columns.Add(new DatasetColumn { Name = "ItemNo_Line", Path = "/Header/Line/ItemNo_Line" });
        header.Children.Add(j1);
        header.Children.Add(line);
        schema.Root.Children.Add(header);

        const int rows = 50;
        const int budget = 100;
        XNamespace ns = schema.Report.Namespace;

        var legacy = SampleDataGenerator.Generate(
            schema, new SampleDataOptions { Seed = 1, Rows = rows, MaxTotalInstances = budget });

        Assert.True(legacy.Truncated, "legacy (all-consumed) generation should exhaust the budget on the unconsumed J-chain");
        var legacyLineCount = legacy.Xml.Root!.Element(ns + "Header")!.Elements(ns + "Line").Count();
        Assert.Equal(0, legacyLineCount); // starved: the repeater's own target got nothing.

        var fixedDataset = SampleDataGenerator.Generate(
            schema,
            new SampleDataOptions
            {
                Seed = 1,
                Rows = rows,
                MaxTotalInstances = budget,
                RepeaterConsumedPaths = new HashSet<string> { "/Header/Line" },
            });

        Assert.False(fixedDataset.Truncated, "the unconsumed J-chain should no longer eat the budget");
        var fixedLineCount = fixedDataset.Xml.Root!.Element(ns + "Header")!.Elements(ns + "Line").Count();
        Assert.Equal(rows, fixedLineCount);
    }

    [Fact]
    public void MaxRowsPerItem_does_not_affect_a_Rows_value_already_under_the_cap()
    {
        // The default (100) - and any cap comfortably above Rows - must not change output at all: proves
        // the waste-elimination safeguard is a pure no-op for ordinary Rows values (matches every
        // corpus/typical usage, where Rows is always far below the default cap).
        var schema = SchemaWithHeaderColumns("Quantity_Line");

        var withDefaultCap = SampleDataGenerator.Generate(schema, new SampleDataOptions { Seed = 7, Rows = 3 });
        var withExplicitSameCap = SampleDataGenerator.Generate(
            schema, new SampleDataOptions { Seed = 7, Rows = 3, MaxRowsPerItem = 100 });

        Assert.Equal(withDefaultCap.Xml.ToString(), withExplicitSameCap.Xml.ToString());
    }

    [Fact]
    public void DataOverrides_valid_file_is_returned_as_is_and_ignores_seed_and_rows()
    {
        var xml =
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
            + "<NavWordReportXmlPart xmlns=\"urn:microsoft-dynamics-nav/reports/TestReport/50000/\">"
            + "<Header><CompanyName>Contoso</CompanyName></Header>"
            + "</NavWordReportXmlPart>";

        var path = Path.Combine(Path.GetTempPath(), $"bcwl-overrides-{Guid.NewGuid():N}.xml");
        File.WriteAllText(path, xml, new UTF8Encoding(false));

        try
        {
            var schema = MinimalSchema();
            var dataset = SampleDataGenerator.Generate(
                schema,
                new SampleDataOptions { Seed = 123, Rows = 5, DataOverridesPath = path });

            Assert.Equal("urn:microsoft-dynamics-nav/reports/TestReport/50000/", dataset.Namespace);

            XNamespace ns = dataset.Namespace;
            Assert.Equal("Contoso", dataset.Xml.Root!.Element(ns + "Header")!.Element(ns + "CompanyName")!.Value);

            // No Rows-driven expansion happened: overrides are returned verbatim, exactly 1 Header.
            Assert.Single(dataset.Xml.Root!.Elements(ns + "Header"));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void DataOverrides_handles_utf16_le_bom_like_real_bc_exports()
    {
        var xml =
            "<?xml version=\"1.0\" encoding=\"utf-16\" standalone=\"yes\"?>"
            + "<NavWordReportXmlPart xmlns=\"urn:microsoft-dynamics-nav/reports/TestReport/50000/\">"
            + "<Header><CompanyName>Contoso</CompanyName></Header>"
            + "</NavWordReportXmlPart>";

        var path = Path.Combine(Path.GetTempPath(), $"bcwl-overrides-utf16-{Guid.NewGuid():N}.xml");
        File.WriteAllText(path, xml, new UnicodeEncoding(bigEndian: false, byteOrderMark: true));

        try
        {
            var schema = MinimalSchema();
            var dataset = SampleDataGenerator.Generate(schema, new SampleDataOptions { DataOverridesPath = path });

            Assert.Equal("urn:microsoft-dynamics-nav/reports/TestReport/50000/", dataset.Namespace);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void DataOverrides_malformed_root_error_names_both_accepted_shapes()
    {
        var xml = "<?xml version=\"1.0\" encoding=\"utf-8\"?><SomeOtherRoot></SomeOtherRoot>";
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-overrides-bad-root-{Guid.NewGuid():N}.xml");
        File.WriteAllText(path, xml, new UTF8Encoding(false));

        try
        {
            var schema = MinimalSchema();
            var ex = Assert.Throws<InvalidDataException>(() =>
                SampleDataGenerator.Generate(schema, new SampleDataOptions { DataOverridesPath = path }));

            // The caller holding the wrong file must learn BOTH encodings that would have worked.
            Assert.Contains("NavWordReportXmlPart", ex.Message);
            Assert.Contains("ReportDataSet", ex.Message);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void DataOverrides_wrong_namespace_throws_InvalidDataException()
    {
        var xml =
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
            + "<NavWordReportXmlPart xmlns=\"urn:some-other-namespace/\"></NavWordReportXmlPart>";
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-overrides-bad-ns-{Guid.NewGuid():N}.xml");
        File.WriteAllText(path, xml, new UTF8Encoding(false));

        try
        {
            var schema = MinimalSchema();
            Assert.Throws<InvalidDataException>(() =>
                SampleDataGenerator.Generate(schema, new SampleDataOptions { DataOverridesPath = path }));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void DataOverrides_missing_file_throws_FileNotFoundException()
    {
        var schema = MinimalSchema();
        var missingPath = Path.Combine(Path.GetTempPath(), $"bcwl-missing-{Guid.NewGuid():N}.xml");

        Assert.Throws<FileNotFoundException>(() =>
            SampleDataGenerator.Generate(schema, new SampleDataOptions { DataOverridesPath = missingPath }));
    }

    // ---- data overrides: the report UI's Send to → XML export shape (GitHub issue #4) ----

    /// <summary>
    /// Writes an export-shape overrides file and loads it through <see cref="SampleDataGenerator.Generate"/>
    /// against <see cref="MinimalSchema"/> (report id "1", namespace <c>…/reports/Test/1/</c>), deleting the
    /// file afterwards.
    /// </summary>
    private static SampleDataset LoadExportOverrides(string exportXml, Encoding? encoding = null)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-overrides-export-{Guid.NewGuid():N}.xml");
        File.WriteAllText(path, exportXml, encoding ?? new UTF8Encoding(false));
        try
        {
            return SampleDataGenerator.Generate(
                MinimalSchema(), new SampleDataOptions { DataOverridesPath = path });
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void DataOverrides_export_shape_is_converted_to_the_dataset_part_shape()
    {
        // The shape BC's report UI actually produces (Send to → XML): namespace-less ReportDataSet root,
        // Labels/Label[@name], DataItems/DataItem[@name]/Columns/Column[@name], one DataItem SIBLING per
        // row — nothing here is in the layout's namespace or element-per-column encoding.
        var dataset = LoadExportOverrides(
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
            + "<ReportDataSet name=\"Test\" id=\"1\" language=\"en-US\" formatRegion=\"en-US\" wordMergeDataItem=\"Header\">"
            + "<BCReportInformation><CompanyDisplayName>Contoso Ltd.</CompanyDisplayName></BCReportInformation>"
            + "<Labels><Label name=\"TotalLbl\">Total</Label></Labels>"
            + "<DataItems><DataItem name=\"Header\">"
            + "<Columns><Column name=\"CompanyName\">Contoso</Column></Columns>"
            + "<DataItems>"
            + "<DataItem name=\"Line\"><Columns><Column name=\"ItemNo\">A</Column></Columns></DataItem>"
            + "<DataItem name=\"Line\"><Columns><Column name=\"ItemNo\">B</Column></Columns></DataItem>"
            + "</DataItems>"
            + "</DataItem></DataItems>"
            + "</ReportDataSet>");

        // Converted into the layout's own data-store part shape, in the SCHEMA's namespace.
        Assert.Equal("urn:microsoft-dynamics-nav/reports/Test/1/", dataset.Namespace);
        XNamespace ns = dataset.Namespace;
        var root = dataset.Xml.Root!;
        Assert.Equal(ns + "NavWordReportXmlPart", root.Name);

        Assert.Equal("Contoso Ltd.", root.Element(ns + "BCReportInformation")!.Element(ns + "CompanyDisplayName")!.Value);
        Assert.Equal("Total", root.Element(ns + "Labels")!.Element(ns + "TotalLbl")!.Value);

        var header = Assert.Single(root.Elements(ns + "Header"));
        Assert.Equal("Contoso", header.Element(ns + "CompanyName")!.Value);

        // One element per DataItem occurrence, document order preserved — the repeater expansion shape.
        var lines = header.Elements(ns + "Line").ToList();
        Assert.Equal(2, lines.Count);
        Assert.Equal(new[] { "A", "B" }, lines.Select(l => l.Element(ns + "ItemNo")!.Value));
    }

    [Fact]
    public void DataOverrides_export_decimalformatter_is_applied_per_column_with_the_formatRegion_culture()
    {
        // All four decimalformatter cases observed in real sandbox exports, under a culture (de-DE) whose
        // separators DIFFER from both the raw encoding and en-US — so applying the wrong culture, or none,
        // fails visibly. The un-attributed column proves the rule is per-column, never global.
        var dataset = LoadExportOverrides(
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
            + "<ReportDataSet name=\"Test\" id=\"1\" language=\"en-US\" formatRegion=\"de-DE\">"
            + "<DataItems><DataItem name=\"Header\"><Columns>"
            + "<Column name=\"Amount\" decimalformatter=\"#,##0.00\">1002060</Column>"
            + "<Column name=\"Total\" decimalformatter=\"$#,##0.00;$-#,##0.00\">-300</Column>"
            + "<Column name=\"Quantity\" decimalformatter=\"#,##0.#####\">3</Column>"
            + "<Column name=\"PreFormatted\">1,800.00</Column>"
            + "<Column name=\"NotANumber\" decimalformatter=\"#,##0.00\">n/a</Column>"
            + "</Columns></DataItem></DataItems>"
            + "</ReportDataSet>");

        XNamespace ns = dataset.Namespace;
        var header = dataset.Xml.Root!.Element(ns + "Header")!;

        Assert.Equal("1.002.060,00", header.Element(ns + "Amount")!.Value);
        // Sectioned pattern: the NEGATIVE section carries its own literal currency symbol and sign.
        Assert.Equal("$-300,00", header.Element(ns + "Total")!.Value);
        // Precision-variable pattern: no forced decimals on a whole number.
        Assert.Equal("3", header.Element(ns + "Quantity")!.Value);
        // No attribute → verbatim, even though it LOOKS numeric (it arrived pre-formatted).
        Assert.Equal("1,800.00", header.Element(ns + "PreFormatted")!.Value);
        // Attribute but unparseable raw text → verbatim, never mangled.
        Assert.Equal("n/a", header.Element(ns + "NotANumber")!.Value);
    }

    [Fact]
    public void DataOverrides_export_formatRegion_falls_back_to_language_then_invariant()
    {
        const string Columns =
            "<DataItems><DataItem name=\"Header\"><Columns>"
            + "<Column name=\"Amount\" decimalformatter=\"#,##0.00\">1002060</Column>"
            + "</Columns></DataItem></DataItems>";

        var languageOnly = LoadExportOverrides(
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
            + $"<ReportDataSet name=\"Test\" id=\"1\" language=\"de-DE\">{Columns}</ReportDataSet>");
        var neither = LoadExportOverrides(
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
            + $"<ReportDataSet name=\"Test\" id=\"1\">{Columns}</ReportDataSet>");

        XNamespace ns = languageOnly.Namespace;
        Assert.Equal("1.002.060,00", languageOnly.Xml.Root!.Element(ns + "Header")!.Element(ns + "Amount")!.Value);
        Assert.Equal("1,002,060.00", neither.Xml.Root!.Element(ns + "Header")!.Element(ns + "Amount")!.Value);
    }

    [Fact]
    public void DataOverrides_export_from_a_different_report_throws_with_both_report_ids_named()
    {
        // Feeding report 1304's export to a report-1 layout would otherwise produce a preview where every
        // binding is silently unresolved — the id cross-check turns that into an actionable error.
        var ex = Assert.Throws<InvalidDataException>(() => LoadExportOverrides(
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
            + "<ReportDataSet name=\"Standard Sales - Quote\" id=\"1304\" language=\"en-US\" formatRegion=\"en-US\">"
            + "<DataItems><DataItem name=\"Header\"><Columns><Column name=\"No\">X</Column></Columns></DataItem></DataItems>"
            + "</ReportDataSet>"));

        Assert.Contains("1304", ex.Message);
        Assert.Contains("'Test'", ex.Message);
    }

    [Fact]
    public void DataOverrides_export_handles_utf16_le_bom_like_real_bc_exports()
    {
        // Real Send to → XML exports are UTF-16 LE with a BOM; the encoding declaration must be honored.
        var dataset = LoadExportOverrides(
            "<?xml version=\"1.0\" encoding=\"utf-16\" standalone=\"yes\"?>"
            + "<ReportDataSet name=\"Test\" id=\"1\" language=\"en-US\" formatRegion=\"en-US\">"
            + "<DataItems><DataItem name=\"Header\"><Columns><Column name=\"CompanyName\">Contoso</Column></Columns></DataItem></DataItems>"
            + "</ReportDataSet>",
            new UnicodeEncoding(bigEndian: false, byteOrderMark: true));

        XNamespace ns = dataset.Namespace;
        Assert.Equal("Contoso", dataset.Xml.Root!.Element(ns + "Header")!.Element(ns + "CompanyName")!.Value);
    }

    [Fact]
    public void DataOverrides_export_invalid_element_name_throws_naming_the_offender()
    {
        // Export attribute values become ELEMENT names in the converted shape, so a name no XML element
        // can carry must fail with the offender named — not a bare XmlException from deep inside XLinq.
        var ex = Assert.Throws<InvalidDataException>(() => LoadExportOverrides(
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
            + "<ReportDataSet name=\"Test\" id=\"1\">"
            + "<DataItems><DataItem name=\"Header\"><Columns><Column name=\"Bad Name\">X</Column></Columns></DataItem></DataItems>"
            + "</ReportDataSet>"));

        Assert.Contains("Bad Name", ex.Message);
    }

    [Fact]
    public void Whole_word_matching_avoids_substring_false_positives()
    {
        // "LastUpdatedBy" contains the raw substring "date" (Up-date-d); "PrivateNote" and "Innovation"
        // both contain the raw substring "vat" (pri-vat-e / Inno-vat-ion). Whole-word matching must NOT
        // classify any of these as Date / VAT-amount columns.
        var schema = SchemaWithHeaderColumns("LastUpdatedBy", "PrivateNote", "Innovation");
        var dataset = SampleDataGenerator.Generate(schema, new SampleDataOptions { Seed = 21, Rows = 1 });

        XNamespace ns = schema.Report.Namespace;
        var header = dataset.Xml.Root!.Element(ns + "Header")!;

        var lastUpdatedBy = header.Element(ns + "LastUpdatedBy")!.Value;
        Assert.False(
            DateTime.TryParseExact(lastUpdatedBy, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out _),
            $"'{lastUpdatedBy}' should NOT be treated as a date");

        var privateNote = header.Element(ns + "PrivateNote")!.Value;
        Assert.False(
            decimal.TryParse(privateNote, NumberStyles.Number, CultureInfo.InvariantCulture, out _),
            $"'{privateNote}' should NOT be treated as a VAT/amount decimal");

        var innovation = header.Element(ns + "Innovation")!.Value;
        Assert.False(
            decimal.TryParse(innovation, NumberStyles.Number, CultureInfo.InvariantCulture, out _),
            $"'{innovation}' should NOT be treated as a VAT/amount decimal");
    }

    // ---- Image-ish column names generate a marker, not fake text ----

    [Theory]
    [InlineData("CompanyPicture")]
    [InlineData("ProductImage")]
    [InlineData("PaymentServiceLogo")]
    [InlineData("CustomerPhoto")]
    [InlineData("ScannedBitmap")]
    public void Image_named_column_generates_the_image_marker_not_fake_text(string columnName)
    {
        var schema = SchemaWithHeaderColumns(columnName);
        var dataset = SampleDataGenerator.Generate(schema, new SampleDataOptions { Seed = 1, Rows = 1 });

        XNamespace ns = schema.Report.Namespace;
        var value = dataset.Xml.Root!.Element(ns + "Header")!.Element(ns + columnName)!.Value;

        Assert.Equal("[image]", value);
    }

    [Fact]
    public void Label_suffixed_column_that_also_looks_image_named_still_renders_as_a_humanized_caption()
    {
        // Label precedence: column.IsLabel is checked BEFORE the image-token check (same as every other
        // type-inference branch), so a hypothetical "<image-word>Lbl" column still gets its humanized
        // caption text, never the "[image]" marker - a *Lbl column is always a caption, never typed data.
        var schema = SchemaWithHeaderColumns("CompanyPictureLbl");
        var dataset = SampleDataGenerator.Generate(schema, new SampleDataOptions { Seed = 1, Rows = 1 });

        XNamespace ns = schema.Report.Namespace;
        var value = dataset.Xml.Root!.Element(ns + "Header")!.Element(ns + "CompanyPictureLbl")!.Value;

        Assert.Equal("Company Picture", value);
        Assert.NotEqual("[image]", value);
    }

    [Fact]
    public void Non_image_columns_are_unaffected_by_the_image_marker_check()
    {
        // "LastUpdatedBy"/"PrivateNote"-style guard, but for the NEW image tokens: whole-word matching must
        // not fire on a column that merely shares letters with "Picture"/"Image"/"Logo"/"Photo"/"Bitmap"
        // without containing any of them as a whole camelCase/underscore word.
        var schema = SchemaWithHeaderColumns("CustomerName", "TotalAmount");
        var dataset = SampleDataGenerator.Generate(schema, new SampleDataOptions { Seed = 1, Rows = 1 });

        XNamespace ns = schema.Report.Namespace;
        var header = dataset.Xml.Root!.Element(ns + "Header")!;

        Assert.NotEqual("[image]", header.Element(ns + "CustomerName")!.Value);
        Assert.NotEqual("[image]", header.Element(ns + "TotalAmount")!.Value);
    }

    [Fact]
    public void Real_corpus_unbound_image_columns_generate_the_marker()
    {
        // Grounded against the real corpus: StandardStatement.docx's schema (/Customer/Integer/...) has
        // several leaf columns that look like image data by name - CompanyInfo1Picture, CompanyInfo2Picture,
        // CompanyInfo3Picture - which, unlike CompanyPicture itself, are never bound by any control anywhere
        // in that layout's document/header/footer parts (verified directly against the real corpus file).
        // SampleDataGenerator has no notion of "bound" at all - it walks the whole schema tree regardless of
        // what any control references - so this is exactly the scenario the marker exists to harden: these
        // columns would otherwise get lorem-ish fallback text despite genuinely being image-shaped data in BC.
        var schema = SchemaProvider.FromLayout(Corpus.Path(Corpus.StandardStatement));
        var dataset = SampleDataGenerator.Generate(schema, new SampleDataOptions { Seed = 1, Rows = 1 });

        XNamespace ns = schema.Report.Namespace;
        var integer = dataset.Xml.Root!.Element(ns + "Customer")!.Element(ns + "Integer")!;

        foreach (var columnName in new[]
                 {
                     "CompanyInfo1Picture", "CompanyInfo2Picture", "CompanyInfo3Picture", "CompanyPicture",
                 })
        {
            var element = integer.Element(ns + columnName);
            Assert.NotNull(element);
            Assert.Equal("[image]", element!.Value);
        }
    }

    [Fact]
    public void Email_column_produces_an_at_sign_value()
    {
        var schema = SchemaWithHeaderColumns("ContactEmail");
        var dataset = SampleDataGenerator.Generate(schema, new SampleDataOptions { Seed = 11, Rows = 1 });

        XNamespace ns = schema.Report.Namespace;
        var value = dataset.Xml.Root!.Element(ns + "Header")!.Element(ns + "ContactEmail")!.Value;

        Assert.Contains('@', value);
        Assert.EndsWith("@example.com", value, StringComparison.Ordinal);
    }

    [Fact]
    public void Phone_column_produces_a_digit_bearing_value()
    {
        var schema = SchemaWithHeaderColumns("ContactPhoneNo");
        var dataset = SampleDataGenerator.Generate(schema, new SampleDataOptions { Seed = 12, Rows = 1 });

        XNamespace ns = schema.Report.Namespace;
        var value = dataset.Xml.Root!.Element(ns + "Header")!.Element(ns + "ContactPhoneNo")!.Value;

        Assert.True(value.Any(char.IsDigit), $"'{value}' should contain at least one digit");
    }

    [Fact]
    public void Code_column_produces_an_uppercase_code()
    {
        var schema = SchemaWithHeaderColumns("DocumentNo");
        var dataset = SampleDataGenerator.Generate(schema, new SampleDataOptions { Seed = 13, Rows = 1 });

        XNamespace ns = schema.Report.Namespace;
        var value = dataset.Xml.Root!.Element(ns + "Header")!.Element(ns + "DocumentNo")!.Value;

        Assert.Equal(value.ToUpperInvariant(), value);
        Assert.Matches(@"^[A-Z0-9]+-\d{4,}$", value);
    }

    [Fact]
    public void Empty_schema_with_only_root_yields_a_bare_root_element_without_throwing()
    {
        var schema = MinimalSchema();

        var dataset = SampleDataGenerator.Generate(schema, new SampleDataOptions { Seed = 1, Rows = 3 });

        var root = dataset.Xml.Root!;
        Assert.Equal("NavWordReportXmlPart", root.Name.LocalName);
        Assert.False(root.HasElements);
        Assert.Equal(string.Empty, root.Value);
    }

    [Fact]
    public void Column_directly_under_root_is_emitted()
    {
        var schema = MinimalSchema();
        schema.Root.Columns.Add(new DatasetColumn { Name = "RootLevelField", Path = "/RootLevelField" });

        var dataset = SampleDataGenerator.Generate(schema, new SampleDataOptions { Seed = 2, Rows = 1 });

        XNamespace ns = schema.Report.Namespace;
        var element = dataset.Xml.Root!.Element(ns + "RootLevelField");

        Assert.NotNull(element);
        Assert.False(string.IsNullOrEmpty(element!.Value));
    }
}
