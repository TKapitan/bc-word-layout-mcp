using System.Xml.Linq;
using BcWordLayout.Domain;
using BcWordLayout.Merge;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// A real dataset printed for several documents holds several top-level records, and BC renders the report
/// body once per record. The mock renders record [1] only, and now says so with a
/// <c>multiple-top-level-records</c> warning (GitHub issue #105). It used to stay silent. The body is not
/// repeated per record: there is no observed OOXML for how BC joins the documents (ADR-0005).
/// </summary>
public class TopLevelRecordsWarningTests : IDisposable
{
    private const string Ns = "urn:microsoft-dynamics-nav/reports/TestReport/50000/";
    private const string Kind = "multiple-top-level-records";

    private readonly string _dir = Directory.CreateTempSubdirectory("bcwl-top-level-records-").FullName;

    public void Dispose() => Directory.Delete(_dir, recursive: true);

    private string CopyOfCorpus(string corpusFile)
    {
        var path = Path.Combine(_dir, corpusFile);
        File.Copy(Corpus.Path(corpusFile), path);
        return path;
    }

    /// <summary>The layout's own BC data-store part: its dataset skeleton, one instance of everything.</summary>
    private static XDocument DatasetPartOf(string layoutPath)
    {
        using var doc = WordprocessingDocument.Open(layoutPath, false);
        foreach (var part in doc.MainDocumentPart!.CustomXmlParts)
        {
            using var stream = part.GetStream();
            var xml = XDocument.Load(stream);
            if (xml.Root?.Name.LocalName == "NavWordReportXmlPart")
            {
                return xml;
            }
        }

        throw new InvalidOperationException("No BC dataset part.");
    }

    private static string BodyText(string docxPath)
    {
        using var doc = WordprocessingDocument.Open(docxPath, false);
        return string.Concat(doc.MainDocumentPart!.Document!.Body!.Descendants<Text>().Select(t => t.Text));
    }

    [Fact]
    public void A_purchase_order_dataset_with_two_Purchase_Header_records_warns_that_only_the_first_is_rendered()
    {
        // The issue's reproduction: StandardPurchaseOrder with an override holding two orders. The merged copy
        // shows FIRST-DOC only, and the only warning used to be external-relationship-stripped.
        var layout = CopyOfCorpus(Corpus.StandardPurchaseOrder);
        var dataset = DatasetPartOf(layout);
        var first = dataset.Root!.Elements().Single(e => e.Name.LocalName == "Purchase_Header");
        first.Elements().Single(e => e.Name.LocalName == "BuyFromAddr1").Value = "FIRST-DOC";
        var second = new XElement(first);
        second.Elements().Single(e => e.Name.LocalName == "BuyFromAddr1").Value = "SECOND-DOC";
        first.AddAfterSelf(second);
        var overridesPath = Path.Combine(_dir, "two-orders.xml");
        dataset.Save(overridesPath);
        var merged = Path.Combine(_dir, "merged.docx");

        var result = MergeEngine.Merge(layout, merged, new MergeOptions { DataOverridesPath = overridesPath });

        var warning = Assert.Single(result.Warnings, w => w.Kind == Kind);
        Assert.Equal("/Purchase_Header", warning.Location);
        Assert.StartsWith(
            "The dataset has 2 top-level 'Purchase_Header' records, but the preview renders the first only",
            warning.Message,
            StringComparison.Ordinal);

        // What the warning describes: the first order is in the copy, the second nowhere.
        var text = BodyText(merged);
        Assert.Contains("FIRST-DOC", text, StringComparison.Ordinal);
        Assert.DoesNotContain("SECOND-DOC", text, StringComparison.Ordinal);
    }

    [Fact]
    public void The_layouts_own_single_record_dataset_and_generated_sample_data_raise_no_warning()
    {
        // Generation never multiplies a top-level item no repeater repeats, so a generated preview has
        // nothing to warn about; neither does a real export of one document.
        var layout = CopyOfCorpus(Corpus.StandardPurchaseOrder);
        var overridesPath = Path.Combine(_dir, "one-order.xml");
        DatasetPartOf(layout).Save(overridesPath);

        var generated = MergeEngine.Merge(layout, Path.Combine(_dir, "generated.docx"), new MergeOptions { Rows = 5 });
        var overridden = MergeEngine.Merge(
            layout, Path.Combine(_dir, "overridden.docx"), new MergeOptions { DataOverridesPath = overridesPath });

        Assert.DoesNotContain(generated.Warnings, w => w.Kind == Kind);
        Assert.DoesNotContain(overridden.Warnings, w => w.Kind == Kind);
    }

    [Fact]
    public void A_top_level_item_a_repeater_repeats_renders_every_record_and_raises_no_warning()
    {
        // SalespersonCommission repeats its top-level Salesperson_Purchaser item in a row repeater, so every
        // salesperson in the dataset gets its rows. Its single Labels/Totals items are not repeated at all.
        var layout = CopyOfCorpus(Corpus.SalespersonCommission);
        var dataset = DatasetPartOf(layout);
        var salesperson = dataset.Root!.Elements().Single(e => e.Name.LocalName == "Salesperson_Purchaser");
        salesperson.AddAfterSelf(new XElement(salesperson), new XElement(salesperson));
        var overridesPath = Path.Combine(_dir, "three-salespeople.xml");
        dataset.Save(overridesPath);

        var merged = Path.Combine(_dir, "merged.docx");

        var result = MergeEngine.Merge(layout, merged, new MergeOptions { DataOverridesPath = overridesPath });

        Assert.DoesNotContain(result.Warnings, w => w.Kind == Kind);

        // One Salesperson row per record: each carries the salesperson's own name control.
        using var doc = WordprocessingDocument.Open(merged, false);
        var nameControls = doc.MainDocumentPart!.Document!.Body!.Descendants<SdtElement>().Count(s =>
            BindingXPath.Segments(SdtInspector.ReadXPath(s))
                .SequenceEqual(new[] { "NavWordReportXmlPart", "Salesperson_Purchaser", "Salesperson_Purchaser_Name" }));
        Assert.Equal(3, nameControls);
    }

    [Fact]
    public void Each_unrepeated_top_level_item_held_more_than_once_warns_once_with_its_count_in_dataset_order()
    {
        // Bound only at top level (the implicit top-level repeater): Header held 3 times, Totals twice,
        // Labels and the BCReportInformation system node once each.
        var body = SyntheticLayout.BoundField("/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:No[1]", SyntheticLayout.GoodItemId)
            + SyntheticLayout.BoundField("/ns0:NavWordReportXmlPart[1]/ns0:Totals[1]/ns0:Amount[1]", SyntheticLayout.GoodItemId);
        var layoutPath = SyntheticLayout.Create(body);

        XNamespace ns = Ns;
        XElement Info() => new(ns + "BCReportInformation", new XElement(ns + "CreationDateTime", "2026-01-01"));
        XElement Header(string no) => new(ns + "Header", new XElement(ns + "No", no));
        XElement Totals(string amount) => new(ns + "Totals", new XElement(ns + "Amount", amount));
        var xml = new XDocument(new XElement(
            ns + "NavWordReportXmlPart",
            Info(),
            new XElement(ns + "Labels", new XElement(ns + "NoCaption", "No.")),
            Header("H1"),
            Totals("1.00"),
            Header("H2"),
            Totals("2.00"),
            Header("H3")));

        try
        {
            using var doc = WordprocessingDocument.Open(layoutPath, true);
            var result = MergeEngine.Merge(doc, new SampleDataset { Xml = xml, Namespace = Ns });

            var warnings = result.Warnings.Where(w => w.Kind == Kind).ToList();
            Assert.Equal(new[] { "/Header", "/Totals" }, warnings.Select(w => w.Location));
            Assert.StartsWith("The dataset has 3 top-level 'Header' records", warnings[0].Message, StringComparison.Ordinal);
            Assert.StartsWith("The dataset has 2 top-level 'Totals' records", warnings[1].Message, StringComparison.Ordinal);
            Assert.Equal(2, result.Stats.FieldsFilled);
            Assert.Equal(0, result.Stats.Unresolved);
        }
        finally
        {
            File.Delete(layoutPath);
        }
    }
}
