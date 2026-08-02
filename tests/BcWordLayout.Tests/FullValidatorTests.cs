using BcWordLayout.Domain.Models;
using BcWordLayout.Merge;

namespace BcWordLayout.Tests;

/// <summary>
/// Covers <see cref="FullValidator"/> — the <c>validate_layout</c> "full" level, which layers a real
/// dry-run merge (<see cref="MergeEngine"/>) on top of <c>LayoutValidator.Quick</c>'s structural/binding
/// checks. <see cref="FullValidator.Full"/> manages its own temp merge output internally (create, use,
/// delete in a try/finally), so these tests only need to clean up the synthetic layout files they create.
/// </summary>
public class FullValidatorTests
{
    [Theory]
    [InlineData(Corpus.SalesInvoice)]
    [InlineData(Corpus.InventoryOrderDetails)]
    [InlineData(Corpus.StandardStatement)]
    public void All_corpus_layouts_full_validate_with_zero_errors(string fileName)
    {
        var result = FullValidator.Full(Corpus.Path(fileName));

        Assert.Equal("full", result.Level);
        Assert.Equal(0, result.ErrorCount);
        Assert.True(result.Passed,
            "expected pass; errors: " +
            string.Join(" | ", result.Findings.Where(f => f.Severity == FindingSeverity.Error).Select(f => f.Message)));
    }

    [Fact]
    public void Synthetic_layout_with_bogus_binding_xpath_yields_an_error_finding_at_full_level()
    {
        const string BogusXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:BogusField[1]";

        var body = SyntheticLayout.BoundField(BogusXPath, SyntheticLayout.GoodItemId);
        var layoutPath = SyntheticLayout.Create(body);

        try
        {
            var result = FullValidator.Full(layoutPath);

            Assert.Equal("full", result.Level);
            Assert.False(result.Passed);

            var errorFindings = result.Findings.Where(f => f.Severity == FindingSeverity.Error).ToList();
            Assert.Contains(errorFindings, f => f.Check is "unresolved-binding" or "xpath-error");
            Assert.Contains(errorFindings, f => f.Message.Contains("BogusField"));
        }
        finally
        {
            File.Delete(layoutPath);
        }
    }

    [Fact]
    public void Quick_findings_are_preserved_alongside_merge_findings()
    {
        // SalesInvoice's known quick-level finding (an external attachedTemplate relationship, a warning)
        // must still be present at "full" level — full is additive over quick, not a replacement.
        var result = FullValidator.Full(Corpus.Path(Corpus.SalesInvoice));

        Assert.Contains(result.Findings, f => f.Check == "attached-template");
    }

    [Fact]
    public void Structurally_divergent_in_repeater_binding_does_not_raise_a_false_unresolved_binding_error()
    {
        // Re-anchoring shape (b) - the DEEPER-divergent path: a field bound to Header/Totals/Amount —
        // deeper than the enclosing Header/Line repeater but branching off a DIFFERENT element — used to
        // have its remainder evaluated against the wrong context node (the Line row), typically resolving
        // nothing and reporting a FALSE "unresolved-binding" ERROR at validate_layout level=full for a
        // layout BC itself renders fine. After the fix, the value resolves correctly via a document-root
        // fallback: only the (Warning-severity) xpath-fallback finding should appear, never the Error.
        const string RepeaterXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Line";
        const string FieldXPath = "/ns0:NavWordReportXmlPart[1]/ns0:Header[1]/ns0:Totals[1]/ns0:Amount[1]";
        var schemaXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<NavWordReportXmlPart xmlns=\"urn:microsoft-dynamics-nav/reports/TestReport/50000/\">"
            + "<BCReportInformation><CreationDateTime>2026-01-01</CreationDateTime></BCReportInformation>"
            + "<Header><CompanyName>Contoso</CompanyName>"
            + "<Totals><Amount>0</Amount></Totals>"
            + "<Line><ItemNo_Line>X</ItemNo_Line></Line></Header>"
            + "</NavWordReportXmlPart>";

        var body = SyntheticLayout.RepeaterWithField(RepeaterXPath, FieldXPath, SyntheticLayout.GoodItemId);
        var layoutPath = SyntheticLayout.Create(body, datasetXml: schemaXml);

        try
        {
            var result = FullValidator.Full(layoutPath);

            Assert.DoesNotContain(result.Findings, f => f.Check == "unresolved-binding");
            Assert.Equal(0, result.ErrorCount);
            Assert.True(result.Passed,
                "expected pass; errors: "
                + string.Join(" | ", result.Findings.Where(f => f.Severity == FindingSeverity.Error).Select(f => f.Message)));

            // FullValidator's dry-run merge uses the default MergeOptions.Rows (3), so the repeater clones
            // three times and the structurally-divergent field falls back independently in each clone —
            // one xpath-fallback finding per row, every one a Warning, never an Error.
            var fallbacks = result.Findings.Where(f => f.Check == "xpath-fallback").ToList();
            Assert.NotEmpty(fallbacks);
            Assert.All(fallbacks, f => Assert.Equal(FindingSeverity.Warning, f.Severity));
        }
        finally
        {
            File.Delete(layoutPath);
        }
    }
}
