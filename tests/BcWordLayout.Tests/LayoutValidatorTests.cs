using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using BcWordLayout.Merge;

namespace BcWordLayout.Tests;

public class LayoutValidatorTests
{
    [Theory]
    [InlineData(Corpus.SalesInvoice)]
    [InlineData(Corpus.InventoryOrderDetails)]
    [InlineData(Corpus.StandardStatement)]
    [InlineData(Corpus.SalespersonCommission)]
    [InlineData(Corpus.JobQuote)]
    [InlineData(Corpus.PaymentPracticeByPeriod)]
    [InlineData(Corpus.SubcontractorDispatchList)]
    public void All_corpus_layouts_quick_validate_with_zero_errors(string fileName)
    {
        var result = LayoutValidator.Quick(Corpus.Path(fileName));

        Assert.Equal("quick", result.Level);
        Assert.True(result.Passed,
            "expected pass; errors: " +
            string.Join(" | ", result.Findings.Where(f => f.Severity == FindingSeverity.Error).Select(f => f.Message)));
        Assert.Equal(0, result.ErrorCount);
    }

    [Fact]
    public void SalesInvoice_surfaces_attachedTemplate_as_warning_not_error()
    {
        var result = LayoutValidator.Quick(Corpus.Path(Corpus.SalesInvoice));

        var warning = result.Findings.SingleOrDefault(f => f.Check == "attached-template");
        Assert.NotNull(warning);
        Assert.Equal(FindingSeverity.Warning, warning!.Severity);
        Assert.True(result.Passed);
    }

    [Fact]
    public void PaymentPractice_orphaned_bindings_are_reported_instead_of_passing_silently()
    {
        // The B36 regression against the real layout that exposed it. This base-app layout binds 20 of its
        // 25 controls to the report's superseded namespace (Payment_Practice/590) while its embedded part
        // declares Payment_Practice/685, and its BC part carries no DataStoreItem, so every binding's
        // storeItemID is unverifiable too. Before the fix this validated as "passed, 0 errors, 0 warnings".
        var result = LayoutValidator.Quick(Corpus.Path(Corpus.PaymentPracticeByPeriod));

        var nsFindings = result.Findings.Where(f => f.Check == "binding-namespace").ToList();
        Assert.Equal(20, nsFindings.Count);
        Assert.All(nsFindings, f => Assert.Equal(FindingSeverity.Warning, f.Severity));
        Assert.All(nsFindings, f => Assert.Contains("Payment_Practice/590", f.Message));

        Assert.Contains(result.Findings, f => f.Check == "single-bc-part"
                                              && f.Severity == FindingSeverity.Warning);

        // All warnings: the layout is Microsoft's own and BC prints it, so it must not be called invalid.
        Assert.Equal(0, result.ErrorCount);
        Assert.True(result.Passed);
    }

    [Fact]
    public void A_layout_whose_bindings_all_name_its_own_namespace_gets_no_namespace_warnings()
    {
        // Non-tautology guard for the test above: the check must be capable of staying silent on a real
        // layout, or "20 findings" would prove nothing but that the check fires indiscriminately.
        var result = LayoutValidator.Quick(Corpus.Path(Corpus.JobQuote));

        Assert.Empty(result.Findings.Where(f => f.Check == "binding-namespace"));
        Assert.Empty(result.Findings.Where(f => f.Check == "single-bc-part"));
        Assert.True(result.Passed);
    }

    [Fact]
    public void Full_level_runs_a_real_dry_run_merge_and_returns_a_findings_list()
    {
        // "full" is now implemented (BcWordLayout.Merge.FullValidator) rather than a placeholder — this
        // just proves the plumbing: level is "full", and Findings is a real (possibly empty) list rather
        // than the fixed one-item "not implemented" placeholder the old version of this test targeted.
        var result = FullValidator.Full(Corpus.Path(Corpus.SalesInvoice));

        Assert.Equal("full", result.Level);
        Assert.DoesNotContain(result.Findings, f => f.Check == "full-not-implemented");
        Assert.True(result.Passed,
            "expected the known-good corpus layout to full-validate clean; errors: " +
            string.Join(" | ", result.Findings.Where(f => f.Severity == FindingSeverity.Error).Select(f => f.Message)));
    }

    [Fact]
    public void Missing_file_throws_file_not_found()
    {
        Assert.Throws<FileNotFoundException>(() => LayoutValidator.Quick("Z:\\does-not-exist.docx"));
    }
}
