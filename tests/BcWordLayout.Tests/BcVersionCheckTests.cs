using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;

namespace BcWordLayout.Tests;

/// <summary>
/// The target BC version in the validator and the tools (GitHub issue #152): the <c>bc-version</c> check, the
/// <c>target</c> the read tools report, the post-edit validation of the mutating tools and <c>create_layout</c>, and
/// the <c>bc_version_unsupported</c> envelope.
/// </summary>
[Collection("bc-version-environment-seam")]
public sealed class BcVersionCheckTests : IDisposable
{
    private const string TenantReportDefaultsXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        + "<NavWordReportXmlPart xmlns=\"urn:microsoft-dynamics-nav/reports/Tenant_Report_Defaults/2000000001/\">"
        + "<BCReportInformation><ReportMetadata><ReportName>Standard Sales - Invoice</ReportName></ReportMetadata>"
        + "</BCReportInformation></NavWordReportXmlPart>";

    private const string Paragraph = "<w:p><w:r><w:t>Header</w:t></w:r></w:p>";

    public BcVersionCheckTests() => ToolGuards.BcVersionEnvironmentValue = null;

    public void Dispose() => ToolGuards.BcVersionEnvironmentValue = null;

    private static TargetVersion Known(int major, int minor) =>
        TargetVersion.FromEnvironment(new BcVersion(major, minor), $"{major}.{minor}", note: null);

    /// <summary>A header/footer layout for report themes: bound to Tenant_Report_Defaults, under <paramref name="repo"/>.</summary>
    private static string TenantReportDefaultsLayout(TargetFolder repo)
    {
        var synthetic = SyntheticLayout.Create(Paragraph, datasetXml: TenantReportDefaultsXml);
        try
        {
            return repo.CopyIn(synthetic, "layouts/HeaderFooter.docx");
        }
        finally
        {
            File.Delete(synthetic);
        }
    }

    private static List<ValidationFinding> BcVersionFindings(ValidationResult result) =>
        result.Findings.Where(f => f.Check == LayoutValidator.BcVersionCheck).ToList();

    // ---- the bc-version check ----

    [Fact]
    public void A_Tenant_Report_Defaults_layout_is_an_error_below_29()
    {
        using var repo = new TargetFolder();
        var path = TenantReportDefaultsLayout(repo);

        var finding = Assert.Single(BcVersionFindings(LayoutValidator.Quick(path, Known(28, 3))));

        Assert.Equal(FindingSeverity.Error, finding.Severity);
        Assert.Contains("bound to Tenant_Report_Defaults (report 2000000001)", finding.Message, StringComparison.Ordinal);
        Assert.Contains("('report-themes'), which Business Central has from BC 29.0 (2026 release wave 2)", finding.Message, StringComparison.Ordinal);
        Assert.Contains("The target is BC 28.3 (2026 release wave 1), from environment (BCWL_BC_VERSION=28.3)", finding.Message, StringComparison.Ordinal);
        Assert.Equal("urn:microsoft-dynamics-nav/reports/Tenant_Report_Defaults/2000000001/", finding.Location);
    }

    [Fact]
    public void A_Tenant_Report_Defaults_layout_is_a_warning_for_an_unknown_target()
    {
        using var repo = new TargetFolder();
        var path = TenantReportDefaultsLayout(repo);

        var finding = Assert.Single(BcVersionFindings(LayoutValidator.Quick(path, TargetVersion.Unknown("BCWL_BC_VERSION is not set"))));

        Assert.Equal(FindingSeverity.Warning, finding.Severity);
        Assert.Contains("The target BC version is unknown (BCWL_BC_VERSION is not set)", finding.Message, StringComparison.Ordinal);
        Assert.Contains("Set application in the AL project's app.json, or BCWL_BC_VERSION", finding.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_Tenant_Report_Defaults_layout_passes_from_29_and_without_a_target()
    {
        using var repo = new TargetFolder();
        var path = TenantReportDefaultsLayout(repo);

        Assert.Empty(BcVersionFindings(LayoutValidator.Quick(path, Known(29, 0))));
        Assert.Empty(BcVersionFindings(LayoutValidator.Quick(path, Known(30, 0))));
        Assert.Empty(BcVersionFindings(LayoutValidator.Quick(path)));
    }

    [Theory]
    [InlineData(Corpus.SalesInvoice)]
    [InlineData(Corpus.InventoryOrderDetails)]
    [InlineData(Corpus.StandardStatement)]
    [InlineData(Corpus.StandardSalesQuote)]
    [InlineData(Corpus.StandardPurchaseOrder)]
    [InlineData(Corpus.SalespersonCommission)]
    [InlineData(Corpus.SalesInvoiceVatSpec)]
    [InlineData(Corpus.JobQuote)]
    [InlineData(Corpus.PaymentPracticeByPeriod)]
    [InlineData(Corpus.SubcontractorDispatchList)]
    [InlineData(Corpus.SalesInvoiceBlueSimpleThemable)]
    [InlineData(Corpus.SalesShipmentBlueThemable)]
    [InlineData(Corpus.StandardPurchaseOrderThemable)]
    public void Theme_colours_and_paragraph_styles_already_in_a_layout_are_not_flagged(string fileName)
    {
        // BC 28 prints them, and Microsoft's own BC 28 layouts carry them; themable-formatting gates only what the
        // tools write. The three themable captures colour everything through accent1 and still pass.
        Assert.Empty(BcVersionFindings(LayoutValidator.Quick(Corpus.Path(fileName), Known(28, 0))));
        Assert.Empty(BcVersionFindings(LayoutValidator.Quick(Corpus.Path(fileName), TargetVersion.Unknown("not set"))));
    }

    // ---- the target the read tools report ----

    [Fact]
    public void get_layout_info_reports_the_target_from_the_nearest_app_json()
    {
        using var repo = new TargetFolder();
        var appJson = repo.AppJson("app", "29.0.0.0");
        var path = repo.CopyIn(Corpus.Path(Corpus.JobQuote), "app/layouts/JobQuote.docx");

        var response = ReadTools.GetLayoutInfo(path);

        Assert.True(response.Ok, response.Error?.Message);
        var target = Assert.IsType<LayoutInfoDto>(response.Data).Target;
        Assert.Equal(new TargetVersionDto("29.0", "app.json", $"{appJson}: application 29.0.0.0"), target);
    }

    [Fact]
    public void validate_layout_reports_an_unknown_target_with_the_reason()
    {
        using var repo = new TargetFolder();
        var path = repo.CopyIn(Corpus.Path(Corpus.JobQuote), "JobQuote.docx");

        var response = ReadTools.ValidateLayout(path);

        Assert.True(response.Ok, response.Error?.Message);
        var target = Assert.IsType<ValidationResultDto>(response.Data).Target;
        Assert.Null(target.BcVersion);
        Assert.Equal("unknown", target.Source);
        Assert.EndsWith("and BCWL_BC_VERSION is not set", target.Detail, StringComparison.Ordinal);
    }

    [Fact]
    public void The_environment_variable_names_the_target_outside_an_AL_project()
    {
        using var repo = new TargetFolder();
        var path = repo.CopyIn(Corpus.Path(Corpus.JobQuote), "JobQuote.docx");
        ToolGuards.BcVersionEnvironmentValue = "29";

        var target = Assert.IsType<ValidationResultDto>(ReadTools.ValidateLayout(path, "full").Data).Target;

        Assert.Equal(new TargetVersionDto("29.0", "environment", "BCWL_BC_VERSION=29"), target);
    }

    [Fact]
    public void validate_layout_fails_a_Tenant_Report_Defaults_layout_for_a_BC_28_project()
    {
        using var repo = new TargetFolder();
        repo.AppJson(".", "28.0.0.0");
        var path = TenantReportDefaultsLayout(repo);

        var quick = Assert.IsType<ValidationResultDto>(ReadTools.ValidateLayout(path).Data);
        var full = Assert.IsType<ValidationResultDto>(ReadTools.ValidateLayout(path, "full").Data);

        Assert.All(new[] { quick, full }, result =>
        {
            Assert.False(result.Passed);
            var finding = Assert.Single(result.Findings, f => f.Check == "bc-version");
            Assert.Equal("Error", finding.Severity);
        });
    }

    // ---- the validation a writing tool returns ----

    [Fact]
    public void A_mutating_tool_validates_against_the_target_of_the_layout_it_edits()
    {
        using var repo = new TargetFolder();
        var appJson = repo.AppJson(".", "28.0.0.0");
        var path = TenantReportDefaultsLayout(repo);

        var below = EditTools.InsertText(path, " / ", locationType: "documentEnd");
        Assert.True(below.Ok, below.Error?.Message); // reported, never refused: nothing shipped is gated
        Assert.False(Assert.IsType<EditResultDto>(below.Data).QuickValidation.Passed);

        File.WriteAllText(appJson, """{ "application": "29.0.0.0" }""");
        var from29 = EditTools.InsertText(path, " / ", locationType: "documentEnd");
        Assert.True(Assert.IsType<EditResultDto>(from29.Data).QuickValidation.Passed);
    }

    [Fact]
    public void create_layout_reports_the_target_of_its_output_folder_and_never_mistakes_it_for_a_stale_template()
    {
        using var repo = new TargetFolder();
        repo.AppJson("bc28", "28.0.0.0");
        repo.AppJson("bc29", "29.0.0.0");
        var schema = repo.Write("TenantReportDefaults.xml", TenantReportDefaultsXml);

        var below = LifecycleTools.CreateLayout(schema, repo.Path("bc28/HeaderFooter.docx"));
        var from29 = LifecycleTools.CreateLayout(schema, repo.Path("bc29/HeaderFooter.docx"));

        Assert.True(below.Ok, below.Error?.Message);
        var belowValidation = Assert.IsType<CreateResultDto>(below.Data).QuickValidation;
        Assert.Equal(1, belowValidation.ErrorCount);
        Assert.True(Assert.IsType<CreateResultDto>(from29.Data).QuickValidation.Passed);
    }

    // ---- the bc_version_unsupported envelope ----

    [Fact]
    public void A_gated_call_is_refused_with_bc_version_unsupported_and_a_hint_naming_both_ways_to_change_the_target()
    {
        using var repo = new TargetFolder();
        var appJson = repo.AppJson(".", "28.0.0.0");
        var path = repo.Path("Invoice.docx");

        var response = ToolGuards.Guard(() =>
        {
            BcFeatures.Require(BcFeatures.ThemableFormatting, ToolGuards.ResolveTarget(path));
            return ToolResponse.Success(new object());
        }, path);

        Assert.False(response.Ok);
        Assert.Equal(ErrorCodes.BcVersionUnsupported, response.Error!.Code);
        Assert.Contains("('themable-formatting')", response.Error.Message, StringComparison.Ordinal);
        Assert.StartsWith(
            $"'themable-formatting' needs BC 29.0 (2026 release wave 2) or later, and the target is BC 28.0 (2026 release "
            + $"wave 1), from app.json ({appJson}: application 28.0.0.0).",
            response.Error.Hint,
            StringComparison.Ordinal);
        Assert.Contains("`application` in the AL project's app.json", response.Error.Hint, StringComparison.Ordinal);
        Assert.Contains("BCWL_BC_VERSION in the server's env block", response.Error.Hint, StringComparison.Ordinal);
    }

    [Fact]
    public void An_unknown_target_is_refused_with_its_reason_in_the_hint()
    {
        using var repo = new TargetFolder();
        var path = repo.Path("Invoice.docx");

        var response = ToolGuards.Guard(() =>
        {
            BcFeatures.Require(BcFeatures.ReportThemes, ToolGuards.ResolveTarget(path));
            return ToolResponse.Success(new object());
        }, path);

        Assert.Equal(ErrorCodes.BcVersionUnsupported, response.Error!.Code);
        Assert.Contains("and the target is unknown: no app.json above the layout", response.Error.Hint, StringComparison.Ordinal);
    }
}
