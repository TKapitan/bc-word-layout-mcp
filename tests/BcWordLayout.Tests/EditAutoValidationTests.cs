using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;

namespace BcWordLayout.Tests;

/// <summary>
/// Cross-cutting guarantee: EVERY mutating tool - <c>insert_field</c>, <c>insert_label</c>,
/// <c>insert_repeater_table</c>, <c>remove_control</c>, and <c>create_layout</c> - returns a non-null,
/// populated <see cref="ValidationSummaryDto"/> (<c>QuickValidation</c>) on a successful call. This is a
/// STRUCTURAL guarantee, not a per-tool convention: the four editing tools (<see cref="EditTools"/>'s
/// <c>insert_field</c>/<c>insert_label</c>/<c>remove_control</c> and <see cref="TableTools"/>'s
/// <c>insert_repeater_table</c>) all route through <see cref="ToolGuards"/>'s single <c>GuardEdit</c>
/// helper, which is the only place <see cref="EditResultDto"/> is ever constructed (see its own
/// doc-comment); <c>create_layout</c> is likewise the only place
/// <see cref="CreateResultDto"/> is constructed, sourced from <c>LayoutBuilder.Create</c>'s own
/// <c>QuickValidation</c>. A future tool literally cannot report success without a populated summary
/// unless it bypasses these shared helpers entirely - this file locks that behavior down with one test
/// per mutating tool.
/// <para>
/// Per-operation mechanics (read-back shape, error-code mapping, file-untouched-on-failure, ...) are
/// covered elsewhere (<see cref="McpHostToolTests"/>, <c>LayoutEditorTests</c>, <c>RepeaterTableTests</c>,
/// <c>LayoutBuilderTests</c>); this file exists solely to lock down the auto-validation contract itself,
/// so it deliberately asserts nothing beyond "did this call succeed, and is QuickValidation populated and
/// sensible."
/// </para>
/// </summary>
public class EditAutoValidationTests
{
    private static string CopyOfCorpus(string corpusFile)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-autovalidation-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    /// <summary>
    /// The actual contract check, shared by every test below: non-null, level "quick", and sensible
    /// (non-negative, and - since every call below targets a known-clean operation - a real error-free)
    /// counts. <paramref name="quickValidation"/> is declared nullable here (even though the DTOs
    /// themselves declare it non-nullable) so this helper still means something as a runtime tripwire if
    /// that guarantee were ever weakened.
    /// </summary>
    private static void AssertPopulatedQuickValidation(ValidationSummaryDto? quickValidation)
    {
        Assert.NotNull(quickValidation);
        Assert.Equal("quick", quickValidation!.Level);
        Assert.True(quickValidation.ErrorCount >= 0);
        Assert.True(quickValidation.WarningCount >= 0);
        Assert.True(quickValidation.Passed, "expected a clean quick validation for this known-good operation");
        Assert.Equal(0, quickValidation.ErrorCount);
    }

    [Fact]
    public void InsertField_success_always_returns_a_populated_quickValidation_summary()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = EditTools.InsertField(path, "/Header/CustomerAddress1", "documentEnd");

            Assert.True(response.Ok);
            var dto = Assert.IsType<EditResultDto>(response.Data);
            AssertPopulatedQuickValidation(dto.QuickValidation);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertLabel_success_always_returns_a_populated_quickValidation_summary()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = EditTools.InsertLabel(path, "/Header/Contact_Lbl", "documentEnd");

            Assert.True(response.Ok);
            var dto = Assert.IsType<EditResultDto>(response.Data);
            AssertPopulatedQuickValidation(dto.QuickValidation);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertRepeaterTable_success_always_returns_a_populated_quickValidation_summary()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.InsertRepeaterTable(
                path, "/Header/Line", "ItemNo_Line,Description_Line,TransHeaderAmount", "documentEnd");

            Assert.True(response.Ok);
            var dto = Assert.IsType<EditResultDto>(response.Data);
            AssertPopulatedQuickValidation(dto.QuickValidation);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveControl_success_always_returns_a_populated_quickValidation_summary()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            // Insert a fresh control first so there is always something safe to remove regardless of the
            // corpus's own pre-existing content.
            var inserted = EditTools.InsertField(path, "/Header/CustomerAddress1", "documentEnd");
            var insertedDto = Assert.IsType<EditResultDto>(inserted.Data);

            var response = EditTools.RemoveControl(path, insertedDto.ControlId);

            Assert.True(response.Ok);
            var dto = Assert.IsType<EditResultDto>(response.Data);
            AssertPopulatedQuickValidation(dto.QuickValidation);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void CreateLayout_success_always_returns_a_populated_quickValidation_summary()
    {
        var outputPath = Path.Combine(Path.GetTempPath(), $"bcwl-autovalidation-create-{Guid.NewGuid():N}.docx");
        try
        {
            var response = LifecycleTools.CreateLayout(Corpus.Path(Corpus.SalesInvoice), outputPath);

            Assert.True(response.Ok);
            var dto = Assert.IsType<CreateResultDto>(response.Data);
            AssertPopulatedQuickValidation(dto.QuickValidation);
        }
        finally
        {
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }
}
