using System.Text;
using BcWordLayout.Merge;
using BcWordLayout.Render;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace BcWordLayout.Tests;

/// <summary>
/// Fidelity-artifact harness (see docs/FIDELITY-CHECKLIST.md): for each of the 3 corpus
/// layouts, batch-produces the mock-preview artifacts a human reviewer needs for the MANUAL fidelity
/// regression — a merged .docx and, when a PDF converter is available on this machine, a preview.pdf —
/// and asserts the invariants that CAN be automated: the merge actually
/// produced output, zero bindings are unresolved against the known-good corpus, and the merged document
/// stays <see cref="OpenXmlValidator"/>-clean. Deliberately dependency-agnostic, mirroring
/// <see cref="PdfConverterContractTests"/>: whichever converter branch runs (or none), the suite stays
/// fully green. PDF conversion goes entirely through <see cref="PdfConverterFactory"/>/
/// <see cref="WordComConverter"/>, which already guarantees no zombie WINWORD process is left behind on
/// its own (see that type's own remarks) — this class never spawns Word directly and adds no cleanup of
/// its own on top.
/// <para>
/// Output goes to <c>%TEMP%/bcwl-fidelity-output/</c> by default, which keeps the suite hermetic — no
/// writes into the repo tree, so a clean clone and CI behave identically. The human-review workflow
/// opts into a stable location by setting <c>BCWL_FIDELITY_OUTPUT_DIR</c> (conventionally
/// <c>&lt;repo&gt;/fidelity-output</c>, which stays gitignored). Wherever it lands, the output is never
/// deleted or reset by this class, unlike every other test's temp-file teardown — it IS the
/// human-review artifact and must survive the test run; each re-run simply overwrites it.
/// </para>
/// </summary>
public class FidelityHarnessTests
{
    /// <summary>Opt-in override for where the human-review artifacts land (see class remarks).</summary>
    internal const string OutputDirVariable = "BCWL_FIDELITY_OUTPUT_DIR";

    private static string ResolveOutputRoot()
    {
        var overridden = Environment.GetEnvironmentVariable(OutputDirVariable);
        return string.IsNullOrWhiteSpace(overridden)
            ? Path.Combine(Path.GetTempPath(), "bcwl-fidelity-output")
            : Path.GetFullPath(overridden);
    }

    /// <summary>
    /// Fixed disclaimer written into every per-layout summary, for a human reading these artifacts
    /// directly rather than through the <c>preview_layout</c> MCP tool response (whose own
    /// <c>PreviewDisclaimer</c> in <c>LifecycleTools.cs</c> says the same thing).
    /// </summary>
    private const string MockRenderDisclaimer =
        "MOCK RENDER — not a substitute for a BC sandbox render. Sample data is generated (not real "
        + "Business Central data) and conversion happens outside the BC report engine; captions, fonts, "
        + "pagination, and other BC-specific rendering behavior may differ. These artifacts only cover "
        + "the AUTOMATED fidelity dimensions (see docs/FIDELITY-CHECKLIST.md) — final sign-off is always "
        + "a real Business Central sandbox render.";

    /// <summary>
    /// Strips a legacy "Template-" prefix (kept for compatibility with any older-style corpus file name)
    /// and the ".docx" extension — e.g. "InventoryOrderDetails.docx" becomes "InventoryOrderDetails" — for
    /// a human-friendly output subfolder name under <c>fidelity-output/</c>.
    /// </summary>
    private static string LayoutFolderName(string corpusFileName)
    {
        var name = Path.GetFileNameWithoutExtension(corpusFileName);
        return name.StartsWith("Template-", StringComparison.Ordinal) ? name["Template-".Length..] : name;
    }

    [Theory]
    [InlineData(Corpus.SalesInvoice)]
    [InlineData(Corpus.InventoryOrderDetails)]
    [InlineData(Corpus.StandardStatement)]
    public void Corpus_layout_produces_fidelity_artifacts_and_meets_automatable_invariants(string corpusFileName)
    {
        var layoutName = LayoutFolderName(corpusFileName);
        var outputDir = Path.Combine(ResolveOutputRoot(), layoutName);
        Directory.CreateDirectory(outputDir);

        var mergedDocxPath = Path.Combine(outputDir, "merged.docx");

        // Same defaults preview_layout itself uses (LifecycleTools.PreviewLayout: rows=3, seed=12345) so
        // these artifacts look like what an agent normally gets back from that tool.
        var mergeResult = MergeEngine.Merge(
            Corpus.Path(corpusFileName), mergedDocxPath, new MergeOptions { Seed = 12345, Rows = 3 });

        // ---- automatable invariant: merge produced output ----
        Assert.True(File.Exists(mergedDocxPath),
            $"[{layoutName}] MergeEngine.Merge did not produce '{mergedDocxPath}'.");

        // ---- automatable invariant: zero unresolved bindings on the known-good corpus ----
        Assert.Equal(0, mergeResult.Stats.Unresolved);

        // ---- automatable invariant: merged docx stays OpenXmlValidator-clean ----
        // (SalesInvoiceForSubscriptionBilling's PaymentServiceLogo and StandardStatement's CompanyPicture
        // each live inside a body repeater that Rows=3 clones - exercising the clone-id fix, which gives
        // every cloned row's wp:docPr / bookmark w:id its own fresh, unique value.)
        using (var doc = WordprocessingDocument.Open(mergedDocxPath, false))
        {
            var errors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc).ToList();
            Assert.True(errors.Count == 0,
                $"[{layoutName}] expected 0 OpenXmlValidator errors on the merged docx; found: "
                + string.Join(" | ", errors.Select(e => $"{e.Path?.XPath}: {e.Description}")));
        }

        // ---- PDF artifact: produced only when a converter is actually available on this machine ----
        var converter = PdfConverterFactory.Select(PdfConverterKind.Auto);
        var pdfPath = Path.Combine(outputDir, "preview.pdf");
        var pdfProduced = false;
        string? pdfSkippedReason = null;

        if (converter.IsAvailable)
        {
            var conversion = converter.Convert(mergedDocxPath, pdfPath);
            Assert.True(conversion.Ok,
                $"[{layoutName}] converter '{converter.Name}' reported available but the conversion "
                + $"failed: {conversion.Error}");
            Assert.True(File.Exists(pdfPath),
                $"[{layoutName}] converter reported success but '{pdfPath}' is missing.");

            var bytes = File.ReadAllBytes(pdfPath);
            Assert.True(bytes.Length > 1024,
                $"[{layoutName}] expected a non-trivial PDF (>1 KB), was {bytes.Length} bytes.");
            Assert.Equal("%PDF", Encoding.ASCII.GetString(bytes, 0, 4));
            pdfProduced = true;
        }
        else
        {
            // No PDF converter on this machine (neither Word COM nor LibreOffice): the merged docx alone
            // is still the deliverable for the manual comparison — never fail the suite over a missing
            // local install (dependency-agnostic, matching PdfConverterContractTests).
            Assert.True(File.Exists(mergedDocxPath));
            pdfSkippedReason = "no PDF converter available on this machine (neither Word COM nor LibreOffice)";
        }

        var summaryPath = Path.Combine(outputDir, "SUMMARY.md");
        WriteSummary(summaryPath, layoutName, corpusFileName, converter, pdfProduced, pdfSkippedReason, mergeResult);

        // The summary is itself a deliverable of this harness — confirm it was actually written and
        // still carries the fixed mock-render disclaimer verbatim, so a future edit to WriteSummary can't
        // silently drop it.
        Assert.True(File.Exists(summaryPath));
        var summaryText = File.ReadAllText(summaryPath);
        Assert.Contains("MOCK RENDER", summaryText, StringComparison.Ordinal);
        Assert.Contains("not a substitute for a BC sandbox render", summaryText, StringComparison.Ordinal);
    }

    /// <summary>
    /// Builds and writes the per-layout human-readable summary: converter used/availability, whether a
    /// PDF was produced or skipped (and why), merge stats, every merge warning (or "(none)"), and the
    /// fixed <see cref="MockRenderDisclaimer"/>. Pure I/O — asserts nothing itself; the caller verifies
    /// the result.
    /// </summary>
    private static void WriteSummary(
        string summaryPath,
        string layoutName,
        string corpusFileName,
        IPdfConverter converter,
        bool pdfProduced,
        string? pdfSkippedReason,
        MergeResult mergeResult)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"# Fidelity harness summary — {layoutName}");
        sb.AppendLine();
        sb.AppendLine($"Generated (UTC): {DateTime.UtcNow:O}");
        sb.AppendLine($"Corpus file: {corpusFileName}");
        sb.AppendLine();
        sb.AppendLine(MockRenderDisclaimer);
        sb.AppendLine();
        sb.AppendLine("## Converter");
        sb.AppendLine($"- Selected: {converter.Name}");
        sb.AppendLine($"- Available on this machine: {converter.IsAvailable}");
        sb.AppendLine(pdfProduced ? "- PDF: produced (preview.pdf)" : $"- PDF: skipped — {pdfSkippedReason}");
        sb.AppendLine();
        sb.AppendLine("## Merge stats");
        sb.AppendLine($"- FieldsFilled: {mergeResult.Stats.FieldsFilled}");
        sb.AppendLine($"- RepeatersExpanded: {mergeResult.Stats.RepeatersExpanded}");
        sb.AppendLine($"- RowsGenerated: {mergeResult.Stats.RowsGenerated}");
        sb.AppendLine($"- Unresolved: {mergeResult.Stats.Unresolved}");
        sb.AppendLine($"- PicturesFilled: {mergeResult.Stats.PicturesFilled}");
        sb.AppendLine();
        sb.AppendLine("## Merge warnings");
        if (mergeResult.Warnings.Count == 0)
        {
            sb.AppendLine("(none)");
        }
        else
        {
            foreach (var warning in mergeResult.Warnings)
            {
                var location = warning.Location is null ? string.Empty : $" ({warning.Location})";
                sb.AppendLine($"- [{warning.Kind}] {warning.Message}{location}");
            }
        }

        File.WriteAllText(summaryPath, sb.ToString());
    }
}
