using System.Xml.Linq;
using BcWordLayout.Domain;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using BcWordLayout.Render;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <summary>
/// Covers the MCP tool surface (<see cref="ReadTools"/>, <see cref="EditTools"/>, <see cref="TableTools"/>,
/// <see cref="LifecycleTools"/>): the DTO mapping and the <c>Guard()</c> error-envelope translation. Tools
/// are called directly in-process (not over the MCP wire) against the corpus files, asserting both the
/// success shape ({ok, data}) and the failure shape ({ok, error}).
/// </summary>
/// <remarks>Joins the preview-converter-seam collection: this class holds every test that SWAPS
/// <c>LifecycleTools.SelectConverter</c> for a fake, plus real-converter <c>preview_layout</c> tests that
/// read it (see <see cref="PreviewConverterSeamCollection"/> for the rule).</remarks>
[Collection("preview-converter-seam")]
public class McpHostToolTests
{
    [Fact]
    public void GetLayoutInfo_returns_ok_with_expected_shape()
    {
        var response = ReadTools.GetLayoutInfo(Corpus.Path(Corpus.SalesInvoice));

        Assert.True(response.Ok);
        Assert.Null(response.Error);

        var dto = Assert.IsType<LayoutInfoDto>(response.Data);
        Assert.Equal("1306", dto.Report.ReportId);
        Assert.Equal("Standard_Sales_Invoice", dto.Report.ReportName);
        Assert.Equal("quick", dto.Validation.Level);
        Assert.True(dto.Validation.Passed);
        Assert.NotEmpty(dto.Controls);
        Assert.Equal(dto.Controls.Count, dto.ControlSummary.Total);
        Assert.Contains("document.xml", dto.Parts);

        // Every control reports a structural level; the header address fields are cell-level and located
        // in a table.
        Assert.All(dto.Controls, c => Assert.False(string.IsNullOrEmpty(c.Level)));
        var addr6 = dto.Controls.Single(c => c.SdtId == -2064325541); // #Nav: /Header/CustomerAddress6
        Assert.Equal("cell", addr6.Level);
        Assert.NotNull(addr6.TableIndex);

        // The tables section is populated and consistent with the cell-level control's own coordinates.
        Assert.NotEmpty(dto.Tables);
        var owningCell = dto.Tables
            .SelectMany(t => t.Rows)
            .SelectMany(r => r.Cells)
            .SingleOrDefault(cell => cell.ControlId == -2064325541);
        Assert.NotNull(owningCell);
        Assert.True(owningCell!.IsControlCell);
    }

    [Fact]
    public void ValidateLayout_quick_returns_ok_and_passes_for_corpus()
    {
        var response = ReadTools.ValidateLayout(Corpus.Path(Corpus.InventoryOrderDetails), "quick");

        Assert.True(response.Ok);
        var dto = Assert.IsType<ValidationResultDto>(response.Data);
        Assert.Equal("quick", dto.Level);
        Assert.True(dto.Passed);
        Assert.Equal(0, dto.ErrorCount);
    }

    [Fact]
    public void ValidateLayout_unknown_level_returns_invalid_argument()
    {
        var response = ReadTools.ValidateLayout(Corpus.Path(Corpus.SalesInvoice), "bogus");

        Assert.False(response.Ok);
        Assert.NotNull(response.Error);
        Assert.Equal("invalid_argument", response.Error!.Code);
        Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));
    }

    [Fact]
    public void ValidateLayout_full_returns_ok_and_passes_for_corpus()
    {
        var response = ReadTools.ValidateLayout(Corpus.Path(Corpus.SalesInvoice), "full");

        Assert.True(response.Ok);
        var dto = Assert.IsType<ValidationResultDto>(response.Data);
        Assert.Equal("full", dto.Level);
        Assert.True(dto.Passed,
            "expected pass; errors: " + string.Join(" | ",
                dto.Findings.Where(f => f.Severity == "Error").Select(f => f.Message)));
        Assert.Equal(0, dto.ErrorCount);
    }

    [Fact]
    public void PreviewLayout_returns_ok_with_merged_docx_stats_disclaimer_and_quick_validation()
    {
        var outputDir = Path.Combine(Path.GetTempPath(), $"bcwl-preview-test-{Guid.NewGuid():N}");

        try
        {
            var response = LifecycleTools.PreviewLayout(Corpus.Path(Corpus.SalesInvoice), outputDir: outputDir);

            Assert.True(response.Ok);
            Assert.Null(response.Error);
            var dto = Assert.IsType<PreviewResultDto>(response.Data);

            Assert.True(File.Exists(dto.MergedDocxPath), $"expected merged docx to exist at '{dto.MergedDocxPath}'");
            Assert.True(dto.Stats.FieldsFilled > 0, "expected at least one field to be filled");
            Assert.True(dto.Stats.RepeatersExpanded > 0, "expected at least one repeater to expand");
            Assert.True(dto.Stats.RowsGenerated > 0, "expected at least one row to be generated");
            Assert.DoesNotContain(dto.Warnings, w => w.Kind == "unresolved-binding");
            Assert.False(string.IsNullOrWhiteSpace(dto.Disclaimer));
            Assert.Equal("quick", dto.QuickValidation.Level);
            Assert.True(dto.QuickValidation.Passed);
            Assert.Equal(0, dto.QuickValidation.ErrorCount);

            // Dependency-agnostic: assert a real successful PDF only when a converter is actually available
            // on this machine (Word is installed here); otherwise the tool must still report Ok=true overall
            // with a clean, explained conversion failure and the merged docx still present.
            if (PdfConverterFactory.Select(PdfConverterKind.Auto).IsAvailable)
            {
                Assert.True(dto.ConversionOk, $"expected conversion to succeed but got: {dto.ConversionError}");
                Assert.NotNull(dto.PdfPath);
                Assert.True(File.Exists(dto.PdfPath!), $"expected pdf to exist at '{dto.PdfPath}'");

                var bytes = File.ReadAllBytes(dto.PdfPath!);
                Assert.True(bytes.Length > 4, "expected a non-trivial pdf");
                Assert.Equal("%PDF", System.Text.Encoding.ASCII.GetString(bytes, 0, 4));
            }
            else
            {
                Assert.False(dto.ConversionOk);
                Assert.Null(dto.PdfPath);
                Assert.False(string.IsNullOrWhiteSpace(dto.ConversionError));
                Assert.True(File.Exists(dto.MergedDocxPath), "merged docx must survive a conversion failure");
            }
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
    public void PreviewLayout_unknown_converter_returns_invalid_argument()
    {
        // No outputDir needed: converter validation happens before any directory/file is touched, so there
        // is nothing for this test to clean up.
        var response = LifecycleTools.PreviewLayout(Corpus.Path(Corpus.SalesInvoice), converter: "bogus");

        Assert.False(response.Ok);
        Assert.NotNull(response.Error);
        Assert.Equal("invalid_argument", response.Error!.Code);
        Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));
    }

    [Fact]
    public void PreviewLayout_without_outputDir_defaults_under_a_hashed_per_layout_temp_subfolder()
    {
        // No explicit outputDir: exercises the default "%TEMP%\bc-word-layout-mcp\<basename>-<hash>"
        // branch (LifecycleTools.DefaultPreviewOutputDirName), which every other preview test bypasses by
        // always passing its own outputDir.
        var response = LifecycleTools.PreviewLayout(Corpus.Path(Corpus.SalesInvoice));

        Assert.True(response.Ok);
        var dto = Assert.IsType<PreviewResultDto>(response.Data);

        var mergedDir = Path.GetDirectoryName(dto.MergedDocxPath);
        try
        {
            Assert.True(File.Exists(dto.MergedDocxPath), $"expected merged docx to exist at '{dto.MergedDocxPath}'");

            var expectedRoot = Path.Combine(Path.GetTempPath(), "bc-word-layout-mcp");
            Assert.NotNull(mergedDir);
            Assert.StartsWith(
                Path.GetFullPath(expectedRoot) + Path.DirectorySeparatorChar,
                Path.GetFullPath(mergedDir!) + Path.DirectorySeparatorChar,
                StringComparison.OrdinalIgnoreCase);

            // The subfolder name itself is exactly LifecycleTools.DefaultPreviewOutputDirName's output - the
            // fixed "merged.docx"/"preview.pdf" file names inside it stay unhashed (see PreviewLayout: the
            // hash lives in the FOLDER name here, and only moves onto the FILE names when outputDir is
            // explicit).
            var expectedDirName = LifecycleTools.DefaultPreviewOutputDirName(Corpus.Path(Corpus.SalesInvoice));
            Assert.Equal(expectedDirName, Path.GetFileName(mergedDir));
            Assert.Equal("merged.docx", Path.GetFileName(dto.MergedDocxPath));
        }
        finally
        {
            if (mergedDir is not null && Directory.Exists(mergedDir))
            {
                Directory.Delete(mergedDir, recursive: true);
            }
        }
    }

    [Fact]
    public void PreviewPathHash_same_path_different_case_produces_the_same_hash()
    {
        var lower = Corpus.Path(Corpus.SalesInvoice).ToLowerInvariant();
        var upper = Corpus.Path(Corpus.SalesInvoice).ToUpperInvariant();

        Assert.Equal(LifecycleTools.PreviewPathHash(lower), LifecycleTools.PreviewPathHash(upper));
    }

    [Fact]
    public void PreviewPathHash_same_path_via_a_non_normalized_dotdot_segment_produces_the_same_hash()
    {
        var direct = Corpus.Path(Corpus.SalesInvoice);
        var directory = Path.GetDirectoryName(direct)!;
        var viaDotDot = Path.Combine(directory, "some-other-dir", "..", Path.GetFileName(direct));

        Assert.Equal(LifecycleTools.PreviewPathHash(direct), LifecycleTools.PreviewPathHash(viaDotDot));
    }

    [Fact]
    public void PreviewPathHash_different_full_paths_sharing_a_basename_produce_different_hashes()
    {
        var pathA = @"C:\appA\SalesInvoice.docx";
        var pathB = @"C:\appB\SalesInvoice.docx";

        Assert.NotEqual(LifecycleTools.PreviewPathHash(pathA), LifecycleTools.PreviewPathHash(pathB));
    }

    [Fact]
    public void PreviewPathHash_produces_twelve_lowercase_hex_characters()
    {
        var hash = LifecycleTools.PreviewPathHash(Corpus.Path(Corpus.SalesInvoice));

        Assert.Equal(12, hash.Length);
        Assert.Matches("^[0-9a-f]{12}$", hash);
    }

    [Fact]
    public void DefaultPreviewOutputDirName_combines_the_sanitized_basename_with_the_path_hash()
    {
        var layoutPath = Corpus.Path(Corpus.SalesInvoice);
        var expectedBasename = Path.GetFileNameWithoutExtension(layoutPath);
        var expectedHash = LifecycleTools.PreviewPathHash(layoutPath);

        Assert.Equal($"{expectedBasename}-{expectedHash}", LifecycleTools.DefaultPreviewOutputDirName(layoutPath));
    }

    [Fact]
    public void DefaultPreviewOutputDirName_two_layouts_sharing_a_basename_produce_different_dir_names()
    {
        // Path.Combine (not a hardcoded C:\ literal) so the same-basename-different-directory shape
        // holds on every OS - backslashes are not separators on POSIX.
        var root = Path.GetTempPath();
        var dirNameA = LifecycleTools.DefaultPreviewOutputDirName(Path.Combine(root, "appA", "SalesInvoice.docx"));
        var dirNameB = LifecycleTools.DefaultPreviewOutputDirName(Path.Combine(root, "appB", "SalesInvoice.docx"));

        Assert.NotEqual(dirNameA, dirNameB);
        Assert.StartsWith("SalesInvoice-", dirNameA);
        Assert.StartsWith("SalesInvoice-", dirNameB);
    }

    [Fact]
    public void PreviewLayout_explicit_outputDir_keeps_two_layouts_sharing_a_basename_from_colliding()
    {
        // Two DIFFERENT layouts that happen to share a file name ("Invoice.docx" under two different
        // directories), both previewed into the SAME caller-chosen outputDir: PreviewLayout must key the
        // merged/pdf FILE names by each layout's own path hash so the second preview's output doesn't
        // clobber the first's (the explicit-outputDir case).
        var root = Path.Combine(Path.GetTempPath(), $"bcwl-preview-collision-{Guid.NewGuid():N}");
        var dirA = Path.Combine(root, "appA");
        var dirB = Path.Combine(root, "appB");
        Directory.CreateDirectory(dirA);
        Directory.CreateDirectory(dirB);
        var layoutA = Path.Combine(dirA, "Invoice.docx");
        var layoutB = Path.Combine(dirB, "Invoice.docx");
        File.Copy(Corpus.Path(Corpus.SalesInvoice), layoutA);
        File.Copy(Corpus.Path(Corpus.InventoryOrderDetails), layoutB);
        var outputDir = Path.Combine(root, "shared-output");

        try
        {
            var responseA = LifecycleTools.PreviewLayout(layoutA, outputDir: outputDir);
            var responseB = LifecycleTools.PreviewLayout(layoutB, outputDir: outputDir);

            Assert.True(responseA.Ok);
            Assert.True(responseB.Ok);
            var dtoA = Assert.IsType<PreviewResultDto>(responseA.Data);
            var dtoB = Assert.IsType<PreviewResultDto>(responseB.Data);

            Assert.NotEqual(dtoA.MergedDocxPath, dtoB.MergedDocxPath);
            Assert.Equal(Path.GetFullPath(outputDir), Path.GetDirectoryName(dtoA.MergedDocxPath));
            Assert.Equal(Path.GetFullPath(outputDir), Path.GetDirectoryName(dtoB.MergedDocxPath));

            // The first preview's output must still be intact - not overwritten by the second.
            Assert.True(File.Exists(dtoA.MergedDocxPath), "layout A's merged docx must survive layout B's preview");
            Assert.True(File.Exists(dtoB.MergedDocxPath), "layout B's merged docx must exist");
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
    public void PreviewLayout_serializes_against_a_concurrent_holder_of_the_same_layouts_edit_lock()
    {
        // Ordering signal, not wall-clock inference: while another thread holds the SAME lock object
        // PreviewLayout itself acquires (ToolGuards.EditLockFor), PreviewLayout cannot have entered its
        // critical section, so the merged-output file it would create cannot exist yet. We poll for that
        // file's absence for a few seconds WHILE the lock is held, rather than inferring "still blocked"
        // from a short Join() timeout - a Join-based wait is unreliable in the failure direction: on a
        // Word-equipped machine an UNLOCKED preview can take several real seconds (genuine COM conversion),
        // so a short Join() timing out proves nothing about whether the lock actually blocked it.
        // Pinning converter to "libreoffice" (rather than the "auto" default) sidesteps that entirely -
        // conversion happens AFTER the merge and is irrelevant to the file-existence check below, but it
        // also keeps the post-release completion phase fast/deterministic instead of risking a real Word
        // COM conversion.
        //
        // Correctness of the poll-while-held assertion: a held Monitor lock blocks the second thread
        // unconditionally, so on CORRECT code the merged file never appears during the poll window no
        // matter how long that window is - this can never flake false-positive (fail correct code). It can
        // only FALSE-PASS (miss a real regression) if an UNLOCKED merge takes LONGER than the poll window
        // to write the file; a bare merge of one of the small corpus layouts (no PDF conversion involved
        // yet) reliably completes in well under a second, so the several-second poll window below
        // comfortably exceeds that and the test does actually fail if the lock is removed.
        var layoutPath = Corpus.Path(Corpus.SalesInvoice);
        var outputDir = Path.Combine(Path.GetTempPath(), $"bcwl-preview-lock-test-{Guid.NewGuid():N}");
        var expectedMergedPath = Path.Combine(outputDir, $"merged-{LifecycleTools.PreviewPathHash(layoutPath)}.docx");
        var blockerHasLock = new ManualResetEventSlim(false);
        var releaseBlocker = new ManualResetEventSlim(false);
        ToolResponse? previewResponse = null;

        var blockerThread = new Thread(() =>
        {
            lock (ToolGuards.EditLockFor(layoutPath))
            {
                blockerHasLock.Set();
                releaseBlocker.Wait(TimeSpan.FromSeconds(15));
            }
        });
        blockerThread.Start();

        try
        {
            Assert.True(blockerHasLock.Wait(TimeSpan.FromSeconds(5)), "blocker thread failed to acquire the edit lock in time");

            var previewThread = new Thread(() =>
            {
                previewResponse = LifecycleTools.PreviewLayout(layoutPath, converter: "libreoffice", outputDir: outputDir);
            });
            previewThread.Start();

            var pollUntilUtc = DateTime.UtcNow + TimeSpan.FromSeconds(3);
            while (DateTime.UtcNow < pollUntilUtc)
            {
                Assert.False(
                    File.Exists(expectedMergedPath),
                    "merged output must not exist while a concurrent caller still holds the edit lock - "
                    + "PreviewLayout must still be blocked");
                Thread.Sleep(50);
            }

            releaseBlocker.Set();
            Assert.True(previewThread.Join(TimeSpan.FromSeconds(30)), "PreviewLayout should complete once the edit lock is released");
            blockerThread.Join(TimeSpan.FromSeconds(5));

            Assert.NotNull(previewResponse);
            Assert.True(previewResponse!.Ok);
            Assert.True(File.Exists(expectedMergedPath), "merged output should exist once PreviewLayout has run to completion");
        }
        finally
        {
            releaseBlocker.Set();
            blockerThread.Join(TimeSpan.FromSeconds(5));
            if (Directory.Exists(outputDir))
            {
                Directory.Delete(outputDir, recursive: true);
            }
        }
    }

    // ---- preview_layout: deterministic converter-selection tests ----
    //
    // The tests above this point exercise preview_layout against whatever converter is REALLY installed on
    // the machine running the suite (dependency-agnostic, but still machine-dependent for which BRANCH
    // runs). The tests below instead swap LifecycleTools.SelectConverter for a FakePdfConverter, so every
    // branch of the tool's conversion-outcome handling (converterAvailable / conversionOk / conversionError
    // / pdfPath) is exercised deterministically regardless of what is installed. Each test restores the seam
    // to PdfConverterFactory.Select in a finally block; see LifecycleTools.SelectConverter's own remarks for
    // why every such test must live in THIS class (xUnit's default per-class collection keeps them
    // sequential relative to each other and to the real-converter tests above, so this static never races).

    [Fact]
    public void SelectConverter_defaults_to_PdfConverterFactory_Select_for_every_kind()
    {
        // Regression guard for the isolation invariant itself: if an earlier test in this class forgot to
        // restore the seam in its finally block, this test (and the real-converter tests) would resolve the
        // wrong converter TYPE for at least one kind.
        foreach (PdfConverterKind kind in Enum.GetValues<PdfConverterKind>())
        {
            Assert.Equal(PdfConverterFactory.Select(kind).GetType(), LifecycleTools.SelectConverter(kind).GetType());
        }
    }

    [Fact]
    public void PreviewLayout_with_a_fake_available_converter_reports_conversionOk_and_actually_invokes_it()
    {
        var outputDir = Path.Combine(Path.GetTempPath(), $"bcwl-preview-fake-ok-{Guid.NewGuid():N}");
        var fake = new FakePdfConverter { Name = "fake-ok", IsAvailable = true, ConversionSucceeds = true };
        LifecycleTools.SelectConverter = _ => fake;

        try
        {
            var response = LifecycleTools.PreviewLayout(Corpus.Path(Corpus.SalesInvoice), outputDir: outputDir);

            Assert.True(response.Ok);
            var dto = Assert.IsType<PreviewResultDto>(response.Data);

            Assert.Equal("fake-ok", dto.ConverterUsed);
            Assert.True(dto.ConverterAvailable);
            Assert.True(dto.ConversionOk);
            Assert.NotNull(dto.PdfPath);
            Assert.True(File.Exists(dto.PdfPath!), $"expected fake pdf to exist at '{dto.PdfPath}'");
            Assert.Null(dto.ConversionError);

            // Actually invoked, with the real merged working copy - not just an outcome that happens to
            // match by coincidence.
            Assert.Equal(1, fake.ConvertCallCount);
            Assert.Equal(dto.MergedDocxPath, fake.LastDocxPath);
            Assert.Equal(dto.PdfPath, fake.LastPdfPath);
            Assert.True(File.Exists(fake.LastDocxPath!), "the merged docx handed to the converter must actually exist");
        }
        finally
        {
            LifecycleTools.SelectConverter = PdfConverterFactory.Select;
            if (Directory.Exists(outputDir))
            {
                Directory.Delete(outputDir, recursive: true);
            }
        }
    }

    [Fact]
    public void PreviewLayout_with_a_fake_unavailable_converter_still_succeeds_with_no_pdf()
    {
        var outputDir = Path.Combine(Path.GetTempPath(), $"bcwl-preview-fake-unavail-{Guid.NewGuid():N}");
        var fake = new FakePdfConverter
        {
            Name = "fake-unavailable",
            IsAvailable = false,
            ConversionSucceeds = false,
            ConversionErrorMessage = "fake converter's dependency is not installed",
        };
        LifecycleTools.SelectConverter = _ => fake;

        try
        {
            var response = LifecycleTools.PreviewLayout(Corpus.Path(Corpus.SalesInvoice), outputDir: outputDir);

            // The tool call itself still succeeds (ok:true) even though no PDF could be produced - matching
            // every real-converter test above's "no converter available" branch.
            Assert.True(response.Ok);
            var dto = Assert.IsType<PreviewResultDto>(response.Data);

            Assert.Equal("fake-unavailable", dto.ConverterUsed);
            Assert.False(dto.ConverterAvailable);
            Assert.False(dto.ConversionOk);
            Assert.Null(dto.PdfPath);
            Assert.Equal("fake converter's dependency is not installed", dto.ConversionError);
            Assert.True(File.Exists(dto.MergedDocxPath), "merged docx must survive even when no converter is available");
            Assert.Equal(1, fake.ConvertCallCount);
        }
        finally
        {
            LifecycleTools.SelectConverter = PdfConverterFactory.Select;
            if (Directory.Exists(outputDir))
            {
                Directory.Delete(outputDir, recursive: true);
            }
        }
    }

    [Fact]
    public void PreviewLayout_with_a_fake_conversion_failure_reports_conversionOk_false_with_the_error()
    {
        var outputDir = Path.Combine(Path.GetTempPath(), $"bcwl-preview-fake-fail-{Guid.NewGuid():N}");
        // IsAvailable=true but Convert itself fails - distinguishes "no converter installed" (previous test)
        // from "a converter IS installed but this specific conversion attempt failed" (e.g. a real Word/
        // LibreOffice crash or timeout); the tool must surface this identically either way.
        var fake = new FakePdfConverter
        {
            Name = "fake-flaky",
            IsAvailable = true,
            ConversionSucceeds = false,
            ConversionErrorMessage = "simulated conversion crash",
        };
        LifecycleTools.SelectConverter = _ => fake;

        try
        {
            var response = LifecycleTools.PreviewLayout(Corpus.Path(Corpus.SalesInvoice), outputDir: outputDir);

            Assert.True(response.Ok);
            var dto = Assert.IsType<PreviewResultDto>(response.Data);

            Assert.Equal("fake-flaky", dto.ConverterUsed);
            Assert.True(dto.ConverterAvailable);
            Assert.False(dto.ConversionOk);
            Assert.Null(dto.PdfPath);
            Assert.Equal("simulated conversion crash", dto.ConversionError);
            Assert.True(File.Exists(dto.MergedDocxPath), "merged docx must survive a failed conversion attempt");
            Assert.Equal(1, fake.ConvertCallCount);
        }
        finally
        {
            LifecycleTools.SelectConverter = PdfConverterFactory.Select;
            if (Directory.Exists(outputDir))
            {
                Directory.Delete(outputDir, recursive: true);
            }
        }
    }

    /// <summary>Recursively collects every <see cref="ExternalRelationship"/> reachable from
    /// <paramref name="main"/>'s own part tree - the same traversal
    /// <see cref="BcWordLayout.Domain.LayoutValidator"/>'s attachedTemplate check and
    /// <c>ExternalRelationshipStripper</c> both use - so a test can assert none remain anywhere in a merged
    /// package, not just on the main document part itself.</summary>
    private static List<ExternalRelationship> AllExternalRelationships(OpenXmlPart part, HashSet<OpenXmlPart>? visited = null)
    {
        visited ??= new HashSet<OpenXmlPart>();
        if (!visited.Add(part))
        {
            return new List<ExternalRelationship>();
        }

        var found = part.ExternalRelationships.ToList();
        foreach (var child in part.Parts)
        {
            found.AddRange(AllExternalRelationships(child.OpenXmlPart, visited));
        }

        return found;
    }

    [Fact]
    public void PreviewLayout_strips_the_real_corpus_attachedTemplate_relationship_and_leaves_the_original_untouched()
    {
        // SalesInvoiceForSubscriptionBilling.docx carries a REAL external attachedTemplate relationship
        // (see LayoutValidatorTests.SalesInvoice_surfaces_attachedTemplate_as_warning_not_error) - a stale
        // developer template path, exactly the shape a poisoned
        // layout could weaponize with a UNC path or URL instead. This proves preview_layout's merged output
        // never carries it (or ANY external relationship) into the copy handed to a converter, and that the
        // ORIGINAL layout file is never touched in the process - the never-open-the-original-writable
        // invariant.
        var outputDir = Path.Combine(Path.GetTempPath(), $"bcwl-preview-strip-{Guid.NewGuid():N}");
        var fake = new FakePdfConverter { Name = "fake-strip-check" };
        LifecycleTools.SelectConverter = _ => fake;

        var originalPath = Corpus.Path(Corpus.SalesInvoice);
        var originalBytesBefore = File.ReadAllBytes(originalPath);

        try
        {
            var response = LifecycleTools.PreviewLayout(originalPath, outputDir: outputDir);

            Assert.True(response.Ok);
            var dto = Assert.IsType<PreviewResultDto>(response.Data);

            var warning = Assert.Single(dto.Warnings, w => w.Kind == "external-relationship-stripped");
            Assert.Contains("attachedTemplate", warning.Message);

            using (var mergedDoc = WordprocessingDocument.Open(dto.MergedDocxPath, false))
            {
                var remaining = AllExternalRelationships(mergedDoc.MainDocumentPart!);
                Assert.Empty(remaining);

                var settings = mergedDoc.MainDocumentPart!.DocumentSettingsPart?.Settings;
                Assert.True(
                    settings is null || !settings.Elements<AttachedTemplate>().Any(),
                    "expected no dangling w:attachedTemplate element in the merged settings part");

                // SalesInvoiceForSubscriptionBilling's PaymentServiceLogo picture lives inside a repeater
                // that this preview's Rows=3 clones - exercising the clone-id fix (MergeEngine now
                // regenerates each cloned row's wp:docPr id), unrelated to (and unaffected by) external-
                // relationship stripping, which is what this test itself is actually proving.
                var errors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator(FileFormatVersions.Office2019)
                    .Validate(mergedDoc)
                    .ToList();
                Assert.True(errors.Count == 0,
                    "expected zero validation errors in the merged output; found: "
                    + string.Join(" | ", errors.Select(e => $"{e.Path?.XPath}: {e.Description}")));
            }

            // Never-open-the-original-writable invariant: the source corpus file is byte-identical.
            var originalBytesAfter = File.ReadAllBytes(originalPath);
            Assert.Equal(originalBytesBefore, originalBytesAfter);
        }
        finally
        {
            LifecycleTools.SelectConverter = PdfConverterFactory.Select;
            if (Directory.Exists(outputDir))
            {
                Directory.Delete(outputDir, recursive: true);
            }
        }
    }

    // ---- preview_layout retention sweep ----
    // All four tests below seed real filesystem state directly under the tool's actual default preview
    // root ("%TEMP%\bc-word-layout-mcp") - the same root LifecycleTools.PreviewLayout itself sweeps - since
    // the sweep is private to LifecycleTools and only observable through its filesystem side effects. Safe
    // to share with every other preview test in this class: this class is the sole
    // "preview-converter-seam" collection member exercising the default root, so xUnit serializes it against
    // every other test here, and the sweep only ever deletes directories OLDER than the retention window -
    // never anything another test in this run just created.
    //
    // Folders meant to simulate a REAL earlier preview_layout output use
    // LifecycleTools.DefaultPreviewOutputDirName's own naming shape (via a throwaway fake layout path) -
    // since the sweep now only ever considers folders shaped like the tool's own output, a folder named
    // anything else would never even become a delete candidate, regardless of
    // age, and wouldn't actually exercise the age-check/delete-failure paths these tests target.

    private static string FakeToolOwnedPreviewDirName(string discriminator) =>
        LifecycleTools.DefaultPreviewOutputDirName(
            Path.Combine(Path.GetTempPath(), $"bcwl-retention-fake-{discriminator}-{Guid.NewGuid():N}.docx"));

    [Fact]
    public void PreviewLayout_sweeps_a_stale_sibling_dir_under_the_default_root()
    {
        var defaultRoot = Path.Combine(Path.GetTempPath(), "bc-word-layout-mcp");
        Directory.CreateDirectory(defaultRoot);
        var staleDir = Path.Combine(defaultRoot, FakeToolOwnedPreviewDirName("stale"));
        Directory.CreateDirectory(staleDir);
        File.WriteAllText(Path.Combine(staleDir, "leftover.txt"), "old preview output");
        Directory.SetLastWriteTimeUtc(staleDir, DateTime.UtcNow - TimeSpan.FromDays(10));

        var fake = new FakePdfConverter { Name = "fake-sweep-stale" };
        LifecycleTools.SelectConverter = _ => fake;
        // This call's OWN output goes to its own explicit outputDir - proving the sweep runs (and reaches
        // the default root) even when the triggering call itself never touches that root.
        var callOutputDir = Path.Combine(Path.GetTempPath(), $"bcwl-preview-sweep-call-{Guid.NewGuid():N}");

        try
        {
            var response = LifecycleTools.PreviewLayout(Corpus.Path(Corpus.SalesInvoice), outputDir: callOutputDir);

            Assert.True(response.Ok);
            Assert.False(Directory.Exists(staleDir),
                "a stale sibling dir under the default root should be swept by a preview call");
        }
        finally
        {
            LifecycleTools.SelectConverter = PdfConverterFactory.Select;
            if (Directory.Exists(staleDir))
            {
                Directory.Delete(staleDir, recursive: true);
            }

            if (Directory.Exists(callOutputDir))
            {
                Directory.Delete(callOutputDir, recursive: true);
            }
        }
    }

    [Fact]
    public void PreviewLayout_never_sweeps_a_fresh_sibling_dir_under_the_default_root()
    {
        var defaultRoot = Path.Combine(Path.GetTempPath(), "bc-word-layout-mcp");
        Directory.CreateDirectory(defaultRoot);
        var freshDir = Path.Combine(defaultRoot, FakeToolOwnedPreviewDirName("fresh"));
        Directory.CreateDirectory(freshDir);
        File.WriteAllText(Path.Combine(freshDir, "recent.txt"), "recent preview output");
        // No explicit SetLastWriteTimeUtc: just-created, so its LastWriteTimeUtc is "now" - well inside the
        // retention window - the natural "someone previewed this five minutes ago" shape.

        var fake = new FakePdfConverter { Name = "fake-sweep-fresh" };
        LifecycleTools.SelectConverter = _ => fake;
        var callOutputDir = Path.Combine(Path.GetTempPath(), $"bcwl-preview-sweep-fresh-call-{Guid.NewGuid():N}");

        try
        {
            var response = LifecycleTools.PreviewLayout(Corpus.Path(Corpus.SalesInvoice), outputDir: callOutputDir);

            Assert.True(response.Ok);
            Assert.True(Directory.Exists(freshDir),
                "a fresh sibling dir under the default root must survive a preview call");
        }
        finally
        {
            LifecycleTools.SelectConverter = PdfConverterFactory.Select;
            if (Directory.Exists(freshDir))
            {
                Directory.Delete(freshDir, recursive: true);
            }

            if (Directory.Exists(callOutputDir))
            {
                Directory.Delete(callOutputDir, recursive: true);
            }
        }
    }

    [Fact]
    public void PreviewLayout_never_sweeps_a_caller_supplied_outputDir_or_its_neighbors()
    {
        var root = Path.Combine(Path.GetTempPath(), $"bcwl-preview-outputdir-sweep-{Guid.NewGuid():N}");
        var outputDir = Path.Combine(root, "caller-output");
        Directory.CreateDirectory(outputDir);
        // A sibling dir NEXT TO outputDir (same parent - NOT under the tool's default root at all), made to
        // look exactly as stale as the sweep would otherwise remove: proves the sweep never reaches outside
        // the default root, regardless of how old anything near a caller-supplied outputDir looks.
        var staleLookingNeighbor = Path.Combine(root, "stale-looking-neighbor");
        Directory.CreateDirectory(staleLookingNeighbor);
        File.WriteAllText(Path.Combine(staleLookingNeighbor, "unrelated.txt"), "not the tool's default root");
        Directory.SetLastWriteTimeUtc(staleLookingNeighbor, DateTime.UtcNow - TimeSpan.FromDays(30));
        Directory.SetLastWriteTimeUtc(outputDir, DateTime.UtcNow - TimeSpan.FromDays(30));

        // Review follow-up: a stale, ARBITRARILY-NAMED folder that happens to sit DIRECTLY INSIDE
        // the tool's real default root (as if a caller had once pointed their own outputDir there) must
        // ALSO survive - not because it's "this call's own directory" (it isn't; this call's own output goes
        // to outputDir above, elsewhere entirely) but because its name doesn't match
        // LooksLikeToolOwnedPreviewDirName's "<basename>-<12 lowercase hex>" shape. Before the fix, only the
        // this-call exclusion protected a folder from a sweep, which only ever covers the CURRENT call's own
        // directory - an unrelated stale folder like this one, living in the same root, would have been
        // deleted by this very call's sweep pass.
        var defaultRoot = Path.Combine(Path.GetTempPath(), "bc-word-layout-mcp");
        Directory.CreateDirectory(defaultRoot);
        var callerNamedDirInsideDefaultRoot = Path.Combine(defaultRoot, $"caller_owned_folder_{Guid.NewGuid():N}");
        Directory.CreateDirectory(callerNamedDirInsideDefaultRoot);
        File.WriteAllText(
            Path.Combine(callerNamedDirInsideDefaultRoot, "caller-data.txt"),
            "a caller's own folder that happens to live inside the tool's default root");
        Directory.SetLastWriteTimeUtc(callerNamedDirInsideDefaultRoot, DateTime.UtcNow - TimeSpan.FromDays(30));

        var fake = new FakePdfConverter { Name = "fake-outputdir-sweep" };
        LifecycleTools.SelectConverter = _ => fake;

        try
        {
            var response = LifecycleTools.PreviewLayout(Corpus.Path(Corpus.SalesInvoice), outputDir: outputDir);

            Assert.True(response.Ok);
            Assert.True(Directory.Exists(outputDir),
                "a caller-supplied outputDir must never be swept, however old it looks");
            Assert.True(Directory.Exists(staleLookingNeighbor),
                "a directory next to a caller-supplied outputDir sits outside the default root and must never be touched");
            Assert.True(Directory.Exists(callerNamedDirInsideDefaultRoot),
                "a stale, arbitrarily-named caller folder living inside the default root must survive - its "
                + "name doesn't match the tool's own output-folder shape, so it must never be a sweep candidate");
        }
        finally
        {
            LifecycleTools.SelectConverter = PdfConverterFactory.Select;
            if (Directory.Exists(callerNamedDirInsideDefaultRoot))
            {
                Directory.Delete(callerNamedDirInsideDefaultRoot, recursive: true);
            }

            if (Directory.Exists(root))
            {
                Directory.Delete(root, recursive: true);
            }
        }
    }

    [WindowsOnlyFact]
    public void PreviewLayout_succeeds_even_when_a_stale_sibling_dir_fails_to_delete()
    {
        var defaultRoot = Path.Combine(Path.GetTempPath(), "bc-word-layout-mcp");
        Directory.CreateDirectory(defaultRoot);
        var lockedStaleDir = Path.Combine(defaultRoot, FakeToolOwnedPreviewDirName("locked"));
        Directory.CreateDirectory(lockedStaleDir);
        var lockedFilePath = Path.Combine(lockedStaleDir, "locked.bin");
        File.WriteAllBytes(lockedFilePath, new byte[] { 1, 2, 3 });
        Directory.SetLastWriteTimeUtc(lockedStaleDir, DateTime.UtcNow - TimeSpan.FromDays(10));

        var fake = new FakePdfConverter { Name = "fake-sweep-locked" };
        LifecycleTools.SelectConverter = _ => fake;
        var callOutputDir = Path.Combine(Path.GetTempPath(), $"bcwl-preview-sweep-locked-call-{Guid.NewGuid():N}");

        try
        {
            // Hold the file open with FileShare.None so Directory.Delete(recursive: true) on lockedStaleDir
            // throws mid-sweep (another process holding a file open is exactly the real-world shape this
            // guards against) - the "using" scope ends (releasing the lock) before the finally cleanup below
            // runs, so this test can still tidy up after itself.
            using (new FileStream(lockedFilePath, FileMode.Open, FileAccess.Read, FileShare.None))
            {
                var response = LifecycleTools.PreviewLayout(Corpus.Path(Corpus.SalesInvoice), outputDir: callOutputDir);

                Assert.True(response.Ok,
                    "a stale sibling dir failing to delete must never fail the preview_layout call itself");
                var dto = Assert.IsType<PreviewResultDto>(response.Data);
                Assert.True(File.Exists(dto.MergedDocxPath));

                // The locked stale dir survives (its delete genuinely failed) - proving the sweep actually
                // attempted and failed here, rather than this test passing for an unrelated reason.
                Assert.True(Directory.Exists(lockedStaleDir),
                    "the locked stale dir should still exist since deleting it failed");
            }
        }
        finally
        {
            LifecycleTools.SelectConverter = PdfConverterFactory.Select;
            if (Directory.Exists(lockedStaleDir))
            {
                Directory.Delete(lockedStaleDir, recursive: true);
            }

            if (Directory.Exists(callOutputDir))
            {
                Directory.Delete(callOutputDir, recursive: true);
            }
        }
    }

    [Fact]
    public void ListDatasetFields_from_layout_returns_ok_with_bound_flags()
    {
        var response = ReadTools.ListDatasetFields(Corpus.Path(Corpus.SalesInvoice));

        Assert.True(response.Ok);
        var dto = Assert.IsType<DatasetFieldsDto>(response.Data);
        Assert.Equal("layout", dto.SourceType);
        Assert.Equal("NavWordReportXmlPart", dto.Root.Name);
    }

    [Fact]
    public void GetLayoutInfo_missing_file_returns_file_not_found_with_hint()
    {
        var response = ReadTools.GetLayoutInfo("Z:\\does-not-exist.docx");

        Assert.False(response.Ok);
        Assert.Null(response.Data);
        Assert.NotNull(response.Error);
        Assert.Equal("file_not_found", response.Error!.Code);
        Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));
    }

    [Fact]
    public void ValidateLayout_missing_file_returns_file_not_found_with_hint()
    {
        var response = ReadTools.ValidateLayout("Z:\\does-not-exist.docx");

        Assert.False(response.Ok);
        Assert.NotNull(response.Error);
        Assert.Equal("file_not_found", response.Error!.Code);
        Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));
    }

    [Fact]
    public void PreviewLayout_missing_file_returns_file_not_found_with_hint()
    {
        var response = LifecycleTools.PreviewLayout("Z:\\does-not-exist.docx");

        Assert.False(response.Ok);
        Assert.NotNull(response.Error);
        Assert.Equal("file_not_found", response.Error!.Code);
        Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));
    }

    // ---- create_layout ----
    // Every success test below writes to its own temp output path (this tool CREATES a file rather than
    // editing one in place) and deletes it afterward.

    private static string NewLayoutOutputPath() =>
        Path.Combine(Path.GetTempPath(), $"bcwl-createlayout-{Guid.NewGuid():N}.docx");

    [Fact]
    public void CreateLayout_from_a_corpus_docx_returns_ok_with_populated_fields_and_the_output_reopens()
    {
        var outputPath = NewLayoutOutputPath();
        try
        {
            var response = LifecycleTools.CreateLayout(Corpus.Path(Corpus.SalesInvoice), outputPath);

            Assert.True(response.Ok);
            Assert.Null(response.Error);
            var dto = Assert.IsType<CreateResultDto>(response.Data);
            Assert.Equal(Path.GetFullPath(outputPath), dto.OutputPath);
            Assert.Equal("Standard_Sales_Invoice", dto.ReportName);
            Assert.Equal("1306", dto.ReportId);
            Assert.False(string.IsNullOrWhiteSpace(dto.Namespace));
            Assert.False(string.IsNullOrWhiteSpace(dto.StoreItemId));
            Assert.False(dto.UsedTemplate);
            Assert.False(dto.ReplacedExistingBcPart);
            Assert.Equal("quick", dto.QuickValidation.Level);
            Assert.True(dto.QuickValidation.Passed);
            Assert.Equal(0, dto.QuickValidation.ErrorCount);

            Assert.True(File.Exists(outputPath));

            // Reopens via a fresh tool call and reads back the same identity/storeItemID, passing validation.
            var reopened = ReadTools.GetLayoutInfo(outputPath);
            Assert.True(reopened.Ok);
            var info = Assert.IsType<LayoutInfoDto>(reopened.Data);
            Assert.Equal("Standard_Sales_Invoice", info.Report.ReportName);
            Assert.Equal("1306", info.Report.ReportId);
            Assert.Equal(dto.StoreItemId, info.Report.StoreItemId);
            Assert.True(info.Validation.Passed);
            Assert.Equal(0, info.Validation.ErrorCount);
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
    public void CreateLayout_output_is_immediately_editable_via_insert_field_with_a_passing_quickValidation()
    {
        var outputPath = NewLayoutOutputPath();
        try
        {
            var created = LifecycleTools.CreateLayout(Corpus.Path(Corpus.SalesInvoice), outputPath);
            Assert.True(created.Ok);

            var response = EditTools.InsertField(outputPath, "/Header/CustomerAddress1", "documentEnd");

            Assert.True(response.Ok);
            var dto = Assert.IsType<EditResultDto>(response.Data);
            Assert.Equal("Field", dto.Kind);
            Assert.True(dto.QuickValidation.Passed);
            Assert.Equal(0, dto.QuickValidation.ErrorCount);
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
    public void CreateLayout_with_templatePath_that_is_a_full_BC_layout_returns_template_not_unbound_and_writes_nothing()
    {
        var outputPath = NewLayoutOutputPath();
        try
        {
            // StandardStatement (a fully-populated real layout, not an unbound branded shell) is deliberately
            // used as the "template" here: create_layout must REFUSE rather than
            // silently ship a layout whose pre-existing bound controls are now stale.
            var response = LifecycleTools.CreateLayout(
                Corpus.Path(Corpus.SalesInvoice), outputPath, Corpus.Path(Corpus.StandardStatement));

            Assert.False(response.Ok);
            Assert.Null(response.Data);
            Assert.NotNull(response.Error);
            Assert.Equal("template_not_unbound", response.Error!.Code);
            Assert.Contains("full BC layout", response.Error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));
            Assert.Contains("refresh_xml_part", response.Error.Hint, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("remove_control", response.Error.Hint, StringComparison.OrdinalIgnoreCase);

            // The atomic build must leave NOTHING behind on a refusal - no outputPath, no stray temp file.
            Assert.False(File.Exists(outputPath));
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
    public void CreateLayout_with_a_full_BC_layout_templatePath_leaves_a_preexisting_outputPath_byte_identical()
    {
        // Mirrors CreateLayout_from_a_non_BC_xml_leaves_a_preexisting_outputPath_byte_identical, but for the
        // template_not_unbound refusal: a pre-existing file at outputPath must survive this
        // failure byte-for-byte too, not just "no output written when there was nothing there before".
        var outputPath = NewLayoutOutputPath();
        try
        {
            var before = "pre-existing content that must survive"u8.ToArray();
            File.WriteAllBytes(outputPath, before);

            var response = LifecycleTools.CreateLayout(
                Corpus.Path(Corpus.SalesInvoice), outputPath, Corpus.Path(Corpus.StandardStatement));

            Assert.False(response.Ok);
            Assert.Equal("template_not_unbound", response.Error!.Code);
            Assert.Equal(before, File.ReadAllBytes(outputPath));
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
    public void CreateLayout_with_templatePath_that_has_a_BC_part_but_no_bound_controls_succeeds()
    {
        // The edge case the refusal must not trip on: a template whose BC part is present but
        // has nothing bound to it. Built synthetically via create_layout itself with no templatePath - a
        // freshly created layout's body carries only a heading paragraph, so it is guaranteed to have a real
        // BC part and zero content controls.
        var templatePath = NewLayoutOutputPath();
        var outputPath = NewLayoutOutputPath();
        try
        {
            var templateResponse = LifecycleTools.CreateLayout(Corpus.Path(Corpus.SalesInvoice), templatePath);
            Assert.True(templateResponse.Ok);
            var templateDto = Assert.IsType<CreateResultDto>(templateResponse.Data);

            var response = LifecycleTools.CreateLayout(
                Corpus.Path(Corpus.SalesInvoice), outputPath, templatePath);

            Assert.True(response.Ok);
            var dto = Assert.IsType<CreateResultDto>(response.Data);
            Assert.True(dto.UsedTemplate);
            Assert.True(dto.ReplacedExistingBcPart);
            Assert.NotEqual(templateDto.StoreItemId, dto.StoreItemId, StringComparer.OrdinalIgnoreCase);

            Assert.Equal("quick", dto.QuickValidation.Level);
            Assert.True(dto.QuickValidation.Passed);
            Assert.Equal(0, dto.QuickValidation.ErrorCount);

            // Still editable: a fresh control binds to the new storeItemID and reads back correctly.
            var edit = EditTools.InsertField(outputPath, "/Header/CustomerAddress1", "documentEnd");
            Assert.True(edit.Ok);
            var editDto = Assert.IsType<EditResultDto>(edit.Data);
            Assert.Equal("Field", editDto.Kind);

            var afterEdit = ReadTools.GetLayoutInfo(outputPath);
            var info = Assert.IsType<LayoutInfoDto>(afterEdit.Data);
            var insertedControl = Assert.Single(info.Controls, c => c.SdtId == editDto.ControlId);
            Assert.Equal("Field", insertedControl.Kind);
            Assert.Equal(dto.StoreItemId, insertedControl.StoreItemId, StringComparer.OrdinalIgnoreCase);
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
    public void CreateLayout_missing_schemaSource_returns_file_not_found_with_hint()
    {
        var response = LifecycleTools.CreateLayout("Z:\\does-not-exist.docx", NewLayoutOutputPath());

        Assert.False(response.Ok);
        Assert.NotNull(response.Error);
        Assert.Equal("file_not_found", response.Error!.Code);
        Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));
    }

    [Fact]
    public void CreateLayout_missing_templatePath_returns_file_not_found_with_hint()
    {
        var response = LifecycleTools.CreateLayout(
            Corpus.Path(Corpus.SalesInvoice), NewLayoutOutputPath(), "Z:\\does-not-exist.docx");

        Assert.False(response.Ok);
        Assert.NotNull(response.Error);
        Assert.Equal("file_not_found", response.Error!.Code);
        Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));
    }

    [Fact]
    public void CreateLayout_from_a_non_BC_xml_returns_invalid_layout()
    {
        var badXmlPath = Path.Combine(Path.GetTempPath(), $"bcwl-createlayout-badschema-{Guid.NewGuid():N}.xml");
        try
        {
            File.WriteAllText(badXmlPath, "<SomeOtherRoot xmlns=\"urn:not-bc\"><Foo/></SomeOtherRoot>");

            var response = LifecycleTools.CreateLayout(badXmlPath, NewLayoutOutputPath());

            Assert.False(response.Ok);
            Assert.NotNull(response.Error);
            Assert.Equal("invalid_layout", response.Error!.Code);
            Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));
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
    public void CreateLayout_from_a_non_BC_xml_leaves_a_preexisting_outputPath_byte_identical()
    {
        // Mirrors LayoutBuilderTests' Domain-level atomic-write proof (LayoutBuilder.Create reads
        // schemaSource to completion, and fails, before outputPath is ever touched) but through the MCP
        // tool surface specifically: a pre-existing file at outputPath must survive a failed create_layout
        // call byte-for-byte, not just "no stray temp file left behind".
        var badXmlPath = Path.Combine(Path.GetTempPath(), $"bcwl-createlayout-badschema-{Guid.NewGuid():N}.xml");
        var outputPath = NewLayoutOutputPath();
        try
        {
            File.WriteAllText(badXmlPath, "<SomeOtherRoot xmlns=\"urn:not-bc\"><Foo/></SomeOtherRoot>");
            var before = "pre-existing content that must survive"u8.ToArray();
            File.WriteAllBytes(outputPath, before);

            var response = LifecycleTools.CreateLayout(badXmlPath, outputPath);

            Assert.False(response.Ok);
            Assert.Equal("invalid_layout", response.Error!.Code);
            Assert.Equal(before, File.ReadAllBytes(outputPath));
        }
        finally
        {
            if (File.Exists(badXmlPath))
            {
                File.Delete(badXmlPath);
            }

            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }

    // ---- insert_field / insert_label / remove_control ----
    // Every test below runs against its own temp COPY of a corpus file (never the shared corpus itself,
    // since these tools write in place) and deletes it afterward.

    private static string CopyOfCorpus(string corpusFile)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-mcptool-edit-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    [Fact]
    public void InsertField_returns_ok_persists_to_disk_and_includes_a_passing_quickValidation()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = EditTools.InsertField(path, "/Header/CustomerAddress1", "documentEnd");

            Assert.True(response.Ok);
            Assert.Null(response.Error);
            var dto = Assert.IsType<EditResultDto>(response.Data);
            Assert.Equal("insert_field", dto.Operation);
            Assert.Equal("Field", dto.Kind);
            Assert.Equal("document.xml", dto.Part);
            Assert.False(string.IsNullOrWhiteSpace(dto.Summary));
            Assert.Equal("quick", dto.QuickValidation.Level);
            Assert.True(dto.QuickValidation.Passed);
            Assert.Equal(0, dto.QuickValidation.ErrorCount);

            // Persisted to disk: reopen (a fresh tool call, not the same handle) and confirm it's really there.
            var reopened = ReadTools.GetLayoutInfo(path);
            var info = Assert.IsType<LayoutInfoDto>(reopened.Data);
            Assert.Contains(info.Controls, c => c.SdtId == dto.ControlId && c.Kind == "Field" && c.XPath == dto.XPath);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertLabel_returns_ok_persists_to_disk_and_includes_a_passing_quickValidation()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = EditTools.InsertLabel(path, "/Header/Contact_Lbl", "documentEnd");

            Assert.True(response.Ok);
            Assert.Null(response.Error);
            var dto = Assert.IsType<EditResultDto>(response.Data);
            Assert.Equal("insert_label", dto.Operation);
            Assert.Equal("Label", dto.Kind);
            Assert.True(dto.QuickValidation.Passed);

            var reopened = ReadTools.GetLayoutInfo(path);
            var info = Assert.IsType<LayoutInfoDto>(reopened.Data);
            Assert.Contains(info.Controls, c => c.SdtId == dto.ControlId && c.Kind == "Label");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertField_sweeps_a_stale_orphaned_stage_file_in_the_layouts_directory()
    {
        // a hard kill between GuardMutate's stage-copy and its rename leaves a .bcwl-stage-*.docx
        // behind next to the layout - it never self-heals since nothing else revisits it. An isolated
        // directory (not the shared CopyOfCorpus temp root) so this test's planted file can't be seen -
        // or swept - by a concurrently-running test's own mutating call against the same shared folder.
        // Also plants a stale .bcwl-merge-stage-*.docx (interaction) to
        // prove ToolGuards' sweep glob list was broadened to catch that shape too, for the case where a
        // caller happens to point a merge/preview outputDir at a layout's own directory.
        var dir = Path.Combine(Path.GetTempPath(), $"bcwl-stage-sweep-{Guid.NewGuid():N}");
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "layout.docx");
        File.Copy(Corpus.Path(Corpus.SalesInvoice), path, overwrite: true);
        // CreationTimeUtc, not LastWriteTimeUtc, is the sweep's age signal (a File.Copy from an old
        // checked-in corpus file - as CopyOfCorpus/this test's own layout.docx copy above both are -
        // preserves the SOURCE's last-write time, so backdating LastWrite here would prove nothing; only
        // CreationTime reliably reflects how long the STAGING FILE ITSELF has existed).
        var staleStage = Path.Combine(dir, $".bcwl-stage-{Guid.NewGuid():N}.docx");
        File.WriteAllText(staleStage, "orphaned mid-commit artifact from a hard-killed process");
        File.SetCreationTimeUtc(staleStage, DateTime.UtcNow - TimeSpan.FromDays(2));
        var staleMergeStage = Path.Combine(dir, $".bcwl-merge-stage-{Guid.NewGuid():N}.docx");
        File.WriteAllText(staleMergeStage, "orphaned mid-merge artifact from a hard-killed process");
        File.SetCreationTimeUtc(staleMergeStage, DateTime.UtcNow - TimeSpan.FromDays(2));

        try
        {
            var response = EditTools.InsertField(path, "/Header/CustomerAddress1", "documentEnd");

            Assert.True(response.Ok);
            Assert.False(File.Exists(staleStage), "a stale .bcwl-stage-* file must be swept by a later mutating call");
            Assert.False(File.Exists(staleMergeStage), "a stale .bcwl-merge-stage-* file must also be swept");
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void InsertField_never_sweeps_a_fresh_stage_file_in_the_layouts_directory()
    {
        // Symmetric guard: a FRESH .bcwl-stage-* file (well inside the retention window - the shape a
        // genuinely concurrent in-flight commit would have) must never be swept, so a live commit -
        // this process serialized on the same path, or another process entirely mutating a DIFFERENT
        // layout that happens to share this directory - can never be mistaken for an orphan. Same for a
        // fresh .bcwl-merge-stage-*.docx (a live in-flight merge/preview writing into this directory).
        var dir = Path.Combine(Path.GetTempPath(), $"bcwl-stage-sweep-fresh-{Guid.NewGuid():N}");
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "layout.docx");
        File.Copy(Corpus.Path(Corpus.SalesInvoice), path, overwrite: true);
        var freshStage = Path.Combine(dir, $".bcwl-stage-{Guid.NewGuid():N}.docx");
        File.WriteAllText(freshStage, "a live in-flight commit's staged file");
        var freshMergeStage = Path.Combine(dir, $".bcwl-merge-stage-{Guid.NewGuid():N}.docx");
        File.WriteAllText(freshMergeStage, "a live in-flight merge's staged file");
        // No SetCreationTimeUtc on either: just-created, well inside the retention window.

        try
        {
            var response = EditTools.InsertField(path, "/Header/CustomerAddress1", "documentEnd");

            Assert.True(response.Ok);
            Assert.True(File.Exists(freshStage), "a fresh .bcwl-stage-* file must survive a mutating call");
            Assert.True(File.Exists(freshMergeStage), "a fresh .bcwl-merge-stage-* file must also survive");
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void InsertField_afterControl_targets_a_real_control_by_id()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            // Prove the flat afterControl params reach LayoutEditor by targeting a SAFE control: a field
            // first inserted at documentEnd sits in its own body paragraph (not inside a plain-text control),
            // so inserting another field after it does not nest a control inside a plain-text one. (Targeting
            // a cell-level plain-text control instead is now correctly REFUSED - see the regression test
            // InsertField_afterControl_into_a_plaintext_cell_control_is_refused_and_leaves_the_file_untouched.)
            var seed = EditTools.InsertField(path, "/Header/SalesPersonName", "documentEnd");
            var seedDto = Assert.IsType<EditResultDto>(seed.Data);

            var response = EditTools.InsertField(
                path, "/Header/CustomerAddress1", "afterControl", controlId: seedDto.ControlId);

            Assert.True(response.Ok);
            var dto = Assert.IsType<EditResultDto>(response.Data);
            Assert.True(dto.QuickValidation.Passed);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertField_afterControl_into_a_plaintext_cell_control_is_refused_and_leaves_the_file_untouched()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var before = File.ReadAllBytes(path);

            // YourReference_Lbl (id -1130623254) is a cell-level PLAIN-TEXT content control (its w:sdtPr
            // carries <w:text/>). afterControl anchors the new field inside that control's own cell, which
            // would nest a content control inside a plain-text one - well-formed OOXML that OpenXmlValidator
            // accepts but that Word rejects as "The file appears to be corrupted". This is the exact defect
            // behind the reported bug (a sales-invoice edit that reported success yet produced a file Word
            // could not open); the tool must now refuse it and leave the file byte-for-byte untouched.
            const int cellLevelPlainTextLabelId = -1130623254;

            var response = EditTools.InsertField(
                path, "/Header/SalesPersonName", "afterControl", controlId: cellLevelPlainTextLabelId);

            Assert.False(response.Ok);
            Assert.NotNull(response.Error);
            Assert.Equal("edit_would_corrupt", response.Error!.Code);
            Assert.Contains("plain-text", response.Error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));
            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertField_tableCell_targets_a_real_table_cell_in_the_corpus()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            // Table 1, row 4 is [ShippingAgentCode_Lbl (SdtCell) | PackageTrackingNo_Lbl (SdtCell) |
            // JobNo_Lbl (SdtCell) | a genuinely empty, bare w:tc]. Targeting the outer three columns would
            // nest a control inside an existing plain-text control (correctly rejected); col 3 is the plain
            // empty cell that can safely host a new field. Verify that shape holds (rather than just
            // assuming corpus internals) before relying on it: cells are counted the same way
            // get_layout_info / LocationResolver do - bare w:tc AND cell-level SdtCell wrappers.
            using (var probe = WordprocessingDocument.Open(path, false))
            {
                var body = probe.MainDocumentPart!.Document!.Body!;
                var tables = body.Descendants<Table>().ToList();
                Assert.True(tables.Count > 1, "test assumes at least two tables");
                var rows = tables[1].ChildElements.Where(e => e is TableRow or SdtRow).ToList();
                Assert.True(rows.Count > 4, "test assumes table 1 has at least 5 rows");
                var cells = ((TableRow)rows[4]).ChildElements.Where(e => e is TableCell or SdtCell).ToList();
                Assert.True(cells.Count > 3, "test assumes table 1 row 4 has at least four cells");
                Assert.IsType<TableCell>(cells[3]); // the target cell must be a bare (uncontrolled) w:tc.
            }

            var response = EditTools.InsertField(
                path, "/Header/DueDate", "tableCell", tableIndex: 1, row: 4, col: 3);

            Assert.True(response.Ok);
            Assert.Null(response.Error);
            var dto = Assert.IsType<EditResultDto>(response.Data);
            Assert.Equal("insert_field", dto.Operation);
            Assert.True(dto.QuickValidation.Passed);

            var reopened = ReadTools.GetLayoutInfo(path);
            var info = Assert.IsType<LayoutInfoDto>(reopened.Data);
            Assert.Contains(info.Controls, c => c.SdtId == dto.ControlId);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertField_atText_targets_real_existing_text_in_the_corpus()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            // atText needs ordinary static body text that is NOT inside a content control: nearly every BC
            // header field is a PLAIN-TEXT content control, and anchoring a new field inside one produces a
            // document Word rejects as corrupt (see PlainTextNestingGuard) - which the tool now correctly
            // refuses. This particular corpus layout's document.xml body is exceptionally densely bound -
            // verified directly against the file, every single non-whitespace w:t anywhere in the body sits
            // inside some w:sdt - so there is no PRE-EXISTING plain body text to discover here at all. Append
            // one real, uniquely-named static paragraph first (still a genuine round trip against the real
            // corpus file - only the anchor text itself is freshly added) rather than assume one exists.
            const string searchText = "MARKER-ATTEXT-STATIC-TEXT-UNIQUE";
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                var anchor = LocationResolver.Resolve(new Location { Type = LocationKind.DocumentEnd }, doc);
                anchor.InsertBlock(new Paragraph(new Run(new Text(searchText))));
                doc.MainDocumentPart!.Document!.Save();
            }

            var response = EditTools.InsertField(path, "/Header/CustomerAddress1", "atText", searchText: searchText);

            Assert.True(response.Ok);
            Assert.Null(response.Error);
            var dto = Assert.IsType<EditResultDto>(response.Data);
            Assert.Equal("insert_field", dto.Operation);
            Assert.True(dto.QuickValidation.Passed);

            var reopened = ReadTools.GetLayoutInfo(path);
            var info = Assert.IsType<LayoutInfoDto>(reopened.Data);
            Assert.Contains(info.Controls, c => c.SdtId == dto.ControlId);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- layoutPart / partName: header/footer targeting ----

    [Fact]
    public void InsertField_layoutPart_header_lands_in_a_header_part_and_reports_it_in_Part()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            // "Discover, don't assume" cannot apply here: the point of the assertion is WHICH header part a
            // partName-less insert picks, and the only discoverable orderings (the parts list, the package's
            // relationship order) are precisely the wrong answers - header2.xml comes first in both, while
            // header1.xml is this layout's section-default header. So the expectation is named outright.
            const string expectedPartName = "header1.xml";
            var before = Assert.IsType<LayoutInfoDto>(ReadTools.GetLayoutInfo(path).Data);
            Assert.Contains(expectedPartName, before.Parts);

            var response = EditTools.InsertField(
                path, "/Header/CustomerAddress1", "documentEnd", layoutPart: "header");

            Assert.True(response.Ok);
            Assert.Null(response.Error);
            var dto = Assert.IsType<EditResultDto>(response.Data);
            Assert.Equal(expectedPartName, dto.Part);
            Assert.True(dto.QuickValidation.Passed);
            Assert.Equal(0, dto.QuickValidation.ErrorCount);

            var reopened = Assert.IsType<LayoutInfoDto>(ReadTools.GetLayoutInfo(path).Data);
            Assert.Contains(reopened.Controls, c => c.SdtId == dto.ControlId && c.Part == expectedPartName && c.Kind == "Field");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertLabel_layoutPart_footer_with_explicit_partName_lands_in_that_exact_footer_part()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var before = Assert.IsType<LayoutInfoDto>(ReadTools.GetLayoutInfo(path).Data);
            var footerParts = before.Parts.Where(p => p.StartsWith("footer", StringComparison.OrdinalIgnoreCase)).ToList();
            Assert.NotEmpty(footerParts);
            var targetFooter = footerParts[0];

            var response = EditTools.InsertLabel(
                path, "/Header/Contact_Lbl", "documentEnd", layoutPart: "footer", partName: targetFooter);

            Assert.True(response.Ok);
            var dto = Assert.IsType<EditResultDto>(response.Data);
            Assert.Equal(targetFooter, dto.Part);
            Assert.True(dto.QuickValidation.Passed);

            var reopened = Assert.IsType<LayoutInfoDto>(ReadTools.GetLayoutInfo(path).Data);
            Assert.Contains(reopened.Controls, c => c.SdtId == dto.ControlId && c.Part == targetFooter && c.Kind == "Label");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertField_invalid_layoutPart_returns_invalid_argument_and_leaves_the_file_untouched()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var before = File.ReadAllBytes(path);

            var response = EditTools.InsertField(
                path, "/Header/CustomerAddress1", "documentEnd", layoutPart: "bogus-part");

            Assert.False(response.Ok);
            Assert.NotNull(response.Error);
            Assert.Equal("invalid_argument", response.Error!.Code);
            Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));
            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertField_afterControl_scoped_to_header_does_not_find_a_body_only_control()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            // YourReference_Lbl (id -1130623254) lives in document.xml - scoping the search to layoutPart
            //='header' must report it as not found there, not silently fall back to the body.
            const int bodyOnlyControlId = -1130623254;

            var response = EditTools.InsertField(
                path, "/Header/CustomerAddress1", "afterControl", controlId: bodyOnlyControlId, layoutPart: "header");

            Assert.False(response.Ok);
            Assert.NotNull(response.Error);
            Assert.Equal("not_found", response.Error!.Code);
            Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertRepeaterTable_layoutPart_header_is_rejected_as_invalid_argument_v1_scopes_repeaters_to_body_only()
    {
        // Follow-up review (post-Phase-4.3): unlike insert_field/insert_label, a repeater TABLE is v1-scoped
        // to the main body only - repeaters in headers/footers are explicitly deferred - GitHub
        // issue #10 (see LayoutEditor.InsertRepeaterTable's own chokepoint rejection). This inverts what was
        // originally a "lands in a header part" success test.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var before = File.ReadAllBytes(path);

            var response = TableTools.InsertRepeaterTable(
                path, "/Header/Line", "ItemNo_Line,Description_Line", "documentEnd", layoutPart: "header");

            Assert.False(response.Ok);
            Assert.NotNull(response.Error);
            Assert.Equal("invalid_argument", response.Error!.Code);
            Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));
            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- insert_repeater_table ----

    [Fact]
    public void InsertRepeaterTable_returns_ok_persists_to_disk_and_includes_a_passing_quickValidation()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.InsertRepeaterTable(
                path, "/Header/Line", "ItemNo_Line, Description_Line, TransHeaderAmount", "documentEnd");

            Assert.True(response.Ok);
            Assert.Null(response.Error);
            var dto = Assert.IsType<EditResultDto>(response.Data);
            Assert.Equal("insert_repeater_table", dto.Operation);
            Assert.Equal("Repeater", dto.Kind);
            Assert.Equal(3, dto.ColumnCount);
            Assert.Equal("document.xml", dto.Part);
            Assert.False(string.IsNullOrWhiteSpace(dto.Summary));
            Assert.Equal("quick", dto.QuickValidation.Level);
            Assert.True(dto.QuickValidation.Passed);
            Assert.Equal(0, dto.QuickValidation.ErrorCount);

            // Persisted to disk: reopen (a fresh tool call, not the same handle) and confirm it's really there.
            var reopened = ReadTools.GetLayoutInfo(path);
            var info = Assert.IsType<LayoutInfoDto>(reopened.Data);
            Assert.Contains(info.Controls, c => c.SdtId == dto.ControlId && c.Kind == "Repeater" && c.XPath == dto.XPath);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertRepeaterTable_columns_are_split_and_trimmed_from_a_comma_separated_string()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            // Deliberately irregular spacing around commas - the tool must split+trim internally.
            var response = TableTools.InsertRepeaterTable(
                path, "/Header/Line", "  ItemNo_Line ,Description_Line,  TransHeaderAmount  ", "documentEnd");

            Assert.True(response.Ok);
            var dto = Assert.IsType<EditResultDto>(response.Data);
            Assert.Equal(3, dto.ColumnCount);
            Assert.True(dto.QuickValidation.Passed);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertRepeaterTable_bad_dataItem_returns_invalid_argument()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.InsertRepeaterTable(
                path, "/Header/ThisDataItemDoesNotExistAnywhere", "ItemNo_Line", "documentEnd");

            Assert.False(response.Ok);
            Assert.NotNull(response.Error);
            Assert.Equal("invalid_argument", response.Error!.Code);
            Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertRepeaterTable_bad_column_returns_invalid_argument()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.InsertRepeaterTable(
                path, "/Header/Line", "ThisColumnDoesNotExistAnywhere", "documentEnd");

            Assert.False(response.Ok);
            Assert.NotNull(response.Error);
            Assert.Equal("invalid_argument", response.Error!.Code);
            Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertRepeaterTable_empty_columns_returns_invalid_argument()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.InsertRepeaterTable(path, "/Header/Line", "   ", "documentEnd");

            Assert.False(response.Ok);
            Assert.Equal("invalid_argument", response.Error!.Code);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertRepeaterTable_columnWidths_count_mismatch_returns_invalid_argument()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = TableTools.InsertRepeaterTable(
                path, "/Header/Line", "ItemNo_Line,Description_Line,TransHeaderAmount", "documentEnd",
                columnWidths: "1000,2000");

            Assert.False(response.Ok);
            Assert.Equal("invalid_argument", response.Error!.Code);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertRepeaterTable_missing_file_returns_file_not_found_with_hint()
    {
        var response = TableTools.InsertRepeaterTable(
            "Z:\\does-not-exist.docx", "/Header/Line", "ItemNo_Line", "documentEnd");

        Assert.False(response.Ok);
        Assert.NotNull(response.Error);
        Assert.Equal("file_not_found", response.Error!.Code);
        Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));
    }

    [Fact]
    public void RemoveControl_without_keepText_returns_ok_and_the_control_is_gone_on_reopen()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var inserted = EditTools.InsertField(path, "/Header/CustomerAddress1", "documentEnd");
            var insertedDto = Assert.IsType<EditResultDto>(inserted.Data);

            var response = EditTools.RemoveControl(path, insertedDto.ControlId);

            Assert.True(response.Ok);
            Assert.Null(response.Error);
            var dto = Assert.IsType<EditResultDto>(response.Data);
            Assert.Equal("remove_control", dto.Operation);
            Assert.Equal(insertedDto.ControlId, dto.ControlId);
            Assert.True(dto.QuickValidation.Passed);

            var reopened = ReadTools.GetLayoutInfo(path);
            var info = Assert.IsType<LayoutInfoDto>(reopened.Data);
            Assert.DoesNotContain(info.Controls, c => c.SdtId == insertedDto.ControlId);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveControl_on_a_cell_level_address_field_preserves_the_column_end_to_end()
    {
        // End-to-end regression for the reported bug: removing a cell-level address field via the tool must
        // leave the table row's cell count unchanged (the column survives) and still pass validation.
        const int customerAddress6Id = -2064325541; // #Nav: /Header/CustomerAddress6
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            int cellsBefore;
            using (var probe = WordprocessingDocument.Open(path, false))
            {
                var sdt = probe.MainDocumentPart!.Document!.Descendants<SdtElement>()
                    .Single(s => s.GetFirstChild<SdtProperties>()?.GetFirstChild<SdtId>()?.Val?.Value == customerAddress6Id);
                cellsBefore = sdt.Ancestors<TableRow>().First().ChildElements.Count(e => e is TableCell or SdtCell);
            }

            var response = EditTools.RemoveControl(path, customerAddress6Id);

            Assert.True(response.Ok);
            var dto = Assert.IsType<EditResultDto>(response.Data);
            Assert.True(dto.QuickValidation.Passed);
            Assert.Contains("column is preserved", dto.Summary);

            using var reopened = WordprocessingDocument.Open(path, false);
            Assert.DoesNotContain(
                reopened.MainDocumentPart!.Document!.Descendants<SdtElement>(),
                s => s.GetFirstChild<SdtProperties>()?.GetFirstChild<SdtId>()?.Val?.Value == customerAddress6Id);

            var companyAddress6 = reopened.MainDocumentPart!.Document!.Descendants<SdtElement>()
                .First(s => s.GetFirstChild<SdtProperties>()?.GetFirstChild<SdtAlias>()?.Val?.Value == "#Nav: /Header/CompanyAddress6");
            var cellsAfter = companyAddress6.Ancestors<TableRow>().First().ChildElements.Count(e => e is TableCell or SdtCell);
            Assert.Equal(cellsBefore, cellsAfter);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveControl_of_a_cells_sole_block_control_never_leaves_an_empty_cell_end_to_end()
    {
        // End-to-end regression for the reported corruption (removing fields that are a repeater-row
        // cell's only content): ContractBillingDetailsContractNoLbl and ContractBillingDetailsPositionNoLbl
        // are BLOCK-level controls that are the sole content of their repeater-row cells. Removing them via
        // the tool must leave every table cell with at least one paragraph — a w:tc left with only its
        // w:tcPr is a silently-corrupt document (Word rejects it; OpenXmlValidator, and therefore the
        // pre-save gate, does not).
        const int contractNoLblId = -1414548452; // #Nav: /Header/ContractBillingDetailsMapping/ContractBillingDetailsContractNoLbl
        const int positionNoLblId = 563840068;   // #Nav: /Header/ContractBillingDetailsMapping/ContractBillingDetailsPositionNoLbl
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            foreach (var id in new[] { contractNoLblId, positionNoLblId })
            {
                var response = EditTools.RemoveControl(path, id);
                Assert.True(response.Ok);
                var dto = Assert.IsType<EditResultDto>(response.Data);
                Assert.True(dto.QuickValidation.Passed);
                Assert.Contains("column is preserved", dto.Summary);
            }

            using var reopened = WordprocessingDocument.Open(path, false);
            Assert.DoesNotContain(
                reopened.MainDocumentPart!.Document!.Descendants<SdtElement>(),
                s => s.GetFirstChild<SdtProperties>()?.GetFirstChild<SdtId>()?.Val?.Value is contractNoLblId or positionNoLblId);

            // The actual corruption check: no table cell anywhere is left with only its w:tcPr and no
            // block-level content (a cell whose content is a child control legitimately has no DIRECT
            // paragraph, so this looks for any block-level child: paragraph, table, or content control).
            Assert.All(
                reopened.MainDocumentPart!.Document!.Descendants<TableCell>(),
                cell => Assert.Contains(
                    cell.ChildElements,
                    e => e is Paragraph or Table or SdtElement));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveControl_with_keepText_true_removes_the_wrapper_but_the_text_survives_on_disk()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var inserted = EditTools.InsertLabel(path, "/Header/Contact_Lbl", "documentEnd");
            var insertedDto = Assert.IsType<EditResultDto>(inserted.Data);

            var response = EditTools.RemoveControl(path, insertedDto.ControlId, keepText: true);

            Assert.True(response.Ok);
            var dto = Assert.IsType<EditResultDto>(response.Data);
            Assert.Equal("remove_control", dto.Operation);
            Assert.True(dto.QuickValidation.Passed);

            var reopened = ReadTools.GetLayoutInfo(path);
            var info = Assert.IsType<LayoutInfoDto>(reopened.Data);
            Assert.DoesNotContain(info.Controls, c => c.SdtId == insertedDto.ControlId);

            // Deeper than the tool surface exposes: open the saved file directly and confirm the label's
            // default placeholder text (its own leaf segment name) really did survive as bare content,
            // proving keepText did more than just make the control invisible to the inventory.
            using var doc = WordprocessingDocument.Open(path, false);
            var body = doc.MainDocumentPart!.Document!.Body!;
            Assert.Contains(body.Descendants<Text>(), t => t.Text == "Contact_Lbl");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveControl_keepText_true_on_a_repeater_returns_invalid_argument_and_changes_nothing()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var before = File.ReadAllBytes(path);
            var info = Assert.IsType<LayoutInfoDto>(ReadTools.GetLayoutInfo(path).Data);
            var repeater = info.Controls.First(c => c.Kind == "Repeater" && c.SdtId.HasValue);

            var response = EditTools.RemoveControl(path, repeater.SdtId!.Value, keepText: true);

            Assert.False(response.Ok);
            Assert.NotNull(response.Error);
            Assert.Equal("invalid_argument", response.Error!.Code);
            Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));

            var after = File.ReadAllBytes(path);
            Assert.Equal(before, after);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertField_invalid_locationType_returns_invalid_argument()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = EditTools.InsertField(path, "/Header/CustomerAddress1", "bogus-location");

            Assert.False(response.Ok);
            Assert.NotNull(response.Error);
            Assert.Equal("invalid_argument", response.Error!.Code);
            Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertField_afterControl_without_controlId_returns_invalid_argument()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = EditTools.InsertField(path, "/Header/CustomerAddress1", "afterControl");

            Assert.False(response.Ok);
            Assert.NotNull(response.Error);
            Assert.Equal("invalid_argument", response.Error!.Code);
            Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertLabel_bad_dataset_path_returns_invalid_argument_and_leaves_the_file_untouched()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var before = File.ReadAllBytes(path);

            // Not label-shaped -> SdtFactory.BuildLabel rejects it (ArgumentException) before any resolve.
            var response = EditTools.InsertLabel(path, "/Header/CustomerAddress1", "documentEnd");

            Assert.False(response.Ok);
            Assert.Equal("invalid_argument", response.Error!.Code);

            // insert_label is the one mutating tool the combined byte-identical test below does not
            // exercise a failure mode for - proven here in its own right instead.
            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    /// <summary>
    /// The row at <paramref name="rowIndex"/>, counting BOTH plain <c>w:tr</c> AND row-level <c>SdtRow</c>
    /// wrappers as rows — the same convention <c>TableStructureReader</c> (and therefore
    /// <c>get_layout_info</c>'s row indices) uses, so a row index sourced from that tool lines up with the
    /// physical row here even when an earlier row in the same table is a repeater's SdtRow.
    /// </summary>
    private static OpenXmlElement RowAt(Table table, int rowIndex) =>
        table.ChildElements.Where(e => e is TableRow or SdtRow).ElementAt(rowIndex);

    /// <summary>Finds the (tableIndex, row, col) of the first body cell whose visible text contains <paramref name="needle"/>.</summary>
    private static (int TableIndex, int Row, int Col) FindCell(string layoutPath, string needle)
    {
        var info = Assert.IsType<LayoutInfoDto>(ReadTools.GetLayoutInfo(layoutPath).Data);
        foreach (var table in info.Tables.Where(t => t.Part == "document.xml"))
        {
            foreach (var row in table.Rows)
            {
                foreach (var cell in row.Cells)
                {
                    if (cell.Text.Contains(needle, StringComparison.Ordinal))
                    {
                        return (table.TableIndex, row.RowIndex, cell.ColIndex);
                    }
                }
            }
        }

        throw new Xunit.Sdk.XunitException($"No cell containing '{needle}' was found.");
    }

    [Fact]
    public void ClearCellText_blanks_a_plain_text_header_label_and_preserves_the_column_end_to_end()
    {
        // End-to-end for the reported gap: the report's grand-total row label "Total" is PLAIN TEXT (not a
        // content control), so remove_control cannot touch it. clear_cell_text must blank it while leaving
        // the cell (and its column) intact and the document valid.
        var path = CopyOfCorpus(Corpus.InventoryOrderDetails);
        try
        {
            var (tableIndex, row, col) = FindCell(path, "Total");

            int cellsBefore;
            using (var probe = WordprocessingDocument.Open(path, false))
            {
                cellsBefore = RowAt(probe.MainDocumentPart!.Document!.Descendants<Table>().ElementAt(tableIndex), row)
                    .ChildElements.Count(e => e is TableCell or SdtCell);
            }

            var response = EditTools.ClearCellText(path, tableIndex, row, col);

            Assert.True(response.Ok);
            var dto = Assert.IsType<CellEditResultDto>(response.Data);
            Assert.True(dto.QuickValidation.Passed);
            Assert.Equal("Total", dto.PreviousText);
            Assert.Equal(string.Empty, dto.NewText);

            using var reopened = WordprocessingDocument.Open(path, false);
            var editedRow = RowAt(reopened.MainDocumentPart!.Document!.Descendants<Table>().ElementAt(tableIndex), row);
            // Scoped to the target cell specifically (col), not the whole row: the row's OTHER real cell in
            // this corpus - Totals_OutstandingAmt (col 9) - legitimately contains "Total" as a substring of
            // its own placeholder text, which a whole-row check would wrongly trip on.
            var editedCell = editedRow.ChildElements.Where(e => e is TableCell or SdtCell).ElementAt(col);
            Assert.DoesNotContain(editedCell.Descendants<Text>(), t => t.Text.Contains("Total", StringComparison.Ordinal));
            // Column preserved: same cell count, and the cleared cell still has a paragraph.
            Assert.Equal(cellsBefore, editedRow.ChildElements.Count(e => e is TableCell or SdtCell));
            Assert.All(editedRow.Elements<TableCell>(), c => Assert.NotEmpty(c.Elements<Paragraph>()));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SetCellText_relabels_a_plain_text_header_cell_end_to_end()
    {
        var path = CopyOfCorpus(Corpus.InventoryOrderDetails);
        try
        {
            var (tableIndex, row, col) = FindCell(path, "Total");

            var response = EditTools.SetCellText(path, tableIndex, row, col, "Grand Total");

            Assert.True(response.Ok);
            var dto = Assert.IsType<CellEditResultDto>(response.Data);
            Assert.True(dto.QuickValidation.Passed);
            Assert.Equal("Grand Total", dto.NewText);

            // Read back through the same coordinate model the tool used (get_layout_info's tables[]).
            var after = Assert.IsType<LayoutInfoDto>(ReadTools.GetLayoutInfo(path).Data);
            var editedCell = after.Tables.Single(t => t.TableIndex == tableIndex && t.Part == "document.xml")
                .Rows.Single(r => r.RowIndex == row)
                .Cells.Single(c => c.ColIndex == col);
            Assert.Equal("Grand Total", editedCell.Text);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void ClearCellText_on_a_cell_holding_a_control_returns_invalid_argument_and_leaves_the_file_untouched()
    {
        // The header address block: CustomerAddress fields are cell-level controls. Targeting one of those
        // cells with clear_cell_text must be refused (it is not a plain-text cell) and change nothing.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var (tableIndex, row, col) = FindCell(path, "CustomerAddress1");
            var before = File.ReadAllBytes(path);

            var response = EditTools.ClearCellText(path, tableIndex, row, col);

            Assert.False(response.Ok);
            Assert.Equal("invalid_argument", response.Error!.Code);
            Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));
            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SetCellText_out_of_range_cell_returns_not_found()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = EditTools.SetCellText(path, 0, 0, 999, "x");

            Assert.False(response.Ok);
            Assert.Equal("not_found", response.Error!.Code);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RemoveControl_of_a_nonexistent_id_returns_a_not_found_code()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var response = EditTools.RemoveControl(path, 999_999_999);

            Assert.False(response.Ok);
            Assert.NotNull(response.Error);
            Assert.Equal("not_found", response.Error!.Code);
            Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertField_missing_file_returns_file_not_found_with_hint()
    {
        var response = EditTools.InsertField("Z:\\does-not-exist.docx", "/Header/CustomerAddress1", "documentEnd");

        Assert.False(response.Ok);
        Assert.NotNull(response.Error);
        Assert.Equal("file_not_found", response.Error!.Code);
        Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));
    }

    [Fact]
    public void InsertLabel_missing_file_returns_file_not_found_with_hint()
    {
        var response = EditTools.InsertLabel("Z:\\does-not-exist.docx", "/Header/Contact_Lbl", "documentEnd");

        Assert.False(response.Ok);
        Assert.Equal("file_not_found", response.Error!.Code);
        Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));
    }

    [Fact]
    public void RemoveControl_missing_file_returns_file_not_found_with_hint()
    {
        var response = EditTools.RemoveControl("Z:\\does-not-exist.docx", 1);

        Assert.False(response.Ok);
        Assert.NotNull(response.Error);
        Assert.Equal("file_not_found", response.Error!.Code);
        Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));
    }

    [Fact]
    public void A_failed_edit_leaves_the_file_on_disk_byte_identical_to_before_the_call()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var before = File.ReadAllBytes(path);

            var invalidLocation = EditTools.InsertField(path, "/Header/CustomerAddress1", "bogus-location");
            Assert.False(invalidLocation.Ok);

            var missingControlId = EditTools.InsertField(path, "/Header/CustomerAddress1", "afterControl", controlId: 999_999_999);
            Assert.False(missingControlId.Ok);

            var unknownRemoveTarget = EditTools.RemoveControl(path, 999_999_999);
            Assert.False(unknownRemoveTarget.Ok);

            var badRepeaterTable = TableTools.InsertRepeaterTable(
                path, "/Header/ThisDataItemDoesNotExistAnywhere", "ItemNo_Line", "documentEnd");
            Assert.False(badRepeaterTable.Ok);

            var after = File.ReadAllBytes(path);
            Assert.Equal(before, after);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- refresh_xml_part ----
    // Every test below runs against its own temp COPY of a corpus file (never the shared corpus itself,
    // since this tool writes in place) and deletes it afterward.

    /// <summary>
    /// Finds every custom XML part whose root namespace starts with the BC prefix, via PUBLIC API only
    /// (mirrors <c>LayoutBuilderTests.FindBcCustomXmlParts</c>/<c>LayoutRefresherTests.FindBcCustomXmlParts</c> -
    /// <c>SchemaProvider.FindBcParts</c> is internal and not visible from this test assembly).
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

    /// <summary>
    /// Builds a modified schema .xml from <paramref name="corpusDocxPath"/>'s own raw BC-part XML: removes
    /// <paramref name="removedColumn"/> and adds a brand-new <paramref name="addedColumn"/>, both directly
    /// under <c>Header</c> - enough to exercise the orphaned-binding AND new-unbound-field reports through
    /// the tool surface in one call.
    /// </summary>
    private static string BuildModifiedSchemaXml(string corpusDocxPath, string removedColumn, string addedColumn)
    {
        XDocument xdoc;
        using (var doc = WordprocessingDocument.Open(corpusDocxPath, false))
        {
            var bcPart = FindBcCustomXmlParts(doc.MainDocumentPart!).Single();
            using var stream = bcPart.GetStream(FileMode.Open, FileAccess.Read);
            xdoc = XDocument.Load(stream);
        }

        var ns = xdoc.Root!.Name.Namespace;
        var header = xdoc.Root.Element(ns + "Header")!;
        header.Element(ns + removedColumn)!.Remove();
        header.Add(new XElement(ns + addedColumn, addedColumn));

        var path = Path.Combine(Path.GetTempPath(), $"bcwl-refreshtool-schema-{Guid.NewGuid():N}.xml");
        xdoc.Save(path);
        return path;
    }

    [Fact]
    public void RefreshXmlPart_with_its_own_schema_returns_ok_with_populated_RefreshResultDto()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var before = Assert.IsType<LayoutInfoDto>(ReadTools.GetLayoutInfo(path).Data);

            var response = LifecycleTools.RefreshXmlPart(path, Corpus.Path(Corpus.SalesInvoice));

            Assert.True(response.Ok);
            Assert.Null(response.Error);
            var dto = Assert.IsType<RefreshResultDto>(response.Data);

            Assert.Equal("Standard_Sales_Invoice", dto.OldReportName);
            Assert.Equal("Standard_Sales_Invoice", dto.NewReportName);
            Assert.Equal("1306", dto.OldReportId);
            Assert.Equal("1306", dto.NewReportId);
            Assert.False(dto.NamespaceChanged);
            Assert.Empty(dto.OrphanedBindings);
            Assert.True(dto.RemappedCount > 0);
            Assert.Equal(before.Report.StoreItemId, dto.StoreItemId, StringComparer.OrdinalIgnoreCase);
            Assert.Equal("quick", dto.QuickValidation.Level);
            Assert.True(dto.QuickValidation.Passed);
            Assert.Equal(0, dto.QuickValidation.ErrorCount);

            // Persisted to disk: reopen (a fresh tool call, not the same handle) and confirm it matches.
            var after = Assert.IsType<LayoutInfoDto>(ReadTools.GetLayoutInfo(path).Data);
            Assert.Equal(before.Report.StoreItemId, after.Report.StoreItemId);
            Assert.True(after.Validation.Passed);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RefreshXmlPart_reports_orphaned_binding_and_new_unbound_field_without_failing()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        string? schemaPath = null;
        try
        {
            schemaPath = BuildModifiedSchemaXml(
                Corpus.Path(Corpus.SalesInvoice), "SalesPersonName", "BrandNewFieldFromTool456");

            var response = LifecycleTools.RefreshXmlPart(path, schemaPath);

            // STRUCTURAL GATE ONLY: an orphaned binding is a semantic (quick-validation) finding, never a
            // reason to fail the call.
            Assert.True(response.Ok);
            Assert.Null(response.Error);
            var dto = Assert.IsType<RefreshResultDto>(response.Data);

            Assert.True(dto.OrphanedBindings.Count >= 1);
            Assert.Contains(dto.OrphanedBindings, o => o.XPath.EndsWith("SalesPersonName[1]", StringComparison.Ordinal));

            // OLD-vs-NEW diff: exactly the one genuinely-added field, not every other pre-existing unbound
            // leaf the corpus schema happens to carry (the removed column can never appear here either -
            // it no longer exists in the new schema at all).
            Assert.Equal(new[] { "/Header/BrandNewFieldFromTool456" }, dto.NewUnboundFields);

            // The orphan surfaces in quickValidation too (as data, not a tool failure).
            Assert.False(dto.QuickValidation.Passed);
            Assert.True(dto.QuickValidation.ErrorCount > 0);
        }
        finally
        {
            File.Delete(path);
            if (schemaPath is not null)
            {
                File.Delete(schemaPath);
            }
        }
    }

    [Fact]
    public void RefreshXmlPart_missing_layoutPath_returns_file_not_found_with_hint()
    {
        var response = LifecycleTools.RefreshXmlPart("Z:\\does-not-exist.docx", Corpus.Path(Corpus.SalesInvoice));

        Assert.False(response.Ok);
        Assert.NotNull(response.Error);
        Assert.Equal("file_not_found", response.Error!.Code);
        Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));
    }

    [Fact]
    public void RefreshXmlPart_missing_newSchemaSource_returns_file_not_found_and_leaves_the_file_untouched()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var before = File.ReadAllBytes(path);

            var response = LifecycleTools.RefreshXmlPart(path, "Z:\\does-not-exist.xml");

            Assert.False(response.Ok);
            Assert.NotNull(response.Error);
            Assert.Equal("file_not_found", response.Error!.Code);
            Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));

            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RefreshXmlPart_invalid_newSchemaSource_returns_invalid_layout_and_leaves_the_file_untouched()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        var badXmlPath = Path.Combine(Path.GetTempPath(), $"bcwl-refreshtool-badschema-{Guid.NewGuid():N}.xml");
        try
        {
            File.WriteAllText(badXmlPath, "<SomeOtherRoot xmlns=\"urn:not-bc\"><Foo/></SomeOtherRoot>");
            var before = File.ReadAllBytes(path);

            var response = LifecycleTools.RefreshXmlPart(path, badXmlPath);

            Assert.False(response.Ok);
            Assert.NotNull(response.Error);
            Assert.Equal("invalid_layout", response.Error!.Code);
            Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));

            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally
        {
            File.Delete(path);
            if (File.Exists(badXmlPath))
            {
                File.Delete(badXmlPath);
            }
        }
    }

    [Fact]
    public void RefreshXmlPart_newSchemaSource_with_correct_root_name_but_non_BC_namespace_returns_invalid_layout_and_leaves_the_file_untouched()
    {
        // Distinct from the test above: right root LOCAL NAME (NavWordReportXmlPart), WRONG namespace (not
        // "urn:microsoft-dynamics-nav/reports/..."). Without SchemaProvider.FromSchemaXml's namespace check,
        // this would have been silently ACCEPTED as a valid schema (root name matches) and would have
        // orphaned every existing binding on refresh instead of being rejected up front.
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        var badXmlPath = Path.Combine(Path.GetTempPath(), $"bcwl-refreshtool-badns-{Guid.NewGuid():N}.xml");
        try
        {
            File.WriteAllText(
                badXmlPath,
                "<NavWordReportXmlPart xmlns=\"urn:not-bc\"><Header><Foo>x</Foo></Header></NavWordReportXmlPart>");
            var before = File.ReadAllBytes(path);

            var response = LifecycleTools.RefreshXmlPart(path, badXmlPath);

            Assert.False(response.Ok);
            Assert.NotNull(response.Error);
            Assert.Equal("invalid_layout", response.Error!.Code);
            Assert.False(string.IsNullOrWhiteSpace(response.Error.Hint));

            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally
        {
            File.Delete(path);
            if (File.Exists(badXmlPath))
            {
                File.Delete(badXmlPath);
            }
        }
    }
}
