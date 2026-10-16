using BcWordLayout.Domain;
using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;
using BcWordLayout.Merge;
using BcWordLayout.Render;

namespace BcWordLayout.Tests;

/// <summary>
/// GitHub issue #94: a layout that is open in Word or carries the read-only attribute (a TFVC/Perforce
/// workspace file that is not checked out). Four guarantees, each pinned below:
/// <list type="bullet">
/// <item>A write that would REPLACE such a file fails as <c>file_locked</c> with a hint naming the three
/// causes (open elsewhere, read-only, no write permission), never as <c>internal_error</c>, and the file is
/// left byte-identical. Word's own hold is reproduced as a second handle with read/write access that shares
/// read only: copies and reads of the file still succeed under it, so the denial shows up only at the
/// commit's rename (<c>CrossProcessLockTests</c>' <c>FileShare.None</c> handle fails much earlier, at the
/// copy, and never reached that rename).</item>
/// <item>Operations that never write the source (an edit's working copy, <c>preview_layout</c>'s merge
/// stage, <c>validate_layout level=full</c>'s dry run, a <c>create_layout</c> template) work on a read-only
/// source exactly as on a writable one, and leave no read-only private copy behind.</item>
/// <item>Every private-copy family is swept once orphaned, read-only or not: <c>.bcwl-build-*</c> next to a
/// <c>create_layout</c> output and <c>%TEMP%\bcwl-edit-*</c> were never swept before.</item>
/// <item><see cref="ScratchFiles"/>' own mechanics: a writable copy of a read-only source, and a cleanup
/// delete that never throws.</item>
/// </list>
/// Every test works in its own private directory, never on the shared corpus or a shared folder. The
/// read-only and held-handle tests are Windows-only: on POSIX a rename replaces a read-only or open file
/// happily, so there is nothing to refuse.
/// </summary>
/// <remarks>Joins the preview-converter-seam collection because some tests call
/// <c>LifecycleTools.PreviewLayout</c> with a <see cref="FakePdfConverter"/> (see
/// <see cref="PreviewConverterSeamCollection"/>); no test here ever reaches a real converter.</remarks>
[Collection("preview-converter-seam")]
public class ReadOnlyAndHeldFileTests
{
    private static string PrivateDirectory() => Directory.CreateTempSubdirectory("bcwl-readonly-").FullName;

    private static string CopyOfCorpusInto(string dir, string corpusFile, string fileName = "layout.docx")
    {
        var path = Path.Combine(dir, fileName);
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    private static void SetReadOnly(string path) =>
        File.SetAttributes(path, File.GetAttributes(path) | FileAttributes.ReadOnly);

    private static bool IsReadOnly(string path) => (File.GetAttributes(path) & FileAttributes.ReadOnly) != 0;

    /// <summary>Opens <paramref name="path"/> the way Word holds an open document: read/write, sharing read only.</summary>
    private static FileStream HoldLikeWord(string path) =>
        new(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);

    /// <summary>Plants an orphan of a private-copy family, backdated past the sweep's retention window.</summary>
    private static string PlantOrphan(string dir, string fileName, bool stale, bool readOnly)
    {
        var path = Path.Combine(dir, fileName);
        File.WriteAllText(path, "orphaned private copy from a hard-killed process");
        if (stale)
        {
            // CreationTimeUtc is the sweep's age signal - see ScratchFiles.OrphanRetentionWindow.
            File.SetCreationTimeUtc(path, DateTime.UtcNow - TimeSpan.FromDays(2));
        }

        if (readOnly)
        {
            SetReadOnly(path);
        }

        return path;
    }

    private static IEnumerable<string> PrivateCopiesIn(string dir) =>
        Directory.EnumerateFiles(dir, ".bcwl-*", SearchOption.TopDirectoryOnly);

    /// <summary>Deletes <paramref name="dir"/> even when a test left read-only files in it.</summary>
    private static void DeleteDirectory(string dir)
    {
        if (!Directory.Exists(dir))
        {
            return;
        }

        foreach (var file in Directory.EnumerateFiles(dir, "*", SearchOption.AllDirectories))
        {
            File.SetAttributes(file, FileAttributes.Normal);
        }

        Directory.Delete(dir, recursive: true);
    }

    private static void AssertFileLockedNamingTheThreeCauses(ToolResponse response, string targetPath)
    {
        Assert.False(response.Ok);
        Assert.Null(response.Data);
        var error = response.Error!;
        Assert.Equal("file_locked", error.Code);
        Assert.Contains(Path.GetFullPath(targetPath), error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("left unchanged", error.Message, StringComparison.Ordinal);

        Assert.Contains(Path.GetFullPath(targetPath), error.Hint, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("open in Word or another program", error.Hint, StringComparison.Ordinal);
        Assert.Contains("read-only", error.Hint, StringComparison.Ordinal);
        Assert.Contains("not checked out", error.Hint, StringComparison.Ordinal);
        Assert.Contains("no write permission", error.Hint, StringComparison.Ordinal);
        Assert.DoesNotContain("bug in the tool", error.Hint, StringComparison.Ordinal);
    }

    // ---- the commit: a target held open the way Word holds it, or read-only -> file_locked ----

    [WindowsOnlyFact]
    public void InsertField_on_a_layout_held_open_the_way_Word_holds_it_returns_file_locked_and_leaves_it_byte_identical()
    {
        var dir = PrivateDirectory();
        try
        {
            var path = CopyOfCorpusInto(dir, Corpus.SalesInvoice);
            var before = File.ReadAllBytes(path);

            ToolResponse response;
            using (HoldLikeWord(path))
            {
                response = EditTools.InsertField(path, "/Header/CustomerAddress1", "documentEnd");
            }

            AssertFileLockedNamingTheThreeCauses(response, path);
            Assert.Equal(before, File.ReadAllBytes(path));
            Assert.Empty(PrivateCopiesIn(dir));

            // Once the handle is closed the identical call succeeds: the refusal left nothing behind.
            var followUp = EditTools.InsertField(path, "/Header/CustomerAddress1", "documentEnd");
            Assert.True(followUp.Ok, followUp.Error?.Message);
        }
        finally
        {
            DeleteDirectory(dir);
        }
    }

    [WindowsOnlyFact]
    public void InsertText_on_a_read_only_layout_returns_file_locked_keeps_it_read_only_and_leaks_no_copy()
    {
        var dir = PrivateDirectory();
        try
        {
            var path = CopyOfCorpusInto(dir, Corpus.StandardSalesQuote);
            var before = File.ReadAllBytes(path);
            SetReadOnly(path);
            var startedUtc = DateTime.UtcNow - TimeSpan.FromSeconds(2);

            var response = EditTools.InsertText(path, "Read-only check", "documentEnd");

            AssertFileLockedNamingTheThreeCauses(response, path);
            Assert.Equal(before, File.ReadAllBytes(path));
            Assert.True(IsReadOnly(path), "the caller's own file must keep its read-only attribute");
            Assert.Empty(PrivateCopiesIn(dir));

            // The %TEMP% working copy is shared with concurrently running tests, so only a READ-ONLY copy
            // created during this test can be this call's leak: no other test makes its source read-only.
            var leakedReadOnlyCopies = Directory.EnumerateFiles(Path.GetTempPath(), "bcwl-edit-*.docx")
                .Where(f =>
                {
                    try
                    {
                        return File.GetCreationTimeUtc(f) >= startedUtc && IsReadOnly(f);
                    }
                    catch (IOException)
                    {
                        return false; // deleted by its own call meanwhile
                    }
                })
                .ToList();
            Assert.Empty(leakedReadOnlyCopies);
        }
        finally
        {
            DeleteDirectory(dir);
        }
    }

    [WindowsOnlyFact]
    public void CreateLayout_over_a_read_only_outputPath_returns_file_locked_and_leaves_it_byte_identical()
    {
        var dir = PrivateDirectory();
        try
        {
            var outputPath = CopyOfCorpusInto(dir, Corpus.StandardStatement, "existing.docx");
            var before = File.ReadAllBytes(outputPath);
            SetReadOnly(outputPath);

            var response = LifecycleTools.CreateLayout(Corpus.Path(Corpus.SalesInvoice), outputPath);

            AssertFileLockedNamingTheThreeCauses(response, outputPath);
            Assert.Equal(before, File.ReadAllBytes(outputPath));
            Assert.True(IsReadOnly(outputPath));
            Assert.Empty(PrivateCopiesIn(dir));
        }
        finally
        {
            DeleteDirectory(dir);
        }
    }

    [WindowsOnlyFact]
    public void CreateLayout_over_an_outputPath_held_open_the_way_Word_holds_it_returns_file_locked()
    {
        var dir = PrivateDirectory();
        try
        {
            var outputPath = CopyOfCorpusInto(dir, Corpus.StandardStatement, "existing.docx");
            var before = File.ReadAllBytes(outputPath);

            ToolResponse response;
            using (HoldLikeWord(outputPath))
            {
                response = LifecycleTools.CreateLayout(Corpus.Path(Corpus.SalesInvoice), outputPath);
            }

            AssertFileLockedNamingTheThreeCauses(response, outputPath);
            Assert.Equal(before, File.ReadAllBytes(outputPath));
            Assert.Empty(PrivateCopiesIn(dir));
        }
        finally
        {
            DeleteDirectory(dir);
        }
    }

    [WindowsOnlyFact]
    public void PreviewLayout_when_the_previous_merged_docx_is_held_open_the_way_Word_holds_it_returns_file_locked()
    {
        // The realistic case: the developer opened the last preview's merged.docx in Word, then previewed again.
        var dir = PrivateDirectory();
        var outputDir = Path.Combine(dir, "preview");
        var fake = new FakePdfConverter { Name = "fake", IsAvailable = true, ConversionSucceeds = true };
        LifecycleTools.SelectConverter = _ => fake;
        try
        {
            var layoutPath = CopyOfCorpusInto(dir, Corpus.SalesInvoice);
            var first = LifecycleTools.PreviewLayout(layoutPath, outputDir: outputDir);
            Assert.True(first.Ok, first.Error?.Message);
            var mergedPath = Assert.IsType<PreviewResultDto>(first.Data).MergedDocxPath;
            var before = File.ReadAllBytes(mergedPath);

            ToolResponse second;
            using (HoldLikeWord(mergedPath))
            {
                second = LifecycleTools.PreviewLayout(layoutPath, outputDir: outputDir);
            }

            AssertFileLockedNamingTheThreeCauses(second, mergedPath);
            Assert.Equal(before, File.ReadAllBytes(mergedPath));
            Assert.Empty(PrivateCopiesIn(outputDir));
        }
        finally
        {
            LifecycleTools.SelectConverter = PdfConverterFactory.Select;
            DeleteDirectory(dir);
        }
    }

    [Fact]
    public void Guard_maps_only_the_target_scoped_denial_to_file_locked_never_a_bare_access_denied()
    {
        // The same UnauthorizedAccessException type is thrown when a SOURCE cannot be read, where "close it in
        // Word" would be the wrong advice - so only a throw site that knows it was writing the target maps it.
        var bare = ToolGuards.Guard(() => throw new UnauthorizedAccessException("Access to the path is denied."), "x.docx");
        Assert.Equal("internal_error", bare.Error!.Code);

        var target = Path.GetFullPath("target.docx");
        var scoped = ToolGuards.Guard(
            () => throw new TargetFileNotWritableException(target, new UnauthorizedAccessException("Access to the path is denied.")),
            "x.docx");
        AssertFileLockedNamingTheThreeCauses(scoped, target);
    }

    // ---- operations that never write the source work on a read-only one ----

    [WindowsOnlyFact]
    public void PreviewLayout_of_a_read_only_layout_merges_normally_and_leaves_no_staging_copy()
    {
        var dir = PrivateDirectory();
        var outputDir = Path.Combine(dir, "preview");
        var fake = new FakePdfConverter { Name = "fake", IsAvailable = true, ConversionSucceeds = true };
        LifecycleTools.SelectConverter = _ => fake;
        try
        {
            var layoutPath = CopyOfCorpusInto(dir, Corpus.SalesInvoice);
            var before = File.ReadAllBytes(layoutPath);
            SetReadOnly(layoutPath);

            var response = LifecycleTools.PreviewLayout(layoutPath, outputDir: outputDir);

            Assert.True(response.Ok, response.Error?.Message);
            var dto = Assert.IsType<PreviewResultDto>(response.Data);
            Assert.True(dto.Stats.FieldsFilled > 0);
            Assert.Equal(1, fake.ConvertCallCount);
            Assert.False(IsReadOnly(dto.MergedDocxPath), "the merged output is the tool's own file, never read-only");
            Assert.Empty(PrivateCopiesIn(outputDir));
            Assert.Equal(before, File.ReadAllBytes(layoutPath));
            Assert.True(IsReadOnly(layoutPath));
        }
        finally
        {
            LifecycleTools.SelectConverter = PdfConverterFactory.Select;
            DeleteDirectory(dir);
        }
    }

    [WindowsOnlyFact]
    public void ValidateLayout_full_on_a_read_only_layout_reports_exactly_what_the_writable_copy_does()
    {
        var dir = PrivateDirectory();
        try
        {
            var writable = CopyOfCorpusInto(dir, Corpus.SalesInvoice, "writable.docx");
            var readOnly = CopyOfCorpusInto(dir, Corpus.SalesInvoice, "readonly.docx");
            SetReadOnly(readOnly);

            var expected = Assert.IsType<ValidationResultDto>(ReadTools.ValidateLayout(writable, "full").Data);
            var response = ReadTools.ValidateLayout(readOnly, "full");

            Assert.True(response.Ok, response.Error?.Message);
            var actual = Assert.IsType<ValidationResultDto>(response.Data);
            Assert.DoesNotContain(actual.Findings, f => f.Check == "dry-run-merge");
            Assert.Equal(expected.Passed, actual.Passed);
            Assert.Equal(expected.ErrorCount, actual.ErrorCount);
            Assert.Equal(expected.WarningCount, actual.WarningCount);
            Assert.Equal(expected.Findings.Select(f => f.Check), actual.Findings.Select(f => f.Check));
            Assert.Empty(PrivateCopiesIn(dir));
        }
        finally
        {
            DeleteDirectory(dir);
        }
    }

    [WindowsOnlyFact]
    public void CreateLayout_with_a_read_only_templatePath_succeeds_and_writes_a_writable_output()
    {
        var dir = PrivateDirectory();
        try
        {
            // An unbound shell: a blank build carries a BC part but zero bound controls of its own.
            var templatePath = Path.Combine(dir, "letterhead.docx");
            LayoutBuilder.Create(Corpus.Path(Corpus.StandardStatement), templatePath);
            var templateBefore = File.ReadAllBytes(templatePath);
            SetReadOnly(templatePath);
            var outputPath = Path.Combine(dir, "created.docx");

            var response = LifecycleTools.CreateLayout(Corpus.Path(Corpus.SalesInvoice), outputPath, templatePath);

            Assert.True(response.Ok, response.Error?.Message);
            var dto = Assert.IsType<CreateResultDto>(response.Data);
            Assert.True(dto.UsedTemplate);
            Assert.Equal("Standard_Sales_Invoice", dto.ReportName);
            Assert.False(IsReadOnly(outputPath), "a read-only template must not make the new layout read-only");
            Assert.Equal(templateBefore, File.ReadAllBytes(templatePath));
            Assert.True(IsReadOnly(templatePath));
            Assert.Empty(PrivateCopiesIn(dir));
        }
        finally
        {
            DeleteDirectory(dir);
        }
    }

    // ---- every private-copy family is swept once orphaned, read-only or not ----

    [WindowsOnlyFact]
    public void InsertField_sweeps_stale_read_only_orphans_of_every_family_from_the_layouts_directory()
    {
        var dir = PrivateDirectory();
        try
        {
            var path = CopyOfCorpusInto(dir, Corpus.SalesInvoice);
            var staleStage = PlantOrphan(dir, $".bcwl-stage-{Guid.NewGuid():N}.docx", stale: true, readOnly: true);
            var staleMerge = PlantOrphan(dir, $".bcwl-merge-stage-{Guid.NewGuid():N}.docx", stale: true, readOnly: true);
            var staleBuild = PlantOrphan(dir, $".bcwl-build-{Guid.NewGuid():N}.docx", stale: true, readOnly: true);
            var freshBuild = PlantOrphan(dir, $".bcwl-build-{Guid.NewGuid():N}.docx", stale: false, readOnly: true);

            var response = EditTools.InsertField(path, "/Header/CustomerAddress1", "documentEnd");

            Assert.True(response.Ok, response.Error?.Message);
            Assert.False(File.Exists(staleStage), "a stale read-only .bcwl-stage-* orphan must be swept");
            Assert.False(File.Exists(staleMerge), "a stale read-only .bcwl-merge-stage-* orphan must be swept");
            Assert.False(File.Exists(staleBuild), "a stale read-only .bcwl-build-* orphan must be swept");
            Assert.True(File.Exists(freshBuild), "a fresh .bcwl-build-* file may be another call's live build");
        }
        finally
        {
            DeleteDirectory(dir);
        }
    }

    [Fact]
    public void CreateLayout_sweeps_a_stale_build_file_next_to_outputPath_and_keeps_a_fresh_one()
    {
        var dir = PrivateDirectory();
        try
        {
            var staleBuild = PlantOrphan(dir, $".bcwl-build-{Guid.NewGuid():N}.docx", stale: true, readOnly: true);
            var freshBuild = PlantOrphan(dir, $".bcwl-build-{Guid.NewGuid():N}.docx", stale: false, readOnly: false);

            var response = LifecycleTools.CreateLayout(Corpus.Path(Corpus.SalesInvoice), Path.Combine(dir, "created.docx"));

            Assert.True(response.Ok, response.Error?.Message);
            Assert.False(File.Exists(staleBuild), "a stale .bcwl-build-* orphan next to outputPath must be swept");
            Assert.True(File.Exists(freshBuild), "a fresh .bcwl-build-* file may be another call's live build");
        }
        finally
        {
            DeleteDirectory(dir);
        }
    }

    [Fact]
    public void Merge_sweeps_a_stale_read_only_merge_stage_file_from_its_output_directory()
    {
        var dir = PrivateDirectory();
        try
        {
            var staleMerge = PlantOrphan(dir, $".bcwl-merge-stage-{Guid.NewGuid():N}.docx", stale: true, readOnly: true);

            MergeEngine.Merge(Corpus.Path(Corpus.SalesInvoice), Path.Combine(dir, "merged.docx"));

            Assert.False(File.Exists(staleMerge), "a stale read-only .bcwl-merge-stage-* orphan must be swept");
        }
        finally
        {
            DeleteDirectory(dir);
        }
    }

    [Fact]
    public void SweepStaleEditWorkingCopies_removes_only_stale_files_of_the_exact_working_copy_shape()
    {
        // Driven against a private folder standing in for %TEMP% - which is shared with every other program,
        // so a candidate must match the exact generated name, not just the glob.
        var dir = PrivateDirectory();
        try
        {
            var staleOurs = PlantOrphan(dir, $"bcwl-edit-{Guid.NewGuid():N}.docx", stale: true, readOnly: true);
            var freshOurs = PlantOrphan(dir, $"bcwl-edit-{Guid.NewGuid():N}.docx", stale: false, readOnly: false);
            var staleOtherShape = PlantOrphan(dir, "bcwl-edit-someone-elses-file.docx", stale: true, readOnly: false);
            var staleUpperHex = PlantOrphan(
                dir, $"bcwl-edit-{Guid.NewGuid().ToString("N").ToUpperInvariant()}.docx", stale: true, readOnly: false);

            ToolGuards.SweepStaleEditWorkingCopies(dir);

            Assert.False(File.Exists(staleOurs), "a stale read-only working copy must be swept");
            Assert.True(File.Exists(freshOurs), "a fresh working copy may be another call's live edit");
            Assert.True(File.Exists(staleOtherShape), "a file that merely matches the glob is not ours to delete");
            Assert.True(File.Exists(staleUpperHex), "the generated GUID is lowercase; anything else is not ours");
        }
        finally
        {
            DeleteDirectory(dir);
        }
    }

    // ---- ScratchFiles' own mechanics ----

    [WindowsOnlyFact]
    public void ScratchFiles_CopyWritable_of_a_read_only_source_yields_a_writable_copy_with_the_same_bytes()
    {
        var dir = PrivateDirectory();
        try
        {
            var source = CopyOfCorpusInto(dir, Corpus.SalesInvoice, "source.docx");
            SetReadOnly(source);
            var copy = Path.Combine(dir, "copy.docx");

            ScratchFiles.CopyWritable(source, copy);

            Assert.False(IsReadOnly(copy));
            Assert.True(IsReadOnly(source));
            Assert.Equal(File.ReadAllBytes(source), File.ReadAllBytes(copy));
        }
        finally
        {
            DeleteDirectory(dir);
        }
    }

    [WindowsOnlyFact]
    public void ScratchFiles_Delete_removes_a_read_only_copy_and_never_throws_for_one_held_open()
    {
        var dir = PrivateDirectory();
        try
        {
            var readOnly = PlantOrphan(dir, "read-only.docx", stale: false, readOnly: true);
            ScratchFiles.Delete(readOnly);
            Assert.False(File.Exists(readOnly));

            var held = PlantOrphan(dir, "held.docx", stale: false, readOnly: false);
            using (new FileStream(held, FileMode.Open, FileAccess.Read, FileShare.None))
            {
                // A cleanup runs in a finally block: throwing here would replace the call's real outcome.
                ScratchFiles.Delete(held);
            }

            Assert.True(File.Exists(held), "the held copy survives, for a later sweep to remove");
            ScratchFiles.Delete(Path.Combine(dir, "never-existed.docx"));
        }
        finally
        {
            DeleteDirectory(dir);
        }
    }
}
