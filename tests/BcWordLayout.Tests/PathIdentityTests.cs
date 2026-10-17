using BcWordLayout.Domain;
using BcWordLayout.McpHost.Tools;
using BcWordLayout.Merge;
using DocumentFormat.OpenXml.Packaging;

namespace BcWordLayout.Tests;

/// <summary>
/// GitHub issue #92: the per-path locks keyed on <see cref="Path.GetFullPath(string)"/>, which leaves a junction,
/// a symlink and a subst or mapped drive as distinct spellings of one file, so two hosts editing one layout
/// through two spellings never serialized and one host's commit replaced the other's.
/// Everything that asks "is this the same file" now goes through <see cref="PathIdentity"/>. These tests pin the
/// resolution itself, the lock keys built on it, the commit target, and the merge's same-file refusal; the
/// cross-spelling lock contention is in <c>CrossProcessLockTests</c> (it shares that class's timeout seam), and
/// <c>create_layout</c>'s same-file refusal is in <c>CreateLayoutAliasingTests</c>. Subst and mapped drives are
/// not arranged here (both change machine-wide state); they resolve through the same API call.
/// </summary>
public class PathIdentityTests
{
    private static string PrivateDirectory() => Directory.CreateTempSubdirectory("bcwl-pathidentity-").FullName;

    private static string CopyOfCorpusInto(string dir, string corpusFile, string fileName)
    {
        var path = Path.Combine(dir, fileName);
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    private static string BodyText(string layoutPath)
    {
        using var doc = WordprocessingDocument.Open(layoutPath, false);
        return doc.MainDocumentPart!.Document!.Body!.InnerText;
    }

    private static void AssertSameSpelling(string expected, string actual) =>
        Assert.Equal(expected, actual, ignoreCase: true);

    // ---- the resolution itself ----

    [Fact]
    public void Canonical_agrees_with_GetFullPath_for_relative_and_differently_cased_spellings()
    {
        var dir = PrivateDirectory();
        try
        {
            var layout = CopyOfCorpusInto(dir, Corpus.SalesInvoice, "layout.docx");
            var relative = Path.GetRelativePath(Environment.CurrentDirectory, layout);

            AssertSameSpelling(PathIdentity.Canonical(layout), PathIdentity.Canonical(relative));
            AssertSameSpelling(PathIdentity.Canonical(layout), PathIdentity.Canonical(Path.Combine(dir, "sub", "..", "layout.docx")));
            Assert.True(PathIdentity.SameFile(layout, relative));
            Assert.False(PathIdentity.SameFile(layout, Path.Combine(dir, "other.docx")));
            if (OperatingSystem.IsWindows())
            {
                Assert.True(PathIdentity.SameFile(layout, layout.ToUpperInvariant()));
            }
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void Canonical_keeps_the_segments_that_do_not_exist_yet()
    {
        var dir = PrivateDirectory();
        try
        {
            var notYet = Path.Combine(dir, "new-folder", "deeper", "layout.docx");

            var canonical = PathIdentity.Canonical(notYet);

            Assert.EndsWith(Path.Combine("new-folder", "deeper", "layout.docx"), canonical, StringComparison.Ordinal);
            AssertSameSpelling(Path.Combine(PathIdentity.Canonical(dir), "new-folder", "deeper", "layout.docx"), canonical);
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void Canonical_throws_what_GetFullPath_throws_for_an_empty_path()
    {
        // Callers rely on GetFullPath's own argument exception (e.g. create_layout's empty-outputPath check runs
        // before the lock is keyed), so resolution must not change which exception a malformed path produces.
        Assert.Throws<ArgumentException>(() => PathIdentity.Canonical(string.Empty));
    }

    [WindowsOnlyFact]
    public void Canonical_resolves_a_junction_anywhere_along_the_path()
    {
        var dir = PrivateDirectory();
        var real = Directory.CreateDirectory(Path.Combine(dir, "real")).FullName;
        var link = Path.Combine(dir, "link");
        try
        {
            var layout = CopyOfCorpusInto(real, Corpus.SalesInvoice, "J.docx");
            FileSystemLinks.CreateJunction(link, real);
            var throughLink = Path.Combine(link, "J.docx");

            AssertSameSpelling(PathIdentity.Canonical(layout), PathIdentity.Canonical(throughLink));
            Assert.EndsWith(Path.Combine("real", "J.docx"), PathIdentity.Canonical(throughLink), StringComparison.OrdinalIgnoreCase);
            Assert.True(PathIdentity.SameFile(layout, throughLink));

            // A file that does not exist yet resolves through its nearest existing folder - the junction.
            AssertSameSpelling(
                PathIdentity.Canonical(Path.Combine(real, "new", "out.docx")),
                PathIdentity.Canonical(Path.Combine(link, "new", "out.docx")));
        }
        finally
        {
            FileSystemLinks.DeleteJunction(link);
            Directory.Delete(dir, recursive: true);
        }
    }

    [WindowsOnlyFact]
    public void Canonical_expands_an_8_3_short_name()
    {
        // GetFullPath already expanded these; this pins that the resolution keeps doing so. The file name is
        // longer than 8.3 allows, so the volume gives it a short alias when it generates them at all; on a
        // volume without 8.3 names there is no second spelling to compare.
        var dir = PrivateDirectory();
        try
        {
            var layout = CopyOfCorpusInto(dir, Corpus.SalesInvoice, "layout-with-a-long-name.docx");
            var shortSpelling = FileSystemLinks.ShortPathOrNull(layout);
            if (shortSpelling is null)
            {
                return;
            }

            AssertSameSpelling(PathIdentity.Canonical(layout), PathIdentity.Canonical(shortSpelling));
            Assert.True(PathIdentity.SameFile(layout, shortSpelling));
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [WindowsOnlyFact]
    public void Canonical_resolves_a_file_another_program_holds_without_sharing()
    {
        // Word holds an open layout without write sharing; the resolution opens it with no access rights, so a
        // held file is still resolved rather than falling back to its unresolved spelling.
        var dir = PrivateDirectory();
        var real = Directory.CreateDirectory(Path.Combine(dir, "real")).FullName;
        var link = Path.Combine(dir, "link");
        try
        {
            var layout = CopyOfCorpusInto(real, Corpus.SalesInvoice, "held.docx");
            FileSystemLinks.CreateJunction(link, real);

            using (new FileStream(layout, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                AssertSameSpelling(PathIdentity.Canonical(layout), PathIdentity.Canonical(Path.Combine(link, "held.docx")));
            }
        }
        finally
        {
            FileSystemLinks.DeleteJunction(link);
            Directory.Delete(dir, recursive: true);
        }
    }

    // ---- what is keyed on it ----

    [WindowsOnlyFact]
    public void Both_lock_keys_and_the_preview_hash_are_shared_by_every_spelling_of_a_layout()
    {
        var dir = PrivateDirectory();
        var real = Directory.CreateDirectory(Path.Combine(dir, "real-folder-with-a-long-name")).FullName;
        var link = Path.Combine(dir, "link");
        try
        {
            var layout = CopyOfCorpusInto(real, Corpus.SalesInvoice, "J.docx");
            FileSystemLinks.CreateJunction(link, real);
            var spellings = new List<string> { Path.Combine(link, "J.docx") };
            if (FileSystemLinks.ShortPathOrNull(layout) is { } shortSpelling)
            {
                spellings.Add(shortSpelling);
            }

            foreach (var spelling in spellings)
            {
                Assert.Same(ToolGuards.EditLockFor(layout), ToolGuards.EditLockFor(spelling));
                Assert.Equal(CrossProcessLock.MutexName(layout), CrossProcessLock.MutexName(spelling));
                Assert.Equal(LifecycleTools.PreviewPathHash(layout), LifecycleTools.PreviewPathHash(spelling));
            }
        }
        finally
        {
            FileSystemLinks.DeleteJunction(link);
            Directory.Delete(dir, recursive: true);
        }
    }

    [WindowsOnlyFact]
    public void An_edit_through_a_junction_commits_to_the_real_file_and_stages_beside_it()
    {
        var dir = PrivateDirectory();
        var real = Directory.CreateDirectory(Path.Combine(dir, "real")).FullName;
        var link = Path.Combine(dir, "link");
        try
        {
            var layout = CopyOfCorpusInto(real, Corpus.SalesInvoice, "J.docx");
            FileSystemLinks.CreateJunction(link, real);

            var response = EditTools.InsertText(Path.Combine(link, "J.docx"), "Inserted through the junction", "documentEnd");

            Assert.True(response.Ok, response.Error?.Message);
            Assert.Contains("Inserted through the junction", BodyText(layout), StringComparison.Ordinal);
            Assert.Empty(Directory.EnumerateFiles(real, ".bcwl-*"));
        }
        finally
        {
            FileSystemLinks.DeleteJunction(link);
            Directory.Delete(dir, recursive: true);
        }
    }

    [WindowsOnlyFact]
    public void An_edit_through_a_file_symlink_updates_the_target_and_keeps_the_link()
    {
        // Renaming the staged copy over the LINK replaced it with a plain file: the real layout never saw the
        // edit, and another host editing the real name behind the same (resolved) lock kept its own version.
        var dir = PrivateDirectory();
        try
        {
            var layout = CopyOfCorpusInto(dir, Corpus.SalesInvoice, "real.docx");
            var link = Path.Combine(dir, "link.docx");
            if (!FileSystemLinks.TryCreateFileSymlink(link, layout))
            {
                return; // This account may not create symlinks; junctions cover the lock-key half above.
            }

            var response = EditTools.InsertText(link, "Inserted through the symlink", "documentEnd");

            Assert.True(response.Ok, response.Error?.Message);
            Assert.Contains("Inserted through the symlink", BodyText(layout), StringComparison.Ordinal);
            Assert.NotNull(new FileInfo(link).LinkTarget);
            Assert.True(PathIdentity.SameFile(link, layout));
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [WindowsOnlyFact]
    public void Merge_refuses_an_outputPath_that_is_the_layout_under_another_spelling()
    {
        var dir = PrivateDirectory();
        var real = Directory.CreateDirectory(Path.Combine(dir, "real")).FullName;
        var link = Path.Combine(dir, "link");
        try
        {
            var layout = CopyOfCorpusInto(real, Corpus.SalesInvoice, "J.docx");
            var before = File.ReadAllBytes(layout);
            FileSystemLinks.CreateJunction(link, real);

            var ex = Assert.Throws<ArgumentException>(() => MergeEngine.Merge(layout, Path.Combine(link, "J.docx")));

            Assert.Equal("outputPath", ex.ParamName);
            Assert.Equal(before, File.ReadAllBytes(layout));
        }
        finally
        {
            FileSystemLinks.DeleteJunction(link);
            Directory.Delete(dir, recursive: true);
        }
    }
}
