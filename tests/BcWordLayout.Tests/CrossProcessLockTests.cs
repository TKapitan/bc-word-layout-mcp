using BcWordLayout.McpHost;
using BcWordLayout.McpHost.Tools;

namespace BcWordLayout.Tests;

/// <summary>
/// Covers the cross-process half of the edit lock and its read/edit-coordination
/// use: <see cref="CrossProcessLock"/>'s own naming and abandoned-mutex handling,
/// the <c>file_locked</c> envelope a timed-out lock wait produces for mutating/read/preview/create tools, and
/// the <c>file_locked</c> mapping for a genuine Windows sharing violation (the layout open in another program,
/// e.g. Word). Every test targets its own temp COPY of a corpus file, never the shared corpus itself.
/// </summary>
/// <remarks>Joins the preview-converter-seam collection because one test calls
/// <c>LifecycleTools.PreviewLayout</c> (see <see cref="PreviewConverterSeamCollection"/> for the rule that
/// governs membership).</remarks>
[Collection("preview-converter-seam")]
public class CrossProcessLockTests
{
    private static string CopyOfCorpus(string corpusFile)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bcwl-crosslock-{Guid.NewGuid():N}.docx");
        File.Copy(Corpus.Path(corpusFile), path, overwrite: true);
        return path;
    }

    /// <summary>
    /// Starts a background thread that immediately acquires <see cref="CrossProcessLock"/> for
    /// <paramref name="path"/> and holds it until <paramref name="releaseHolder"/> is set (or a 15s safety
    /// timeout elapses) - the same "hold on one thread, contend from another" shape
    /// <c>McpHostToolTests.PreviewLayout_serializes_against_a_concurrent_holder_of_the_same_layouts_edit_lock</c>
    /// already uses for the in-process lock. Blocks the caller until the holder thread reports it actually
    /// acquired the lock, so the caller's own contended attempt is never racing the holder's own acquire.
    /// </summary>
    private static Thread StartHolderThread(string path, ManualResetEventSlim holderReady, ManualResetEventSlim releaseHolder)
    {
        var thread = new Thread(() =>
        {
            using var held = CrossProcessLock.TryAcquire(path, TimeSpan.FromSeconds(15));
            holderReady.Set();
            releaseHolder.Wait(TimeSpan.FromSeconds(15));
        });
        thread.Start();
        return thread;
    }

    // ---- CrossProcessLock itself: naming + abandoned-mutex handling ----

    [Fact]
    public void MutexName_is_stable_case_insensitive_and_distinct_per_path()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        var otherPath = CopyOfCorpus(Corpus.StandardStatement);
        try
        {
            var name = CrossProcessLock.MutexName(path);

            Assert.StartsWith("Local\\bcwl-edit-", name, StringComparison.Ordinal);
            Assert.Equal(name, CrossProcessLock.MutexName(path.ToUpperInvariant()));
            Assert.Equal(name, CrossProcessLock.MutexName(path.ToLowerInvariant()));
            Assert.NotEqual(name, CrossProcessLock.MutexName(otherPath));
        }
        finally
        {
            File.Delete(path);
            File.Delete(otherPath);
        }
    }

    [Fact]
    public void TryAcquire_blocks_a_different_thread_until_released_then_succeeds_again()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var holderReady = new ManualResetEventSlim(false);
            var releaseHolder = new ManualResetEventSlim(false);
            var holderThread = StartHolderThread(path, holderReady, releaseHolder);

            try
            {
                Assert.True(holderReady.Wait(TimeSpan.FromSeconds(5)), "holder thread failed to acquire in time");

                using var contended = CrossProcessLock.TryAcquire(path, TimeSpan.FromMilliseconds(200));
                Assert.False(contended.Acquired, "a second thread must NOT acquire while the first still holds it");
            }
            finally
            {
                releaseHolder.Set();
                Assert.True(holderThread.Join(TimeSpan.FromSeconds(5)), "holder thread failed to exit in time");
            }

            using var afterRelease = CrossProcessLock.TryAcquire(path, TimeSpan.FromSeconds(5));
            Assert.True(afterRelease.Acquired, "the path must be acquirable again once the holder released it");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void An_abandoned_mutex_is_treated_as_a_successful_acquire()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            var name = CrossProcessLock.MutexName(path);
            var abandonerAcquired = new ManualResetEventSlim(false);

            var abandoner = new Thread(() =>
            {
                var mutex = new Mutex(initiallyOwned: false, name);
                Assert.True(mutex.WaitOne(TimeSpan.FromSeconds(5)));
                abandonerAcquired.Set();
                // Deliberately exit WITHOUT ReleaseMutex/Dispose - the OS marks the mutex abandoned the
                // moment this owning thread terminates, exactly like a host process that crashes mid-edit
                // (see CrossProcessLock's own remarks on why that is still safe to treat as an ordinary
                // successful acquire).
            })
            {
                IsBackground = true,
            };
            abandoner.Start();

            Assert.True(abandonerAcquired.Wait(TimeSpan.FromSeconds(5)), "abandoner thread failed to acquire in time");
            Assert.True(abandoner.Join(TimeSpan.FromSeconds(5)), "abandoner thread should have exited by now");

            using var handle = CrossProcessLock.TryAcquire(path, TimeSpan.FromSeconds(5));
            Assert.True(handle.Acquired, "an abandoned mutex must still be reported as a successful acquire");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ---- tool-level: a held cross-process lock times out to file_locked, and clears once released ----

    [Fact]
    public void InsertField_times_out_with_file_locked_while_another_thread_holds_the_cross_process_lock()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        var originalTimeout = CrossProcessLock.MutatingTimeout;
        try
        {
            CrossProcessLock.MutatingTimeout = TimeSpan.FromMilliseconds(300);

            var holderReady = new ManualResetEventSlim(false);
            var releaseHolder = new ManualResetEventSlim(false);
            var holderThread = StartHolderThread(path, holderReady, releaseHolder);

            try
            {
                Assert.True(holderReady.Wait(TimeSpan.FromSeconds(5)), "holder thread failed to acquire in time");

                var response = EditTools.InsertField(path, "/Header/CustomerAddress1", "documentEnd");

                Assert.False(response.Ok);
                Assert.Null(response.Data);
                Assert.Equal("file_locked", response.Error!.Code);
                Assert.Contains("another process", response.Error!.Hint, StringComparison.OrdinalIgnoreCase);
            }
            finally
            {
                releaseHolder.Set();
                Assert.True(holderThread.Join(TimeSpan.FromSeconds(5)), "holder thread failed to exit in time");
            }

            // The failed attempt must have left the file untouched - a normal call now succeeds identically
            // to a fresh corpus copy (see InsertField_returns_ok_persists_to_disk_and_includes_a_passing_quickValidation).
            var followUp = EditTools.InsertField(path, "/Header/CustomerAddress1", "documentEnd");
            Assert.True(followUp.Ok, followUp.Error?.Message);
        }
        finally
        {
            CrossProcessLock.MutatingTimeout = originalTimeout;
            File.Delete(path);
        }
    }

    [Fact]
    public void GetLayoutInfo_times_out_with_file_locked_while_another_thread_holds_the_cross_process_lock()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        var originalTimeout = CrossProcessLock.ReadTimeout;
        try
        {
            CrossProcessLock.ReadTimeout = TimeSpan.FromMilliseconds(300);

            var holderReady = new ManualResetEventSlim(false);
            var releaseHolder = new ManualResetEventSlim(false);
            var holderThread = StartHolderThread(path, holderReady, releaseHolder);

            try
            {
                Assert.True(holderReady.Wait(TimeSpan.FromSeconds(5)), "holder thread failed to acquire in time");

                var response = ReadTools.GetLayoutInfo(path);

                Assert.False(response.Ok);
                Assert.Null(response.Data);
                Assert.Equal("file_locked", response.Error!.Code);
                Assert.Contains("another process", response.Error!.Hint, StringComparison.OrdinalIgnoreCase);
            }
            finally
            {
                releaseHolder.Set();
                Assert.True(holderThread.Join(TimeSpan.FromSeconds(5)), "holder thread failed to exit in time");
            }

            var followUp = ReadTools.GetLayoutInfo(path);
            Assert.True(followUp.Ok, followUp.Error?.Message);
        }
        finally
        {
            CrossProcessLock.ReadTimeout = originalTimeout;
            File.Delete(path);
        }
    }

    [Fact]
    public void CreateLayout_times_out_with_file_locked_while_another_thread_holds_the_outputPaths_cross_process_lock()
    {
        var outputPath = Path.Combine(Path.GetTempPath(), $"bcwl-crosslock-create-{Guid.NewGuid():N}.docx");
        var originalTimeout = CrossProcessLock.MutatingTimeout;
        try
        {
            CrossProcessLock.MutatingTimeout = TimeSpan.FromMilliseconds(300);

            var holderReady = new ManualResetEventSlim(false);
            var releaseHolder = new ManualResetEventSlim(false);
            var holderThread = StartHolderThread(outputPath, holderReady, releaseHolder);

            try
            {
                Assert.True(holderReady.Wait(TimeSpan.FromSeconds(5)), "holder thread failed to acquire in time");

                var response = LifecycleTools.CreateLayout(Corpus.Path(Corpus.SalesInvoice), outputPath);

                Assert.False(response.Ok);
                Assert.Equal("file_locked", response.Error!.Code);
                Assert.False(File.Exists(outputPath), "create_layout must not have written outputPath while the lock was held");
            }
            finally
            {
                releaseHolder.Set();
                Assert.True(holderThread.Join(TimeSpan.FromSeconds(5)), "holder thread failed to exit in time");
            }

            var followUp = LifecycleTools.CreateLayout(Corpus.Path(Corpus.SalesInvoice), outputPath);
            Assert.True(followUp.Ok, followUp.Error?.Message);
        }
        finally
        {
            CrossProcessLock.MutatingTimeout = originalTimeout;
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
        }
    }

    [Fact]
    public void PreviewLayout_times_out_with_file_locked_while_another_thread_holds_the_cross_process_lock()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        var originalTimeout = CrossProcessLock.MutatingTimeout;
        try
        {
            CrossProcessLock.MutatingTimeout = TimeSpan.FromMilliseconds(300);

            var holderReady = new ManualResetEventSlim(false);
            var releaseHolder = new ManualResetEventSlim(false);
            var holderThread = StartHolderThread(path, holderReady, releaseHolder);

            try
            {
                Assert.True(holderReady.Wait(TimeSpan.FromSeconds(5)), "holder thread failed to acquire in time");

                // Pinning converter to "libreoffice" is irrelevant to this failure (the lock times out before
                // any merge/convert step ever runs) but keeps this test from depending on Word COM being
                // installed, mirroring McpHostToolTests' own in-process lock test.
                var response = LifecycleTools.PreviewLayout(path, converter: "libreoffice");

                Assert.False(response.Ok);
                Assert.Equal("file_locked", response.Error!.Code);
            }
            finally
            {
                releaseHolder.Set();
                Assert.True(holderThread.Join(TimeSpan.FromSeconds(5)), "holder thread failed to exit in time");
            }
        }
        finally
        {
            CrossProcessLock.MutatingTimeout = originalTimeout;
            File.Delete(path);
        }
    }

    // ---- sharing violation (the layout open in another program, e.g. Word) -> file_locked ----

    [Fact]
    public void InsertField_when_the_layout_is_exclusively_open_elsewhere_returns_file_locked_with_a_close_it_in_word_hint()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using (new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None))
            {
                var response = EditTools.InsertField(path, "/Header/CustomerAddress1", "documentEnd");

                Assert.False(response.Ok);
                Assert.Null(response.Data);
                Assert.Equal("file_locked", response.Error!.Code);
                Assert.Contains("open in another program", response.Error!.Hint, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("Word", response.Error!.Hint, StringComparison.Ordinal);
            }

            // Once the exclusive handle is closed, the identical call succeeds - proving the file itself was
            // left untouched by the failed attempt (GuardMutate's working-copy-first design).
            var followUp = EditTools.InsertField(path, "/Header/CustomerAddress1", "documentEnd");
            Assert.True(followUp.Ok, followUp.Error?.Message);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void GetLayoutInfo_when_the_layout_is_exclusively_open_elsewhere_returns_file_locked()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using (new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None))
            {
                var response = ReadTools.GetLayoutInfo(path);

                Assert.False(response.Ok);
                Assert.Null(response.Data);
                Assert.Equal("file_locked", response.Error!.Code);
                Assert.Contains("open in another program", response.Error!.Hint, StringComparison.OrdinalIgnoreCase);
            }

            var followUp = ReadTools.GetLayoutInfo(path);
            Assert.True(followUp.Ok, followUp.Error?.Message);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void ValidateLayout_when_the_layout_is_exclusively_open_elsewhere_returns_file_locked()
    {
        var path = CopyOfCorpus(Corpus.SalesInvoice);
        try
        {
            using (new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None))
            {
                var response = ReadTools.ValidateLayout(path);

                Assert.False(response.Ok);
                Assert.Equal("file_locked", response.Error!.Code);
            }

            var followUp = ReadTools.ValidateLayout(path);
            Assert.True(followUp.Ok, followUp.Error?.Message);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
