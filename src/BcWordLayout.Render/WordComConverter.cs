using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace BcWordLayout.Render;

/// <summary>
/// Converts .docx to PDF via Word COM automation (<c>Word.Application</c>) — the primary, highest-fidelity
/// converter; every target developer machine is expected to have Word installed. Windows-only, driven
/// entirely through reflection (<see cref="Type.InvokeMember(string, BindingFlags, Binder?, object?, object?[]?)"/>)
/// rather than <c>dynamic</c>, so the assembly can still target plain <c>net10.0</c> instead of
/// <c>net10.0-windows</c>. Every failure mode (missing input, Word not installed, timeout, COM error,
/// no/invalid output) is reported as a structured <see cref="PdfConversionResult"/> with <c>Ok = false</c>;
/// <see cref="Convert"/> never throws — its entire body, including the synchronous prologue before any
/// worker task starts, is a hard try/catch boundary — never shows a dialog (<c>Visible = false</c>,
/// <c>DisplayAlerts</c> forced off), and never leaves a zombie WINWORD.EXE behind.
/// </summary>
[SupportedOSPlatform("windows")]
public sealed class WordComConverter : IPdfConverter
{
    private const int WdAlertsNone = 0;
    private const int WdDoNotSaveChanges = 0;
    private const int WdExportFormatPdf = 17;
    private const int MsoAutomationSecurityForceDisable = 3;

    /// <summary>
    /// Serializes every Word COM conversion made through this converter. Word.Application automation is
    /// not safe to drive concurrently, and <see cref="RunConversion"/> identifies "its" newly spawned
    /// WINWORD.EXE process by diffing process-id snapshots taken immediately before/after
    /// <c>Activator.CreateInstance</c> — a second conversion racing that same window could misidentify
    /// which process is its own. Holding this lock for the whole attempt (including the timeout wait)
    /// guarantees at most one conversion is ever actively in flight, which is what makes that pid-diff
    /// dependable.
    /// </summary>
    private static readonly object ConversionLock = new();

    /// <inheritdoc/>
    public string Name => "word-com";

    /// <inheritdoc/>
    public bool IsAvailable => OperatingSystem.IsWindows() && Type.GetTypeFromProgID("Word.Application") is not null;

    /// <inheritdoc/>
    public PdfConversionResult Convert(string docxPath, string pdfPath, PdfConversionOptions? options = null)
    {
        var sw = Stopwatch.StartNew();

        try
        {
            options ??= new PdfConversionOptions();

            if (!File.Exists(docxPath))
            {
                return PdfConversionResult.Failure(Name, $"Input .docx not found: '{docxPath}'.", sw.Elapsed);
            }

            if (!IsAvailable)
            {
                return PdfConversionResult.Failure(
                    Name,
                    "Microsoft Word is not installed, or not registered as a COM automation server ('Word.Application').",
                    sw.Elapsed);
            }

            if (options.Timeout <= TimeSpan.Zero)
            {
                return PdfConversionResult.Failure(Name, "Timeout must be positive.", sw.Elapsed);
            }

            if (File.Exists(pdfPath) && !options.Overwrite)
            {
                return PdfConversionResult.Failure(
                    Name, $"Destination already exists and Overwrite is false: '{pdfPath}'.", sw.Elapsed);
            }

            // Path.GetFullPath throws on an empty/whitespace/otherwise malformed pdfPath, and
            // Directory.CreateDirectory can throw too (e.g. an un-creatable destination) — both run before
            // any worker task exists, so only the outer catch below (not the worker's own) protects them.
            var fullDocxPath = Path.GetFullPath(docxPath);
            var fullPdfPath = Path.GetFullPath(pdfPath);
            var destDir = Path.GetDirectoryName(fullPdfPath);
            if (!string.IsNullOrEmpty(destDir))
            {
                Directory.CreateDirectory(destDir);
            }

            lock (ConversionLock)
            {
                return RunWithTimeout(fullDocxPath, fullPdfPath, options, sw);
            }
        }
        catch (Exception ex)
        {
            // Hard boundary contract: Convert never throws, for any reason, anywhere in its body.
            return PdfConversionResult.Failure(Name, ex.Message, sw.Elapsed);
        }
    }

    /// <summary>
    /// Starts the worker task that drives Word and waits up to <c>options.Timeout</c> for it, killing the
    /// tracked WINWORD process and returning a timeout failure if it does not finish in time. Must only be
    /// called while holding <see cref="ConversionLock"/> — see that field's remarks.
    /// </summary>
    private PdfConversionResult RunWithTimeout(
        string fullDocxPath, string fullPdfPath, PdfConversionOptions options, Stopwatch sw)
    {
        var context = new ConversionContext();
        Exception? failure = null;

        // Run the whole Open -> ExportAsFixedFormat -> Close -> Quit sequence on a worker task so a hang
        // (e.g. Word blocked on a dialog we failed to suppress) can be timed out from here instead of
        // wedging the caller forever.
        var worker = Task.Run(() =>
        {
            try
            {
                RunConversion(fullDocxPath, fullPdfPath, context);
            }
            catch (Exception ex)
            {
                failure = ex;
            }
        });

        var waitMs = (int)Math.Min(options.Timeout.TotalMilliseconds, int.MaxValue);
        if (!worker.Wait(waitMs))
        {
            KillNow(context.Process);

            // Best-effort only, bounded: give the now-killed worker's own teardown (COM release, GC,
            // process dispose in RunConversion's finally) a short extra moment to finish before
            // ConversionLock is released, so the next queued conversion is less likely to start while it is
            // still unwinding. A worker that never unwinds must not hang the lock forever, hence the cap.
            worker.Wait(TimeSpan.FromSeconds(5));

            return PdfConversionResult.Failure(Name, $"Word COM conversion timed out after {options.Timeout}.", sw.Elapsed);
        }

        if (failure is not null)
        {
            return PdfConversionResult.Failure(Name, $"Word COM conversion failed: {failure.Message}", sw.Elapsed);
        }

        if (!PdfFileValidation.LooksLikePdf(fullPdfPath))
        {
            return PdfConversionResult.Failure(Name, $"Word did not produce a valid PDF at '{fullPdfPath}'.", sw.Elapsed);
        }

        return PdfConversionResult.Success(Name, fullPdfPath, sw.Elapsed);
    }

    /// <summary>
    /// Runs the full Open → ExportAsFixedFormat → Close → Quit sequence against a freshly created Word
    /// instance. Records the newly spawned WINWORD process on <paramref name="context"/> (by diffing the
    /// WINWORD pids before/after <c>Activator.CreateInstance</c>) so <see cref="RunWithTimeout"/> can kill
    /// it on timeout, and — regardless of success or failure — always releases every COM reference,
    /// force-stops that same process if it is somehow still alive, and disposes it exactly once.
    /// </summary>
    private static void RunConversion(string docxPath, string pdfPath, ConversionContext context)
    {
        object? app = null;
        object? documents = null;
        object? document = null;

        try
        {
            var beforePids = SnapshotWordProcessIds();

            var wordType = Type.GetTypeFromProgID("Word.Application")
                ?? throw new InvalidOperationException("Word.Application COM type is not registered.");

            app = Activator.CreateInstance(wordType)
                ?? throw new InvalidOperationException("Failed to create a Word.Application COM instance.");

            context.Process = FindNewWordProcess(beforePids);

            SetProperty(app, "Visible", false);
            SetProperty(app, "DisplayAlerts", WdAlertsNone);
            TrySetProperty(app, "ScreenUpdating", false);
            TrySetProperty(app, "AutomationSecurity", MsoAutomationSecurityForceDisable);

            documents = GetProperty(app, "Documents");
            document = InvokeMethod(documents!, "Open", docxPath);

            InvokeMethod(document!, "ExportAsFixedFormat", pdfPath, WdExportFormatPdf);

            InvokeMethod(document!, "Close", WdDoNotSaveChanges);
            ReleaseCom(document);
            document = null;

            InvokeMethod(app, "Quit", 0);
        }
        finally
        {
            ReleaseCom(document);
            ReleaseCom(documents);
            ReleaseCom(app);

            // COM RCWs release their underlying interface pointers on finalization; force it now rather
            // than waiting on the GC so a graceful Quit() truly lets WINWORD.EXE exit before the belt-and-
            // braces kill check below.
            GC.Collect();
            GC.WaitForPendingFinalizers();

            // Dispose exactly once, here, having finished every use of the tracked process (EnsureStopped's
            // own best-effort kill included) — this is the only place that touches it after RunWithTimeout's
            // timeout branch may have already (also best-effort) tried to kill it.
            var proc = context.Process;
            if (proc is not null)
            {
                EnsureStopped(proc);
                proc.Dispose();
                context.Process = null;
            }
        }
    }

    /// <summary>
    /// Carries the tracked WINWORD process between the worker thread (<see cref="RunConversion"/>, which
    /// sets it once, early on, and clears it after disposing it) and the calling thread
    /// (<see cref="RunWithTimeout"/>, which reads it only on a timeout, to kill it). One instance is
    /// created per <see cref="Convert"/> call, so — thanks to <see cref="ConversionLock"/> — at most one
    /// worker/caller pair ever touches a given instance concurrently; <see cref="Process"/> is still backed
    /// by a <c>volatile</c> field so the calling thread is guaranteed to observe the worker's write rather
    /// than a stale cached value.
    /// </summary>
    private sealed class ConversionContext
    {
        private volatile Process? _process;

        public Process? Process
        {
            get => _process;
            set => _process = value;
        }
    }

    // ---- WINWORD process tracking (for timeout-kill and zombie cleanup) ----

    private static HashSet<int> SnapshotWordProcessIds()
    {
        var processes = Process.GetProcessesByName("WINWORD");
        var ids = new HashSet<int>(processes.Select(p => p.Id));
        foreach (var p in processes)
        {
            p.Dispose();
        }

        return ids;
    }

    /// <summary>Returns the first WINWORD process not present in <paramref name="beforePids"/>, disposing every other (unrelated, pre-existing) candidate this query returns.</summary>
    private static Process? FindNewWordProcess(HashSet<int> beforePids)
    {
        Process? found = null;
        foreach (var p in Process.GetProcessesByName("WINWORD"))
        {
            if (found is null && !beforePids.Contains(p.Id))
            {
                found = p;
            }
            else
            {
                p.Dispose();
            }
        }

        return found;
    }

    private static void KillNow(Process? process)
    {
        if (process is null)
        {
            return;
        }

        try
        {
            if (!process.HasExited)
            {
                process.Kill(entireProcessTree: true);
            }
        }
        catch (Exception)
        {
            // Best-effort kill only, deliberately broad: Kill(entireProcessTree: true) can throw
            // AggregateException per the BCL docs, in addition to InvalidOperationException/Win32Exception
            // for an already-exited/exiting process — none of which may escape Convert.
        }
    }

    /// <summary>Gives a normally-quitting Word a brief grace period to exit on its own, then force-kills it if it is still alive.</summary>
    private static void EnsureStopped(Process process)
    {
        try
        {
            if (process.HasExited)
            {
                return;
            }

            process.WaitForExit(3000);
        }
        catch (Exception)
        {
            return;
        }

        KillNow(process);
    }

    // ---- reflection COM helpers (no `dynamic`, so the project can stay plain net10.0) ----

    private static object? GetProperty(object target, string name) =>
        target.GetType().InvokeMember(name, BindingFlags.GetProperty, null, target, null, CultureInfo.InvariantCulture);

    private static void SetProperty(object target, string name, object? value) =>
        target.GetType().InvokeMember(name, BindingFlags.SetProperty, null, target, new[] { value }, CultureInfo.InvariantCulture);

    /// <summary>
    /// Best-effort counterpart of <see cref="SetProperty"/> for niceties some Word versions/builds may not
    /// expose or may refuse to set — never lets a missing optional property (e.g. <c>ScreenUpdating</c>)
    /// fail the whole conversion. Deliberately swallows every exception shape: late-bound COM property
    /// sets can fail as <see cref="MissingMethodException"/>, a wrapped <see cref="COMException"/>, or
    /// otherwise, none of which are load-bearing here.
    /// </summary>
    private static void TrySetProperty(object target, string name, object? value)
    {
        try
        {
            SetProperty(target, name, value);
        }
        catch (Exception)
        {
            // Intentionally ignored — see summary.
        }
    }

    private static object? InvokeMethod(object target, string name, params object?[] args) =>
        target.GetType().InvokeMember(name, BindingFlags.InvokeMethod, null, target, args, CultureInfo.InvariantCulture);

    private static void ReleaseCom(object? comObject)
    {
        if (comObject is null)
        {
            return;
        }

        try
        {
            if (Marshal.IsComObject(comObject))
            {
                Marshal.FinalReleaseComObject(comObject);
            }
        }
        catch (COMException)
        {
            // Best-effort release only; must not mask the conversion's real outcome.
        }
    }
}
