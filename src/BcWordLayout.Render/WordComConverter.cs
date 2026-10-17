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
/// <c>DisplayAlerts</c> forced off), always asks the Word instance it started to quit, and never kills a
/// WINWORD.EXE it has not positively identified as that instance.
/// </summary>
/// <remarks>
/// <para>
/// WHICH WINWORD.EXE IS OURS. A conversion used to take the first WINWORD.EXE that appeared while
/// <c>Word.Application</c> was being created as its own, and force-killed it after every conversion (or at
/// once on a timeout). A second server process converting at the same time (two IDE windows, or the test
/// suite next to an agent's preview) or the user opening Word in that window made it kill somebody else's
/// Word, possibly with unsaved documents (GitHub issue #103). The instance is now identified POSITIVELY:
/// right after creation it gets a unique <c>Caption</c>, and the process owning the (hidden) Word window that
/// carries that caption is ours, provided it did not exist before the creation. If it did, automation
/// connected to a Word that was already running, and the conversion stops without touching it any further.
/// An instance whose window cannot be found is never killed; it is only asked to <c>Quit</c> through COM.
/// </para>
/// <para>
/// A WORD THAT HANGS BEFORE IT EXISTS AS A COM OBJECT. <c>Activator.CreateInstance</c> can block while Word
/// sits on an invisible first-run, licensing or safe-mode prompt, before any caption can be set. On a timeout
/// in that state the conversion kills the one WINWORD.EXE that appeared since the conversion started AND was
/// launched by COM (<see cref="WordProcessIdentity.IsComLaunched"/>; a Word the user starts never is), and only
/// when exactly one such process exists: two mean another automation client is starting Word at the same
/// moment, and neither is killed.
/// </para>
/// <para>
/// A TIMED-OUT WORKER NEVER FINISHES THE JOB. The worker checks a cancellation flag after creation, before
/// <c>Open</c> and before <c>ExportAsFixedFormat</c>. A worker that regains control after its conversion has
/// already timed out (creation finally returned, or a hung <c>Open</c> recovered) quits its Word instead of
/// exporting to a PDF path the next preview may already own.
/// </para>
/// </remarks>
[SupportedOSPlatform("windows")]
public sealed class WordComConverter : IPdfConverter
{
    private const int WdAlertsNone = 0;
    private const int WdDoNotSaveChanges = 0;
    private const int WdExportFormatPdf = 17;
    private const int MsoAutomationSecurityForceDisable = 3;

    /// <summary>
    /// Serializes every Word COM conversion made in this process. Word.Application automation is not safe to
    /// drive concurrently from one process, and the hung-startup fallback (see the class remarks) compares the
    /// WINWORD.EXE processes that appeared since THIS conversion started, which another conversion of this
    /// process starting Word at the same moment would blur. Held for the whole attempt, timeout wait
    /// included. Conversions in OTHER processes are not serialized: each identifies its own instance by its
    /// own caption.
    /// </summary>
    private static readonly object ConversionLock = new();

    private readonly Func<object?> _createApplication;
    private readonly Func<bool> _isAvailable;
    private readonly TimeSpan _identifyTimeout;

    /// <summary>Creates the converter for the Word installed on this machine.</summary>
    public WordComConverter()
        : this(CreateWordApplication, IsWordRegistered, TimeSpan.FromSeconds(2))
    {
    }

    /// <summary>
    /// Test seam: <paramref name="createApplication"/> stands in for <c>Activator.CreateInstance</c> of
    /// <c>Word.Application</c> (a test passes a plain .NET object with the same members, or one that blocks to
    /// simulate a Word hung at startup), <paramref name="isAvailable"/> for the COM registration check, and
    /// <paramref name="identifyTimeout"/> bounds the search for the instance's window.
    /// </summary>
    internal WordComConverter(Func<object?> createApplication, Func<bool> isAvailable, TimeSpan identifyTimeout)
    {
        _createApplication = createApplication;
        _isAvailable = isAvailable;
        _identifyTimeout = identifyTimeout;
    }

    /// <inheritdoc/>
    public string Name => "word-com";

    /// <inheritdoc/>
    public bool IsAvailable => _isAvailable();

    private static bool IsWordRegistered() =>
        OperatingSystem.IsWindows() && Type.GetTypeFromProgID("Word.Application") is not null;

    private static object? CreateWordApplication()
    {
        var wordType = Type.GetTypeFromProgID("Word.Application")
            ?? throw new InvalidOperationException("Word.Application COM type is not registered.");
        return Activator.CreateInstance(wordType);
    }

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
    /// Starts the worker task that drives Word and waits up to <c>options.Timeout</c> for it. On a timeout it
    /// cancels the worker and kills the Word instance only when that instance is positively identified (see
    /// the class remarks). Must only be called while holding <see cref="ConversionLock"/>.
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
            // Cancel FIRST: from here on a worker that regains control quits its Word instead of carrying on.
            context.Cancelled = true;
            var killed = KillOnTimeout(context);

            // Best-effort only, bounded: give the worker's own teardown (COM release, GC, process dispose in
            // RunConversion's finally) a short extra moment to finish before ConversionLock is released, so the
            // next queued conversion is less likely to start while it is still unwinding. A worker that never
            // unwinds must not hang the lock forever, hence the cap.
            worker.Wait(TimeSpan.FromSeconds(5));

            var message = $"Word COM conversion timed out after {options.Timeout}.";
            if (!killed)
            {
                message += " The Word instance it started could not be identified with certainty, so it was not "
                    + "killed; it has been told to quit, but a hidden WINWORD.EXE may remain until it does.";
            }

            return PdfConversionResult.Failure(Name, message, sw.Elapsed);
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
    /// Kills this conversion's Word on a timeout, when it can be told apart from every other WINWORD.EXE: the
    /// instance identified by its caption, or, while creation is still blocked, the single new COM-launched
    /// WINWORD.EXE. Returns false when it killed nothing.
    /// </summary>
    private static bool KillOnTimeout(ConversionContext context)
    {
        var identified = context.Process;
        if (identified is not null)
        {
            KillNow(identified);
            return true;
        }

        if (context.Stage == ConversionStage.Creating && context.BeforePids is { } beforePids)
        {
            using var starting = FindSingleNewComLaunchedWordProcess(beforePids);
            if (starting is not null)
            {
                KillNow(starting);
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Runs the full create → identify → Open → ExportAsFixedFormat → Close → Quit sequence. Records the
    /// identified WINWORD process on <paramref name="context"/> so <see cref="RunWithTimeout"/> can kill it on a
    /// timeout, and — regardless of success or failure — asks the instance to quit, releases every COM
    /// reference, and force-stops the identified process if it is somehow still alive.
    /// </summary>
    private void RunConversion(string docxPath, string pdfPath, ConversionContext context)
    {
        object? app = null;
        object? documents = null;
        object? document = null;

        // Set once `app` is known to be an instance this conversion may quit: false only while it is unknown,
        // and when automation turned out to have connected to a Word that was already running.
        var quitOnExit = false;

        try
        {
            context.BeforePids = SnapshotWordProcessIds();
            context.Stage = ConversionStage.Creating;
            app = _createApplication()
                ?? throw new InvalidOperationException("Failed to create a Word.Application COM instance.");
            context.Stage = ConversionStage.Created;

            quitOnExit = IdentifyInstance(app, context);
            ThrowIfCancelled(context);

            SetProperty(app, "Visible", false);
            SetProperty(app, "DisplayAlerts", WdAlertsNone);
            TrySetProperty(app, "ScreenUpdating", false);
            TrySetProperty(app, "AutomationSecurity", MsoAutomationSecurityForceDisable);

            documents = GetProperty(app, "Documents");
            ThrowIfCancelled(context);
            document = InvokeMethod(documents!, "Open", docxPath);

            ThrowIfCancelled(context);
            InvokeMethod(document!, "ExportAsFixedFormat", pdfPath, WdExportFormatPdf);

            InvokeMethod(document!, "Close", WdDoNotSaveChanges);
            ReleaseCom(document);
            document = null;
        }
        finally
        {
            // Asked on every path, not only after a successful export: an instance that is not positively
            // identified can never be killed, so a polite Quit is the only way it goes away on failure too.
            if (app is not null && quitOnExit)
            {
                TryInvokeMethod(app, "Quit", WdDoNotSaveChanges);
            }

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
    /// Identifies <paramref name="app"/>'s process by giving the instance a unique caption and finding the Word
    /// window that carries it (see the class remarks). Records the process on <paramref name="context"/> when
    /// it is confirmed and new since <see cref="ConversionContext.BeforePids"/>. Returns whether the instance
    /// may be quit at the end; throws, without quitting anything, when automation connected to a Word that was
    /// already running.
    /// </summary>
    private bool IdentifyInstance(object app, ConversionContext context)
    {
        var token = "bcwl-" + Guid.NewGuid().ToString("N");
        var originalCaption = TryGetProperty(app, "Caption");
        if (!TrySetProperty(app, "Caption", token))
        {
            return true; // Unconfirmed: quit at the end, never kill.
        }

        var processId = WordProcessIdentity.FindWordWindowProcessId(token, _identifyTimeout);
        if (processId is not { } pid)
        {
            return true; // Unconfirmed: quit at the end, never kill.
        }

        if (context.BeforePids?.Contains(pid) == true)
        {
            // Not an instance this conversion started: leave it exactly as it was, and never quit or kill it.
            TrySetProperty(app, "Caption", originalCaption);
            throw new InvalidOperationException(
                $"Word automation connected to a Word instance that was already running (process {pid}) instead "
                + "of starting its own, so it was left alone. Close Word and retry.");
        }

        try
        {
            context.Process = Process.GetProcessById(pid);
        }
        catch (ArgumentException)
        {
            // The process exited between the window lookup and here; nothing left to track.
        }

        return true;
    }

    private static void ThrowIfCancelled(ConversionContext context)
    {
        if (context.Cancelled)
        {
            throw new OperationCanceledException("The conversion already timed out.");
        }
    }

    /// <summary>Where the worker is, as far as the timeout path needs to know.</summary>
    private enum ConversionStage
    {
        Starting,
        Creating,
        Created,
    }

    /// <summary>
    /// Carries state between the worker thread (<see cref="RunConversion"/>) and the calling thread
    /// (<see cref="RunWithTimeout"/>, which reads it only on a timeout). One instance is created per
    /// <see cref="Convert"/> call, so — thanks to <see cref="ConversionLock"/> — at most one worker/caller pair
    /// ever touches a given instance concurrently; every member is backed by a <c>volatile</c> field so each
    /// thread observes the other's latest write rather than a stale cached value.
    /// </summary>
    private sealed class ConversionContext
    {
        private volatile Process? _process;
        private volatile HashSet<int>? _beforePids;
        private volatile int _stage;
        private volatile bool _cancelled;

        /// <summary>The positively identified WINWORD process, set once by the worker and cleared after disposal.</summary>
        public Process? Process
        {
            get => _process;
            set => _process = value;
        }

        /// <summary>The WINWORD process ids that existed before this conversion created its instance.</summary>
        public HashSet<int>? BeforePids
        {
            get => _beforePids;
            set => _beforePids = value;
        }

        public ConversionStage Stage
        {
            get => (ConversionStage)_stage;
            set => _stage = (int)value;
        }

        /// <summary>Set by the caller on a timeout; the worker stops at its next check and quits its Word.</summary>
        public bool Cancelled
        {
            get => _cancelled;
            set => _cancelled = value;
        }
    }

    // ---- WINWORD process tracking (for the timeout kill and zombie cleanup) ----

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

    /// <summary>
    /// The one WINWORD process that is not in <paramref name="beforePids"/> and was launched by COM, or null
    /// when there is none or more than one (see the class remarks). Disposes every other candidate.
    /// </summary>
    private static Process? FindSingleNewComLaunchedWordProcess(HashSet<int> beforePids)
    {
        var matches = new List<Process>();
        foreach (var p in Process.GetProcessesByName("WINWORD"))
        {
            if (!beforePids.Contains(p.Id) && WordProcessIdentity.IsComLaunched(WordProcessIdentity.TryReadCommandLine(p)))
            {
                matches.Add(p);
            }
            else
            {
                p.Dispose();
            }
        }

        if (matches.Count == 1)
        {
            return matches[0];
        }

        foreach (var p in matches)
        {
            p.Dispose();
        }

        return null;
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
    /// Best-effort counterpart of <see cref="GetProperty"/>: null when the property is missing or refuses.
    /// </summary>
    private static object? TryGetProperty(object target, string name)
    {
        try
        {
            return GetProperty(target, name);
        }
        catch (Exception)
        {
            return null; // Deliberately broad - see TrySetProperty.
        }
    }

    /// <summary>
    /// Best-effort counterpart of <see cref="SetProperty"/> for niceties some Word versions/builds may not
    /// expose or may refuse to set — never lets a missing optional property (e.g. <c>ScreenUpdating</c>)
    /// fail the whole conversion. Deliberately swallows every exception shape: late-bound COM property
    /// sets can fail as <see cref="MissingMethodException"/>, a wrapped <see cref="COMException"/>, or
    /// otherwise, none of which are load-bearing here. Returns whether the set succeeded.
    /// </summary>
    private static bool TrySetProperty(object target, string name, object? value)
    {
        try
        {
            SetProperty(target, name, value);
            return true;
        }
        catch (Exception)
        {
            return false; // Intentionally ignored — see summary.
        }
    }

    private static object? InvokeMethod(object target, string name, params object?[] args) =>
        target.GetType().InvokeMember(name, BindingFlags.InvokeMethod, null, target, args, CultureInfo.InvariantCulture);

    /// <summary>Best-effort <see cref="InvokeMethod"/> for teardown calls (<c>Quit</c>) that must not mask the real outcome.</summary>
    private static void TryInvokeMethod(object target, string name, params object?[] args)
    {
        try
        {
            InvokeMethod(target, name, args);
        }
        catch (Exception)
        {
            // Deliberately broad: a Word that already exited, or was killed by the timeout path, fails here
            // with a COM disconnect; the conversion's own outcome is what gets reported.
        }
    }

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
