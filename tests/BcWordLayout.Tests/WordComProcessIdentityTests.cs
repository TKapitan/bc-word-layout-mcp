using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using BcWordLayout.Merge;
using BcWordLayout.Render;

namespace BcWordLayout.Tests;

/// <summary>
/// GitHub issue #103: <see cref="WordComConverter"/> took the first WINWORD.EXE that appeared while it created
/// <c>Word.Application</c> as its own and killed it after every conversion, so another server's Word, or a Word
/// the user opened at that moment, could be force-killed. These tests drive the converter with a stand-in for
/// <c>Word.Application</c> (<see cref="FakeWordApplication"/>, a plain object with the members the converter
/// calls) and stand-in WINWORD.EXE processes (a renamed cmd.exe that idles), so they run without Word. The last
/// two use the real Word when it is installed and pass trivially otherwise, like every other Word test here.
/// </summary>
[SupportedOSPlatform("windows")]
public class WordComProcessIdentityTests
{
    private static string NewDirectory() => Directory.CreateTempSubdirectory("bcwl-wordid-").FullName;

    private static WordComConverter ConverterCreating(Func<object?> createApplication) =>
        new(createApplication, () => true, TimeSpan.FromMilliseconds(300));

    /// <summary>An input file for the fake: it is never read, but Convert requires it to exist.</summary>
    private static string InputDocx(string dir)
    {
        var path = Path.Combine(dir, "in.docx");
        File.WriteAllText(path, "read only by the real Word");
        return path;
    }

    [WindowsOnlyFact]
    public void A_WINWORD_started_by_someone_else_during_a_conversion_is_left_running()
    {
        // The old tracking took exactly this process as its own (the first WINWORD.EXE to appear while
        // Word.Application was being created) and killed it 3 s after the conversion: a Word the user opens
        // at that moment, with whatever documents it has open.
        var dir = NewDirectory();
        StandInWord? userWord = null;
        var app = new FakeWordApplication();
        try
        {
            var converter = ConverterCreating(() =>
            {
                userWord = StandInWord.Start(dir, comLaunched: false);
                return app;
            });

            var result = converter.Convert(InputDocx(dir), Path.Combine(dir, "out.pdf"));

            Assert.True(result.Ok, result.Error);
            Assert.True(userWord!.IsRunning, "a WINWORD.EXE the conversion did not start must never be killed");
            Assert.Equal(1, app.QuitCount);
        }
        finally
        {
            userWord?.Dispose();
            Directory.Delete(dir, recursive: true);
        }
    }

    [WindowsOnlyFact]
    public void A_timeout_while_Word_is_starting_kills_only_the_new_COM_launched_WINWORD()
    {
        // Word can hang inside CreateInstance (an invisible first-run, licensing or safe-mode prompt), before
        // it can be given a caption. The timeout then kills the one new WINWORD.EXE that COM started; one the
        // user started at the same moment carries no -Embedding switch and is left alone.
        var dir = NewDirectory();
        StandInWord? starting = null;
        StandInWord? userWord = null;
        try
        {
            var converter = ConverterCreating(() =>
            {
                userWord = StandInWord.Start(dir, comLaunched: false);
                starting = StandInWord.Start(dir, comLaunched: true);

                // CreateInstance only returns, with an error, once the server it was waiting for is gone.
                while (starting.IsRunning)
                {
                    Thread.Sleep(50);
                }

                throw new InvalidOperationException("The server process could not be started.");
            });

            var result = converter.Convert(
                InputDocx(dir), Path.Combine(dir, "out.pdf"), new PdfConversionOptions { Timeout = TimeSpan.FromSeconds(2) });

            Assert.False(result.Ok);
            Assert.Contains("timed out", result.Error, StringComparison.Ordinal);
            Assert.DoesNotContain("could not be identified", result.Error, StringComparison.Ordinal);
            Assert.False(starting!.IsRunning, "the COM-launched WINWORD.EXE that hung must be killed");
            Assert.True(userWord!.IsRunning, "a WINWORD.EXE the user started must never be killed");
        }
        finally
        {
            starting?.Dispose();
            userWord?.Dispose();
            Directory.Delete(dir, recursive: true);
        }
    }

    [WindowsOnlyFact]
    public void A_timeout_while_Word_is_starting_kills_nothing_when_two_COM_launched_WINWORDs_appeared()
    {
        // Two new COM-launched processes mean another automation client is starting Word at the same moment;
        // neither can be told apart from this conversion's, so neither is killed.
        var dir = NewDirectory();
        StandInWord? first = null;
        StandInWord? second = null;
        using var release = new ManualResetEventSlim(false);
        try
        {
            var converter = ConverterCreating(() =>
            {
                first = StandInWord.Start(dir, comLaunched: true);
                second = StandInWord.Start(dir, comLaunched: true);

                // Gives up shortly after the 2 s timeout fires (timed from here, not from the test's start: the
                // conversion may first wait for another test's conversion to release the shared lock), so the
                // converter's bounded wait for its worker stays short.
                release.Wait(TimeSpan.FromSeconds(2.5));
                throw new InvalidOperationException("The server process could not be started.");
            });

            var result = converter.Convert(
                InputDocx(dir), Path.Combine(dir, "out.pdf"), new PdfConversionOptions { Timeout = TimeSpan.FromSeconds(2) });

            Assert.False(result.Ok);
            Assert.Contains("could not be identified", result.Error, StringComparison.Ordinal);
            Assert.True(first!.IsRunning);
            Assert.True(second!.IsRunning);
        }
        finally
        {
            release.Set();
            first?.Dispose();
            second?.Dispose();
            Directory.Delete(dir, recursive: true);
        }
    }

    [WindowsOnlyFact]
    public void A_conversion_that_timed_out_never_exports_afterwards()
    {
        // A worker that regains control after its conversion timed out (here: Open recovers) used to carry on
        // and export to pdfPath outside every lock, racing the next preview. It now quits its Word instead.
        var dir = NewDirectory();
        var pdfPath = Path.Combine(dir, "out.pdf");
        using var openGate = new ManualResetEventSlim(false);
        var app = new FakeWordApplication(openGate);
        try
        {
            var converter = ConverterCreating(() => app);

            // Open recovers 1.5 s after it is entered, after the 1 s timeout has fired. Timed from Open, not from
            // the test's start: the conversion may first wait for another test's conversion to release the lock.
            _ = Task.Run(() =>
            {
                if (app.OpenEntered.Wait(TimeSpan.FromSeconds(60)))
                {
                    Thread.Sleep(TimeSpan.FromSeconds(1.5));
                }

                openGate.Set();
            });
            var result = converter.Convert(InputDocx(dir), pdfPath, new PdfConversionOptions { Timeout = TimeSpan.FromSeconds(1) });

            Assert.False(result.Ok);
            Assert.Contains("timed out", result.Error, StringComparison.Ordinal);
            Assert.True(app.Quitted.Wait(TimeSpan.FromSeconds(10)), "the worker must quit its Word once it regains control");
            Assert.False(app.ExportCalled, "a timed-out conversion must never export afterwards");
            Assert.False(File.Exists(pdfPath));
        }
        finally
        {
            openGate.Set();
            Directory.Delete(dir, recursive: true);
        }
    }

    [WindowsOnlyFact]
    public void The_command_line_of_another_process_is_read()
    {
        var dir = NewDirectory();
        try
        {
            using var standIn = StandInWord.Start(dir, comLaunched: true);

            var commandLine = WordProcessIdentity.TryReadCommandLine(standIn.Process);

            Assert.NotNull(commandLine);
            Assert.Contains("ping -n 60 127.0.0.1", commandLine, StringComparison.Ordinal);
            Assert.True(WordProcessIdentity.IsComLaunched(commandLine));
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [WindowsOnlyFact]
    public void Real_Word_is_found_by_its_caption_and_was_started_by_COM()
    {
        // Pins the two observations the identification rests on, against the Word installed here: the
        // instance's hidden main window carries the Caption set through automation, and the process COM
        // starts has -Embedding on its command line.
        var wordType = Type.GetTypeFromProgID("Word.Application");
        if (wordType is null)
        {
            return;
        }

        var app = Activator.CreateInstance(wordType)!;
        try
        {
            var token = "bcwl-test-" + Guid.NewGuid().ToString("N");
            wordType.InvokeMember("Caption", System.Reflection.BindingFlags.SetProperty, null, app, [token], System.Globalization.CultureInfo.InvariantCulture);

            var pid = WordProcessIdentity.FindWordWindowProcessId(token, TimeSpan.FromSeconds(5));

            Assert.NotNull(pid);
            using var process = Process.GetProcessById(pid.Value);
            Assert.Equal("WINWORD", process.ProcessName, ignoreCase: true);
            Assert.True(WordProcessIdentity.IsComLaunched(WordProcessIdentity.TryReadCommandLine(process)));
        }
        finally
        {
            try
            {
                wordType.InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, app, [0], System.Globalization.CultureInfo.InvariantCulture);
            }
            finally
            {
                Marshal.FinalReleaseComObject(app);
            }
        }
    }

    [WindowsOnlyFact]
    public void Real_Word_conversion_leaves_a_WINWORD_started_meanwhile_running_and_its_own_Word_gone()
    {
        var converter = new WordComConverter();
        if (!converter.IsAvailable)
        {
            return;
        }

        var dir = NewDirectory();
        StandInWord? userWord = null;
        try
        {
            var merged = Path.Combine(dir, "merged.docx");
            MergeEngine.Merge(Corpus.Path(Corpus.SalesInvoice), merged, new MergeOptions { Seed = 1, Rows = 1 });
            var before = WordProcessIds();

            // The real-Word counterpart of the first test: a WINWORD.EXE that appears while Word.Application is
            // being created is left alone, and the conversion's own, positively identified Word is gone
            // afterwards. Which new process the old tracking took depended on process-list order, so this run
            // alone does not reproduce GitHub issue #103 reliably; the first test does.
            var starter = Task.Run(async () =>
            {
                await Task.Delay(100);
                userWord = StandInWord.Start(dir, comLaunched: false);
            });
            var result = converter.Convert(merged, Path.Combine(dir, "out.pdf"));
            starter.Wait();

            Assert.True(result.Ok, result.Error);
            Assert.True(userWord!.IsRunning, "a WINWORD.EXE the conversion did not start must never be killed");

            // Still RUNNING, not merely listed: an exited WINWORD.EXE stays in the process list for a moment
            // while the COM launcher still holds a handle to it.
            var stillRunning = WordProcessIds().Except(before).Where(id => id != userWord.Process.Id && IsRunning(id)).ToList();
            Assert.True(stillRunning.Count == 0, $"the conversion's own Word must be gone, found {string.Join(", ", stillRunning)}");
        }
        finally
        {
            userWord?.Dispose();
            Directory.Delete(dir, recursive: true);
        }
    }

    private static bool IsRunning(int processId)
    {
        try
        {
            using var process = Process.GetProcessById(processId);
            return !process.HasExited;
        }
        catch (ArgumentException)
        {
            return false; // Already gone from the process list.
        }
    }

    private static HashSet<int> WordProcessIds()
    {
        var processes = Process.GetProcessesByName("WINWORD");
        var ids = processes.Select(p => p.Id).ToHashSet();
        foreach (var p in processes)
        {
            p.Dispose();
        }

        return ids;
    }

    /// <summary>
    /// A process named WINWORD.EXE that idles for a minute: a copy of cmd.exe running <c>ping</c>. With
    /// <c>comLaunched</c> its command line also ends in COM's <c>/Automation -Embedding</c>, as a Word that
    /// COM starts does.
    /// </summary>
    private sealed class StandInWord : IDisposable
    {
        private StandInWord(Process process) => Process = process;

        public Process Process { get; }

        public bool IsRunning
        {
            get
            {
                Process.Refresh();
                return !Process.HasExited;
            }
        }

        public static StandInWord Start(string dir, bool comLaunched)
        {
            var exe = Path.Combine(dir, "WINWORD.EXE");
            if (!File.Exists(exe))
            {
                File.Copy(Path.Combine(Environment.SystemDirectory, "cmd.exe"), exe);
            }

            var arguments = "/d /c ping -n 60 127.0.0.1 >nul" + (comLaunched ? " & rem /Automation -Embedding" : string.Empty);
            var process = Process.Start(new ProcessStartInfo(exe, arguments) { UseShellExecute = false, CreateNoWindow = true })!;
            return new StandInWord(process);
        }

        public void Dispose()
        {
            try
            {
                if (!Process.HasExited)
                {
                    Process.Kill(entireProcessTree: true);
                }

                Process.WaitForExit(5000);
            }
            catch (InvalidOperationException)
            {
                // Already gone.
            }
            finally
            {
                Process.Dispose();
            }
        }
    }
}

/// <summary>The pure half of <see cref="WordProcessIdentity"/>, which runs on every OS.</summary>
public class WordProcessIdentityTests
{
    [Theory]
    [InlineData("\"C:\\Program Files\\Microsoft Office\\Root\\Office16\\WINWORD.EXE\" /Automation -Embedding", true)]
    [InlineData("\"C:\\Program Files\\Microsoft Office\\Root\\Office16\\WINWORD.EXE\" -Embedding", true)]
    [InlineData("WINWORD.EXE /embedding", true)]
    [InlineData("\"C:\\Program Files\\Microsoft Office\\Root\\Office16\\WINWORD.EXE\"", false)]
    [InlineData("\"C:\\Program Files\\Microsoft Office\\Root\\Office16\\WINWORD.EXE\" /n \"C:\\docs\\Embedding notes.docx\"", false)]
    [InlineData("\"C:\\Program Files\\Microsoft Office\\Root\\Office16\\WINWORD.EXE\" \"C:\\docs\\x-Embedding.docx\"", false)]
    [InlineData(null, false)]
    public void IsComLaunched_recognizes_only_COMs_Embedding_switch(string? commandLine, bool expected) =>
        Assert.Equal(expected, WordProcessIdentity.IsComLaunched(commandLine));
}

/// <summary>
/// Stands in for <c>Word.Application</c>: the converter drives it through the same late-bound calls it makes on
/// the COM object. It owns no window, so the converter can never confirm its identity and must not kill
/// anything on its behalf. <c>Documents.Open</c> optionally blocks on a gate to simulate a hung Word.
/// </summary>
public sealed class FakeWordApplication
{
    private int _quitCount;

    public FakeWordApplication(ManualResetEventSlim? openGate = null) => Documents = new FakeWordDocuments(this, openGate);

    public string Caption { get; set; } = "Word";

    public bool Visible { get; set; } = true;

    public int DisplayAlerts { get; set; }

    public FakeWordDocuments Documents { get; }

    public int QuitCount => Volatile.Read(ref _quitCount);

    public bool ExportCalled { get; set; }

    public bool DocumentClosed { get; set; }

    public ManualResetEventSlim Quitted { get; } = new(false);

    public ManualResetEventSlim OpenEntered { get; } = new(false);

    public void Quit(int saveChanges)
    {
        _ = saveChanges;
        Interlocked.Increment(ref _quitCount);
        Quitted.Set();
    }
}

/// <summary>See <see cref="FakeWordApplication"/>.</summary>
public sealed class FakeWordDocuments(FakeWordApplication application, ManualResetEventSlim? openGate)
{
    public FakeWordDocument Open(string path)
    {
        _ = path;
        application.OpenEntered.Set();
        openGate?.Wait(TimeSpan.FromSeconds(30));
        return new FakeWordDocument(application);
    }
}

/// <summary>See <see cref="FakeWordApplication"/>.</summary>
public sealed class FakeWordDocument(FakeWordApplication application)
{
    public void ExportAsFixedFormat(string outputPath, int format)
    {
        _ = format;
        application.ExportCalled = true;
        File.WriteAllText(outputPath, "%PDF-1.7 written by FakeWordDocument");
    }

    public void Close(int saveChanges)
    {
        _ = saveChanges;
        application.DocumentClosed = true;
    }
}
