using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text.RegularExpressions;
using Microsoft.Win32.SafeHandles;

namespace BcWordLayout.Render;

/// <summary>
/// The Win32 lookups <see cref="WordComConverter"/> uses to tell ITS Word instance apart from every other
/// WINWORD.EXE on the machine (GitHub issue #103): the process that owns a Word main window with a given
/// caption, and whether a process was launched by COM. Each lookup returns null or false when it cannot tell;
/// the caller treats that as "not confirmed" and never kills on it.
/// </summary>
internal static partial class WordProcessIdentity
{
    /// <summary>The window class of Word's main window, hidden or not.</summary>
    private const string WordMainWindowClass = "OpusApp";

    /// <summary>
    /// COM starts an out-of-process server with <c>-Embedding</c> on its command line (Word also gets
    /// <c>/Automation</c>), while a Word the user starts never carries it. Observed on Microsoft 365 Word:
    /// <c>"…\WINWORD.EXE" /Automation -Embedding</c>, parented by the DCOM launcher.
    /// </summary>
    [GeneratedRegex(@"(?:^|\s)[-/]Embedding(?:\s|$)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex EmbeddingSwitch();

    /// <summary>True when <paramref name="commandLine"/> carries COM's <c>-Embedding</c> switch.</summary>
    internal static bool IsComLaunched(string? commandLine) => commandLine is not null && EmbeddingSwitch().IsMatch(commandLine);

    /// <summary>
    /// The id of the process owning a Word main window whose title contains <paramref name="captionToken"/>,
    /// polled until <paramref name="timeout"/> (Word retitles its window as soon as <c>Caption</c> is set, so
    /// the first look normally finds it), or null when there is none.
    /// </summary>
    [SupportedOSPlatform("windows")]
    internal static int? FindWordWindowProcessId(string captionToken, TimeSpan timeout)
    {
        var deadline = Stopwatch.StartNew();
        while (true)
        {
            var window = IntPtr.Zero;
            while ((window = FindWindowEx(IntPtr.Zero, window, WordMainWindowClass, null)) != IntPtr.Zero)
            {
                if (WindowText(window).Contains(captionToken, StringComparison.Ordinal)
                    && GetWindowThreadProcessId(window, out var processId) != 0
                    && processId != 0)
                {
                    return (int)processId;
                }
            }

            if (deadline.Elapsed >= timeout)
            {
                return null;
            }

            Thread.Sleep(50);
        }
    }

    /// <summary>
    /// The command line <paramref name="process"/> was started with, or null when it cannot be read (the
    /// process exited, or this account may not query it).
    /// </summary>
    [SupportedOSPlatform("windows")]
    internal static string? TryReadCommandLine(Process process)
    {
        try
        {
            var size = 2048;
            for (var attempt = 0; attempt < 4; attempt++)
            {
                var buffer = Marshal.AllocHGlobal(size);
                try
                {
                    var status = NtQueryInformationProcess(
                        process.SafeHandle, ProcessCommandLineInformation, buffer, size, out var needed);
                    if (status is StatusInfoLengthMismatch or StatusBufferTooSmall or StatusBufferOverflow)
                    {
                        size = Math.Max(needed, size * 2);
                        continue;
                    }

                    if (status != 0)
                    {
                        return null;
                    }

                    // The buffer starts with a UNICODE_STRING { USHORT Length; USHORT MaximumLength; PWSTR Buffer; }
                    // whose Buffer points into the same allocation; Buffer sits at pointer alignment.
                    var lengthInBytes = (ushort)Marshal.ReadInt16(buffer);
                    var text = Marshal.ReadIntPtr(buffer, IntPtr.Size);
                    return text == IntPtr.Zero ? null : Marshal.PtrToStringUni(text, lengthInBytes / 2);
                }
                finally
                {
                    Marshal.FreeHGlobal(buffer);
                }
            }

            return null;
        }
        catch (Exception)
        {
            // Deliberately broad: an unreadable command line only means "not confirmed as COM-launched", and
            // the caller then never kills the process. Access denied, an exited process (InvalidOperation /
            // Win32Exception from SafeHandle) and marshalling surprises all land here.
            return null;
        }
    }

    private static string WindowText(IntPtr window)
    {
        var length = GetWindowTextLength(window);
        if (length <= 0)
        {
            return string.Empty;
        }

        var buffer = new char[length + 1];
        var copied = GetWindowText(window, buffer, buffer.Length);
        return copied <= 0 ? string.Empty : new string(buffer, 0, copied);
    }

    private const int ProcessCommandLineInformation = 60;
    private const int StatusInfoLengthMismatch = unchecked((int)0xC0000004);
    private const int StatusBufferTooSmall = unchecked((int)0xC0000023);
    private const int StatusBufferOverflow = unchecked((int)0x80000005);

    [DllImport("ntdll.dll", ExactSpelling = true)]
    [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
    private static extern int NtQueryInformationProcess(
        SafeProcessHandle processHandle, int processInformationClass, IntPtr processInformation,
        int processInformationLength, out int returnLength);

    [DllImport("user32.dll", EntryPoint = "FindWindowExW", CharSet = CharSet.Unicode, ExactSpelling = true)]
    [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
    private static extern IntPtr FindWindowEx(IntPtr parent, IntPtr childAfter, string? className, string? windowName);

    [DllImport("user32.dll", ExactSpelling = true)]
    [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
    private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

    [DllImport("user32.dll", EntryPoint = "GetWindowTextLengthW", CharSet = CharSet.Unicode, ExactSpelling = true)]
    [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
    private static extern int GetWindowTextLength(IntPtr window);

    [DllImport("user32.dll", EntryPoint = "GetWindowTextW", CharSet = CharSet.Unicode, ExactSpelling = true)]
    [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
    private static extern int GetWindowText(IntPtr window, [Out] char[] text, int maxCount);
}
