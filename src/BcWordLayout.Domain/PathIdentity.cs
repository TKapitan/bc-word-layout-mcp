using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Microsoft.Win32.SafeHandles;

namespace BcWordLayout.Domain;

/// <summary>
/// Resolves a caller's path to the file system's own spelling of it, so that two spellings of ONE file are
/// recognized as one file. The per-path edit locks (<c>ToolGuards.EditLockFor</c> and <c>CrossProcessLock</c>
/// in <c>BcWordLayout.McpHost</c>), the preview folder key, the commit target of every in-place edit, and the
/// "is this the same file" refusals (<see cref="LayoutBuilder"/>'s outputPath-over-schemaSource check,
/// <c>MergeEngine.Merge</c>'s outputPath-vs-layoutPath check) all go through <see cref="Canonical"/>.
/// </summary>
/// <remarks>
/// <para>
/// WHY <see cref="Path.GetFullPath(string)"/> IS NOT ENOUGH. It resolves relative segments and <c>..</c> and
/// expands the 8.3 short names of existing folders, and the callers compare its result case-insensitively,
/// but it never follows a reparse point or a drive alias: a junction or symlink anywhere along the path, a
/// subst drive, and a mapped drive versus its UNC path all stay distinct spellings. Two host processes
/// editing one layout through two such spellings took two
/// different locks, so neither serialized against the other and one host's commit silently replaced the
/// other's: 12 concurrent <c>insert_text</c> calls, half through the real folder and half through a junction
/// to it, all returned <c>ok:true</c>, and the file kept 6 of the edits (GitHub issue #92).
/// </para>
/// <para>
/// HOW. On Windows the path is opened with NO access rights (so neither Word holding the file open nor a
/// read-only attribute gets in the way) and handed to <c>GetFinalPathNameByHandle</c>, which returns the path
/// the I/O manager actually opened: every junction and symlink followed, a subst or mapped drive replaced by
/// the real volume or UNC share, and every segment in its long form. A path that does not exist yet (a
/// <c>create_layout</c> outputPath) resolves its nearest existing ancestor folder and appends the remaining
/// segments unchanged.
/// </para>
/// <para>
/// WHY A PATH, NOT A FILE ID. The volume serial number plus file ID would identify the file too, but every
/// committed edit REPLACES the layout with a new file (the staged copy is renamed over it), so the ID changes
/// on every edit. A host that read the ID before another host's commit would key a different lock than one
/// that read it after. The resolved path survives the rename.
/// </para>
/// <para>
/// FALLBACK. Whatever stops the resolution (another OS, a folder the account cannot open, an unexpected API
/// failure) falls back to <see cref="Path.GetFullPath(string)"/>, the behavior before this type existed. The
/// method throws only what <see cref="Path.GetFullPath(string)"/> itself throws for a malformed path, so a
/// caller's existing argument checks keep their wording. Off Windows nothing is resolved: the Linux build is a
/// rot guard, not a supported host (see <c>ci.yml</c>).
/// </para>
/// </remarks>
public static class PathIdentity
{
    /// <summary>
    /// The file system's own full spelling of <paramref name="path"/>, or <see cref="Path.GetFullPath(string)"/>
    /// when it cannot be resolved (see the remarks on <see cref="PathIdentity"/>). Two spellings of one file
    /// give the same result up to letter case, so compare results case-insensitively, as every caller does.
    /// </summary>
    public static string Canonical(string path)
    {
        var fullPath = Path.GetFullPath(path);
        if (!OperatingSystem.IsWindows())
        {
            return fullPath;
        }

        try
        {
            return ResolveOnWindows(fullPath) ?? fullPath;
        }
        catch (Exception)
        {
            // Deliberately broad: resolution is an improvement over GetFullPath, never a new way to fail. Any
            // exception here (a P/Invoke surprise, a path shape the API rejects) keeps the previous behavior.
            return fullPath;
        }
    }

    /// <summary>
    /// True when <paramref name="path"/> and <paramref name="otherPath"/> name the same file, compared
    /// case-insensitively after <see cref="Canonical"/>.
    /// </summary>
    /// <remarks>
    /// Case-insensitive on every OS: Windows (the supported platform) and macOS file systems ignore case,
    /// and on a case-sensitive one this can only over-refuse, which costs the caller a different path, never
    /// a destroyed file.
    /// </remarks>
    public static bool SameFile(string path, string otherPath) =>
        string.Equals(Canonical(path), Canonical(otherPath), StringComparison.OrdinalIgnoreCase);

    [SupportedOSPlatform("windows")]
    private static string? ResolveOnWindows(string fullPath)
    {
        // Walk up to the nearest entry that exists, keeping the segments below it that do not exist yet.
        var existing = fullPath;
        var pending = new Stack<string>();
        while (!File.Exists(existing) && !Directory.Exists(existing))
        {
            var parent = Path.GetDirectoryName(existing);
            if (parent is null)
            {
                return null; // Even the root is missing (e.g. an unmapped drive letter): nothing to resolve.
            }

            pending.Push(Path.GetFileName(existing));
            existing = parent;
        }

        var resolved = FinalPathName(existing);
        if (resolved is null)
        {
            return null;
        }

        while (pending.Count > 0)
        {
            resolved = Path.Combine(resolved, pending.Pop());
        }

        return resolved;
    }

    [SupportedOSPlatform("windows")]
    private static string? FinalPathName(string path)
    {
        // Zero desired access plus FILE_FLAG_BACKUP_SEMANTICS opens files and folders alike, even when another
        // program holds the file without sharing it (no read/write/delete access is requested to conflict).
        using var handle = CreateFile(
            path, 0, FileShare.ReadWrite | FileShare.Delete, IntPtr.Zero, FileMode.Open, FileFlagBackupSemantics, IntPtr.Zero);
        if (handle.IsInvalid)
        {
            return null;
        }

        var buffer = new char[512];
        while (true)
        {
            // FILE_NAME_NORMALIZED | VOLUME_NAME_DOS (both 0): the long-form path on its drive letter or UNC share.
            var length = GetFinalPathNameByHandle(handle, buffer, (uint)buffer.Length, 0);
            if (length == 0)
            {
                return null;
            }

            if (length < buffer.Length)
            {
                return StripDevicePrefix(new string(buffer, 0, (int)length));
            }

            // Too small: length is the size needed, terminating null included.
            buffer = new char[length];
        }
    }

    /// <summary>
    /// Turns the API's device-path form back into an ordinary one: <c>\\?\C:\x</c> into <c>C:\x</c>, and
    /// <c>\\?\UNC\server\share\x</c> into <c>\\server\share\x</c>.
    /// </summary>
    private static string StripDevicePrefix(string path)
    {
        const string UncPrefix = @"\\?\UNC\";
        const string DevicePrefix = @"\\?\";
        if (path.StartsWith(UncPrefix, StringComparison.OrdinalIgnoreCase))
        {
            return @"\\" + path[UncPrefix.Length..];
        }

        return path.StartsWith(DevicePrefix, StringComparison.Ordinal) ? path[DevicePrefix.Length..] : path;
    }

    private const uint FileFlagBackupSemantics = 0x02000000;

    [DllImport("kernel32.dll", EntryPoint = "CreateFileW", CharSet = CharSet.Unicode, SetLastError = true, ExactSpelling = true)]
    [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
    private static extern SafeFileHandle CreateFile(
        string fileName, uint desiredAccess, FileShare shareMode, IntPtr securityAttributes,
        FileMode creationDisposition, uint flagsAndAttributes, IntPtr templateFile);

    [DllImport("kernel32.dll", EntryPoint = "GetFinalPathNameByHandleW", CharSet = CharSet.Unicode, SetLastError = true, ExactSpelling = true)]
    [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
    private static extern uint GetFinalPathNameByHandle(SafeFileHandle file, [Out] char[] filePath, uint filePathLength, uint flags);
}
