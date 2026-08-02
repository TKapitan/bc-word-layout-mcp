using System.Diagnostics;

namespace BcWordLayout.Render;

/// <summary>
/// Converts .docx to PDF by shelling out to LibreOffice headless. Cross-platform and process-based — the
/// fallback converter for machines without Word (chiefly CI), selected by <see cref="PdfConverterFactory"/>
/// when <see cref="WordComConverter"/> is unavailable or not preferred. Every failure mode (missing input,
/// missing <c>soffice</c>, timeout, nonzero exit, no/invalid output) is reported as a structured
/// <see cref="PdfConversionResult"/> with <c>Ok = false</c>; this converter never throws.
/// </summary>
public sealed class LibreOfficeConverter : IPdfConverter
{
    /// <inheritdoc/>
    public string Name => "libreoffice";

    /// <inheritdoc/>
    public bool IsAvailable => LibreOfficeCli.FindSoffice() is not null;

    /// <inheritdoc/>
    public PdfConversionResult Convert(string docxPath, string pdfPath, PdfConversionOptions? options = null)
    {
        var sw = Stopwatch.StartNew();
        options ??= new PdfConversionOptions();

        if (!File.Exists(docxPath))
        {
            return PdfConversionResult.Failure(Name, $"Input .docx not found: '{docxPath}'.", sw.Elapsed);
        }

        var soffice = LibreOfficeCli.FindSoffice();
        if (soffice is null)
        {
            return PdfConversionResult.Failure(
                Name,
                "LibreOffice ('soffice') was not found on PATH or in the usual install locations.",
                sw.Elapsed);
        }

        if (File.Exists(pdfPath) && !options.Overwrite)
        {
            return PdfConversionResult.Failure(
                Name, $"Destination already exists and Overwrite is false: '{pdfPath}'.", sw.Elapsed);
        }

        var fullDocxPath = Path.GetFullPath(docxPath);
        var outDir = Path.Combine(Path.GetTempPath(), $"bcwl-lo-out-{Guid.NewGuid():N}");
        var profileDir = Path.Combine(Path.GetTempPath(), $"bcwl-lo-profile-{Guid.NewGuid():N}");
        Process? process = null;

        try
        {
            Directory.CreateDirectory(outDir);
            Directory.CreateDirectory(profileDir);

            var psi = new ProcessStartInfo(soffice)
            {
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
            };
            foreach (var arg in LibreOfficeCli.BuildConvertArgs(fullDocxPath, outDir, profileDir))
            {
                psi.ArgumentList.Add(arg);
            }

            process = Process.Start(psi);
            if (process is null)
            {
                return PdfConversionResult.Failure(Name, "Failed to start the LibreOffice ('soffice') process.", sw.Elapsed);
            }

            // Drain both redirected streams asynchronously so a full pipe buffer can never deadlock the
            // wait below; only actually consumed (awaited) once we know the process has exited.
            var stdErrTask = process.StandardError.ReadToEndAsync();
            var stdOutTask = process.StandardOutput.ReadToEndAsync();

            var waitMs = (int)Math.Min(options.Timeout.TotalMilliseconds, int.MaxValue);
            if (!process.WaitForExit(waitMs))
            {
                KillIfRunning(process);

                // The kill can make either redirected-stream read fault (e.g. the pipe breaking abruptly).
                // Nobody awaits these after a timeout, so observe (and ignore) any exception now rather
                // than risk it surfacing later as an unobserved task exception.
                ObserveAndIgnore(stdOutTask);
                ObserveAndIgnore(stdErrTask);

                return PdfConversionResult.Failure(
                    Name, $"LibreOffice conversion timed out after {options.Timeout}.", sw.Elapsed);
            }

            // The process already exited; this only lets the async reads settle before we touch .Result.
            process.WaitForExit();
            var stdErr = stdErrTask.GetAwaiter().GetResult();
            _ = stdOutTask.GetAwaiter().GetResult();

            if (process.ExitCode != 0)
            {
                return PdfConversionResult.Failure(
                    Name, $"LibreOffice exited with code {process.ExitCode}. {stdErr}".Trim(), sw.Elapsed);
            }

            // soffice --convert-to only accepts --outdir; it names the output after the input's own
            // basename, so the produced file must be located rather than assumed to be at pdfPath.
            var producedPdf = Path.Combine(outDir, Path.GetFileNameWithoutExtension(fullDocxPath) + ".pdf");
            if (!File.Exists(producedPdf))
            {
                return PdfConversionResult.Failure(
                    Name,
                    $"LibreOffice reported success but produced no PDF at '{producedPdf}'. {stdErr}".Trim(),
                    sw.Elapsed);
            }

            if (!PdfFileValidation.LooksLikePdf(producedPdf))
            {
                return PdfConversionResult.Failure(Name, "LibreOffice produced a file that is not a valid PDF.", sw.Elapsed);
            }

            var destDir = Path.GetDirectoryName(Path.GetFullPath(pdfPath));
            if (!string.IsNullOrEmpty(destDir))
            {
                Directory.CreateDirectory(destDir);
            }

            File.Copy(producedPdf, pdfPath, overwrite: options.Overwrite);

            return PdfConversionResult.Success(Name, pdfPath, sw.Elapsed);
        }
        catch (Exception ex)
        {
            // Boundary contract: this method never throws. Anything unexpected (process start failure,
            // locked file, etc.) becomes a structured failure instead.
            return PdfConversionResult.Failure(Name, $"LibreOffice conversion failed: {ex.Message}", sw.Elapsed);
        }
        finally
        {
            process?.Dispose();
            TryDeleteDirectory(outDir);
            TryDeleteDirectory(profileDir);
        }
    }

    private static void KillIfRunning(Process process)
    {
        try
        {
            process.Refresh();
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

    /// <summary>
    /// Attaches a continuation that touches <see cref="Task.Exception"/> (marking it observed) without
    /// blocking the caller — used after a timeout-kill, where nobody will ever await
    /// <paramref name="task"/> again, so a later fault (e.g. a redirected pipe breaking) must not become an
    /// unobserved task exception.
    /// </summary>
    private static void ObserveAndIgnore(Task task) => task.ContinueWith(t => _ = t.Exception, TaskScheduler.Default);

    private static void TryDeleteDirectory(string path)
    {
        try
        {
            if (Directory.Exists(path))
            {
                Directory.Delete(path, recursive: true);
            }
        }
        catch (IOException)
        {
            // Best-effort cleanup only; a locked temp file must not fail the conversion result.
        }
        catch (UnauthorizedAccessException)
        {
            // Best-effort cleanup only.
        }
    }
}

/// <summary>
/// Pure, dependency-free pieces of <see cref="LibreOfficeConverter"/>'s process invocation, split out so
/// tests can assert on exact candidate paths / argument lists without spawning a real <c>soffice</c>
/// process or depending on what happens to be installed on the test machine.
/// </summary>
internal static class LibreOfficeCli
{
    /// <summary>
    /// Locates the <c>soffice</c> executable: the well-known per-OS install locations first, then every
    /// directory on PATH (see <see cref="CandidatePaths"/>). Returns the first candidate that exists on
    /// disk, or null if none do.
    /// </summary>
    internal static string? FindSoffice()
    {
        var isWindows = OperatingSystem.IsWindows();
        var pathVar = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
        var directories = pathVar.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries);

        foreach (var candidate in CandidatePaths(isWindows, directories))
        {
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        return null;
    }

    /// <summary>
    /// Builds the ordered list of paths worth probing for the <c>soffice</c> executable on
    /// <paramref name="isWindows"/>: the well-known default install locations for that OS first, then
    /// every directory in <paramref name="pathDirectories"/> last. Pure — takes the OS flag and PATH
    /// directories as plain parameters instead of reading the live environment, so it is unit-testable
    /// with fixed inputs. Joins path segments by hand with the separator matching
    /// <paramref name="isWindows"/> rather than via <see cref="Path.Combine(string, string)"/>, which is
    /// tied to the actual host OS's separator instead of this parameter — needed so a test can
    /// deterministically ask "what would the Linux candidate list look like" while itself running on
    /// Windows (or vice versa).
    /// </summary>
    /// <remarks>
    /// Ordering rationale (trusted locations before PATH —
    /// a PATH-order hijack): probing every PATH directory first meant a <c>soffice</c>/
    /// <c>soffice.exe</c> planted earlier on PATH than the real install would be resolved and executed
    /// instead of the genuine LibreOffice binary. Checking the fixed, admin-writable install directories
    /// first closes that hijack window; PATH is now only consulted as a fallback for non-standard install
    /// locations (package-manager installs, portable builds, etc.).
    /// </remarks>
    internal static IReadOnlyList<string> CandidatePaths(bool isWindows, IEnumerable<string> pathDirectories)
    {
        var candidates = new List<string>();
        var separator = isWindows ? '\\' : '/';

        string JoinPath(string dir, string fileName)
        {
            var trimmed = dir.TrimEnd('\\', '/');
            return trimmed.Length == 0 ? fileName : $"{trimmed}{separator}{fileName}";
        }

        if (isWindows)
        {
            // .com before .exe: on Windows, LibreOffice's soffice.exe (GUI subsystem) can hand control
            // back to the caller before the conversion has actually finished writing the output file,
            // while soffice.com (console subsystem) blocks until the whole operation truly completes —
            // preferring it avoids a "reported success but produced no PDF yet" race.
            candidates.Add(@"C:\Program Files\LibreOffice\program\soffice.com");
            candidates.Add(@"C:\Program Files\LibreOffice\program\soffice.exe");
            candidates.Add(@"C:\Program Files (x86)\LibreOffice\program\soffice.com");
            candidates.Add(@"C:\Program Files (x86)\LibreOffice\program\soffice.exe");

            foreach (var dir in pathDirectories)
            {
                candidates.Add(JoinPath(dir, "soffice.com"));
                candidates.Add(JoinPath(dir, "soffice.exe"));
            }
        }
        else
        {
            candidates.Add("/usr/bin/soffice");
            candidates.Add("/usr/bin/libreoffice");
            candidates.Add("/opt/libreoffice/program/soffice");
            candidates.Add("/Applications/LibreOffice.app/Contents/MacOS/soffice");

            foreach (var dir in pathDirectories)
            {
                candidates.Add(JoinPath(dir, "soffice"));
                candidates.Add(JoinPath(dir, "libreoffice"));
            }
        }

        return candidates;
    }

    /// <summary>
    /// Builds the argument list for <c>soffice --headless --convert-to pdf ...</c> against
    /// <paramref name="docxPath"/>, writing into <paramref name="outDir"/> under a unique
    /// <paramref name="userInstallDir"/> profile (via <c>-env:UserInstallation</c>) so the conversion works
    /// even while a desktop LibreOffice instance is already running under the same user account. Pure — no
    /// process is started; callers pass each entry through <see cref="ProcessStartInfo.ArgumentList"/> so
    /// no shell quoting is needed even when paths contain spaces.
    /// </summary>
    internal static IReadOnlyList<string> BuildConvertArgs(string docxPath, string outDir, string userInstallDir)
    {
        var profileUri = new Uri(Path.GetFullPath(userInstallDir)).AbsoluteUri;

        return new[]
        {
            "--headless",
            "--norestore",
            "--nolockcheck",
            "--nodefault",
            $"-env:UserInstallation={profileUri}",
            "--convert-to",
            "pdf:writer_pdf_Export",
            "--outdir",
            outDir,
            docxPath,
        };
    }
}
