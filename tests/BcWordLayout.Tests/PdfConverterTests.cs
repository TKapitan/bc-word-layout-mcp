using System.Text;
using BcWordLayout.Merge;
using BcWordLayout.Render;

namespace BcWordLayout.Tests;

/// <summary>
/// Discovery/selection tests for <see cref="PdfConverterFactory"/>. These never touch a real Word or
/// LibreOffice install, so they always run regardless of what is installed on the test machine.
/// </summary>
public class PdfConverterFactoryTests
{
    [Fact]
    public void All_includes_libreoffice_always_and_word_only_on_windows()
    {
        var all = PdfConverterFactory.All();

        Assert.Contains(all, c => c.Name == "libreoffice");
        Assert.Equal(OperatingSystem.IsWindows(), all.Any(c => c.Name == "word-com"));
    }

    [Fact]
    public void AnyAvailable_matches_whether_any_converter_in_All_reports_available()
    {
        Assert.Equal(PdfConverterFactory.All().Any(c => c.IsAvailable), PdfConverterFactory.AnyAvailable);
    }

    [Fact]
    public void Select_Auto_returns_a_converter_whose_availability_matches_AnyAvailable()
    {
        var selected = PdfConverterFactory.Select(PdfConverterKind.Auto);

        // Auto must resolve to SOME available converter whenever one exists, and (since Word is preferred)
        // never claim availability when nothing on the machine actually is.
        Assert.Equal(PdfConverterFactory.AnyAvailable, selected.IsAvailable);
    }

    [Fact]
    public void Select_Auto_prefers_Word_when_it_is_available()
    {
        var wordDirectlyAvailable = OperatingSystem.IsWindows() && new WordComConverter().IsAvailable;
        var selected = PdfConverterFactory.Select(PdfConverterKind.Auto);

        Assert.Equal(wordDirectlyAvailable, selected.Name == "word-com");
    }

    [Theory]
    [InlineData(PdfConverterKind.Word, "word-com")]
    [InlineData(PdfConverterKind.LibreOffice, "libreoffice")]
    public void Select_specific_kind_returns_a_converter_with_the_expected_name(PdfConverterKind kind, string expectedName)
    {
        var selected = PdfConverterFactory.Select(kind);
        Assert.Equal(expectedName, selected.Name);
    }

    [Fact]
    public void Select_Word_on_non_windows_returns_a_converter_that_fails_cleanly_instead_of_throwing()
    {
        if (OperatingSystem.IsWindows())
        {
            return; // Nothing to prove here; Select(Word) legitimately returns the real WordComConverter.
        }

        var selected = PdfConverterFactory.Select(PdfConverterKind.Word);
        Assert.False(selected.IsAvailable);

        var result = selected.Convert("missing.docx", "missing.pdf");
        Assert.False(result.Ok);
        Assert.False(string.IsNullOrWhiteSpace(result.Error));
    }
}

/// <summary>
/// Tests for the pure (no process spawned) argument/discovery builders behind <see cref="LibreOfficeConverter"/>.
/// </summary>
public class LibreOfficeCliTests
{
    [Fact]
    public void CandidatePaths_on_windows_includes_path_directories_and_well_known_install_locations()
    {
        var candidates = LibreOfficeCli.CandidatePaths(isWindows: true, pathDirectories: new[] { @"C:\tools" });

        Assert.Contains(@"C:\tools\soffice.exe", candidates);
        Assert.Contains(@"C:\tools\soffice.com", candidates);
        Assert.Contains(@"C:\Program Files\LibreOffice\program\soffice.exe", candidates);
    }

    [Fact]
    public void CandidatePaths_on_non_windows_includes_path_directories_and_unix_locations()
    {
        var candidates = LibreOfficeCli.CandidatePaths(isWindows: false, pathDirectories: new[] { "/usr/local/bin" });

        Assert.Contains("/usr/local/bin/soffice", candidates);
        Assert.Contains("/usr/local/bin/libreoffice", candidates);
        Assert.Contains("/usr/bin/soffice", candidates);
    }

    [Fact]
    public void CandidatePaths_probes_every_supplied_path_directory_not_just_the_first()
    {
        var candidates = LibreOfficeCli.CandidatePaths(isWindows: true, pathDirectories: new[] { @"C:\a", @"C:\b" });

        Assert.Contains(@"C:\a\soffice.exe", candidates);
        Assert.Contains(@"C:\b\soffice.exe", candidates);
    }

    [Fact]
    public void CandidatePaths_on_windows_prefers_soffice_com_over_soffice_exe_for_each_location()
    {
        // soffice.exe (GUI subsystem) can return before the PDF is fully written; soffice.com (console
        // subsystem) blocks until done — .com must be probed first everywhere a pair of them appears, so
        // FindSoffice() (which returns the FIRST candidate that exists) prefers it when both are installed.
        var candidates = LibreOfficeCli.CandidatePaths(isWindows: true, pathDirectories: new[] { @"C:\tools" }).ToList();

        var pathDirComIndex = candidates.IndexOf(@"C:\tools\soffice.com");
        var pathDirExeIndex = candidates.IndexOf(@"C:\tools\soffice.exe");
        Assert.True(pathDirComIndex >= 0, "expected C:\\tools\\soffice.com to be a candidate");
        Assert.True(pathDirExeIndex >= 0, "expected C:\\tools\\soffice.exe to be a candidate");
        Assert.True(pathDirComIndex < pathDirExeIndex, "expected soffice.com to be probed before soffice.exe");

        var programFilesComIndex = candidates.IndexOf(@"C:\Program Files\LibreOffice\program\soffice.com");
        var programFilesExeIndex = candidates.IndexOf(@"C:\Program Files\LibreOffice\program\soffice.exe");
        Assert.True(programFilesComIndex >= 0 && programFilesExeIndex >= 0);
        Assert.True(programFilesComIndex < programFilesExeIndex, "expected soffice.com to be probed before soffice.exe");
    }

    [Fact]
    public void CandidatePaths_on_windows_probes_trusted_install_locations_before_path_directories()
    {
        // a planted soffice.exe earlier on PATH than the real install must not win. Probing the
        // fixed, admin-writable install directories first (PATH last, as a fallback) closes that hijack
        // window — see the ordering rationale in LibreOfficeCli.CandidatePaths.
        var candidates = LibreOfficeCli.CandidatePaths(isWindows: true, pathDirectories: new[] { @"C:\tools" }).ToList();

        var trustedIndex = candidates.IndexOf(@"C:\Program Files\LibreOffice\program\soffice.com");
        var pathDirIndex = candidates.IndexOf(@"C:\tools\soffice.com");

        Assert.True(trustedIndex >= 0, "expected the trusted Program Files location to be a candidate");
        Assert.True(pathDirIndex >= 0, "expected the PATH directory to be a candidate");
        Assert.True(trustedIndex < pathDirIndex, "expected trusted install locations to be probed before PATH directories");
    }

    [Fact]
    public void CandidatePaths_on_non_windows_probes_trusted_install_locations_before_path_directories()
    {
        // Same ordering rationale as the Windows case, for the non-Windows probe list.
        var candidates = LibreOfficeCli.CandidatePaths(isWindows: false, pathDirectories: new[] { "/usr/local/bin" }).ToList();

        var trustedIndex = candidates.IndexOf("/usr/bin/soffice");
        var pathDirIndex = candidates.IndexOf("/usr/local/bin/soffice");

        Assert.True(trustedIndex >= 0, "expected the trusted /usr/bin location to be a candidate");
        Assert.True(pathDirIndex >= 0, "expected the PATH directory to be a candidate");
        Assert.True(trustedIndex < pathDirIndex, "expected trusted install locations to be probed before PATH directories");
    }

    [Fact]
    public void BuildConvertArgs_produces_the_expected_soffice_arguments()
    {
        var docx = Path.Combine(Path.GetTempPath(), "bcwl-cli-test-input.docx");
        var outDir = Path.Combine(Path.GetTempPath(), "bcwl-cli-test-out");
        var profileDir = Path.Combine(Path.GetTempPath(), "bcwl-cli-test-profile");

        var args = LibreOfficeCli.BuildConvertArgs(docx, outDir, profileDir).ToList();

        Assert.Contains("--headless", args);
        Assert.Contains("--norestore", args);
        Assert.Contains("--nolockcheck", args);
        Assert.Contains("--nodefault", args);
        Assert.Contains("--convert-to", args);
        Assert.Contains("pdf:writer_pdf_Export", args);
        Assert.Contains("--outdir", args);
        Assert.Contains(outDir, args);
        Assert.Contains(docx, args);

        // The convert-to filter is the argument immediately after --convert-to; --outdir's value the
        // argument immediately after it. Order matters to soffice's CLI parser.
        Assert.Equal("pdf:writer_pdf_Export", args[args.IndexOf("--convert-to") + 1]);
        Assert.Equal(outDir, args[args.IndexOf("--outdir") + 1]);

        var envArg = Assert.Single(args, a => a.StartsWith("-env:UserInstallation=", StringComparison.Ordinal));
        Assert.StartsWith("-env:UserInstallation=file:", envArg, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildConvertArgs_userInstallation_uri_has_no_backslashes_even_on_windows_style_paths()
    {
        var args = LibreOfficeCli.BuildConvertArgs(
            docxPath: "in.docx", outDir: "out", userInstallDir: @"C:\Users\someone\AppData\Local\Temp\profile-x");

        var envArg = args.Single(a => a.StartsWith("-env:UserInstallation=", StringComparison.Ordinal));
        Assert.DoesNotContain('\\', envArg);
    }
}

/// <summary>Tests for the shared <c>%PDF</c> output sanity check used by both converters.</summary>
public class PdfFileValidationTests
{
    private static string TempPath() => Path.Combine(Path.GetTempPath(), $"bcwl-pdfcheck-{Guid.NewGuid():N}");

    [Fact]
    public void LooksLikePdf_true_for_a_file_starting_with_the_PDF_magic_number()
    {
        var path = TempPath();
        File.WriteAllBytes(path, Encoding.ASCII.GetBytes("%PDF-1.7\n...rest of a fake pdf..."));

        try
        {
            Assert.True(PdfFileValidation.LooksLikePdf(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void LooksLikePdf_false_for_a_file_without_the_PDF_magic_number()
    {
        var path = TempPath();
        File.WriteAllText(path, "this is not a pdf");

        try
        {
            Assert.False(PdfFileValidation.LooksLikePdf(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void LooksLikePdf_false_for_a_missing_file()
    {
        Assert.False(PdfFileValidation.LooksLikePdf(TempPath()));
    }
}

/// <summary>
/// Behavioral contract tests run against every converter <see cref="PdfConverterFactory"/> knows about.
/// Deliberately dependency-agnostic (no skip framework): whichever branch runs, the test still asserts a
/// specific, correct outcome, so the suite stays fully green whether or not Word/LibreOffice is installed.
/// </summary>
public class PdfConverterContractTests
{
    private static string TempPath(string extension) =>
        Path.Combine(Path.GetTempPath(), $"bcwl-pdfconv-{Guid.NewGuid():N}{extension}");

    [Theory]
    [InlineData(PdfConverterKind.Word)]
    [InlineData(PdfConverterKind.LibreOffice)]
    public void Convert_a_real_merged_layout_succeeds_if_available_else_fails_cleanly(PdfConverterKind kind)
    {
        var converter = PdfConverterFactory.Select(kind);
        var mergedDocx = TempPath(".docx");
        var pdfPath = TempPath(".pdf");

        try
        {
            MergeEngine.Merge(Corpus.Path(Corpus.SalesInvoice), mergedDocx, new MergeOptions { Seed = 12345, Rows = 2 });

            var result = converter.Convert(mergedDocx, pdfPath);

            Assert.Equal(converter.Name, result.Converter);

            if (converter.IsAvailable)
            {
                Assert.True(result.Ok, $"expected a successful conversion via '{converter.Name}' but got: {result.Error}");
                Assert.Equal(pdfPath, result.PdfPath);
                Assert.True(File.Exists(pdfPath), "converter reported success but the PDF file is missing");

                var bytes = File.ReadAllBytes(pdfPath);
                Assert.True(bytes.Length > 1024, $"expected a non-trivial PDF, was {bytes.Length} bytes");
                Assert.Equal("%PDF", Encoding.ASCII.GetString(bytes, 0, 4));
            }
            else
            {
                Assert.False(result.Ok);
                Assert.False(string.IsNullOrWhiteSpace(result.Error));
                Assert.False(File.Exists(pdfPath));
            }
        }
        finally
        {
            File.Delete(mergedDocx);
            File.Delete(pdfPath);
        }
    }

    [Theory]
    [InlineData(PdfConverterKind.Word)]
    [InlineData(PdfConverterKind.LibreOffice)]
    public void Convert_with_missing_input_returns_a_clean_failure_regardless_of_availability(PdfConverterKind kind)
    {
        var converter = PdfConverterFactory.Select(kind);
        var missingDocx = TempPath(".docx");
        var pdfPath = TempPath(".pdf");

        var result = converter.Convert(missingDocx, pdfPath);

        Assert.False(result.Ok);
        Assert.False(string.IsNullOrWhiteSpace(result.Error));
        Assert.Null(result.PdfPath);
        Assert.False(File.Exists(pdfPath));
    }

    [Theory]
    [InlineData(PdfConverterKind.Word)]
    [InlineData(PdfConverterKind.LibreOffice)]
    public void Convert_does_not_overwrite_an_existing_destination_when_Overwrite_is_false(PdfConverterKind kind)
    {
        var converter = PdfConverterFactory.Select(kind);
        var mergedDocx = TempPath(".docx");
        var pdfPath = TempPath(".pdf");
        const string SentinelContent = "not a real pdf yet";

        try
        {
            MergeEngine.Merge(Corpus.Path(Corpus.SalesInvoice), mergedDocx, new MergeOptions { Seed = 1, Rows = 1 });
            File.WriteAllText(pdfPath, SentinelContent);

            var result = converter.Convert(mergedDocx, pdfPath, new PdfConversionOptions { Overwrite = false });

            // Whether the failure came from "already exists" or (on a machine without the dependency) from
            // unavailability, either way Convert must refuse cleanly and must never touch the existing file.
            Assert.False(result.Ok);
            Assert.False(string.IsNullOrWhiteSpace(result.Error));
            Assert.Equal(SentinelContent, File.ReadAllText(pdfPath));
        }
        finally
        {
            File.Delete(mergedDocx);
            File.Delete(pdfPath);
        }
    }

    /// <summary>
    /// Regression test for a hole found in review: WordComConverter's synchronous prologue used to run
    /// outside any try/catch, so <c>Path.GetFullPath</c> on an empty/whitespace <c>pdfPath</c> threw
    /// straight out of <see cref="IPdfConverter.Convert"/> instead of returning a clean failure. Targets
    /// Word specifically (that is where the hole was); on a machine without Word this still passes, just
    /// via the shallower "Word not installed" branch instead of exercising the fixed path-handling code.
    /// </summary>
    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void Convert_with_an_invalid_pdfPath_returns_a_clean_failure_and_never_throws(string invalidPdfPath)
    {
        var converter = PdfConverterFactory.Select(PdfConverterKind.Word);
        var mergedDocx = TempPath(".docx");

        try
        {
            MergeEngine.Merge(Corpus.Path(Corpus.SalesInvoice), mergedDocx, new MergeOptions { Seed = 1, Rows = 1 });

            // Deliberately no try/catch around this call: if Convert ever let an exception escape, this
            // test must fail loudly (an uncaught exception fails the test) rather than mask the very
            // regression it exists to catch.
            var result = converter.Convert(mergedDocx, invalidPdfPath);

            Assert.False(result.Ok);
            Assert.False(string.IsNullOrWhiteSpace(result.Error));
        }
        finally
        {
            File.Delete(mergedDocx);
        }
    }

    [Fact]
    public void Convert_with_a_non_positive_Timeout_returns_a_clean_failure_and_never_throws()
    {
        var converter = PdfConverterFactory.Select(PdfConverterKind.Word);
        var mergedDocx = TempPath(".docx");
        var pdfPath = TempPath(".pdf");

        try
        {
            MergeEngine.Merge(Corpus.Path(Corpus.SalesInvoice), mergedDocx, new MergeOptions { Seed = 1, Rows = 1 });

            var result = converter.Convert(mergedDocx, pdfPath, new PdfConversionOptions { Timeout = TimeSpan.Zero });

            Assert.False(result.Ok);
            Assert.False(string.IsNullOrWhiteSpace(result.Error));
            Assert.False(File.Exists(pdfPath));
        }
        finally
        {
            File.Delete(mergedDocx);
            File.Delete(pdfPath);
        }
    }
}
