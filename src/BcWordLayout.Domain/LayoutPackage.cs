using DocumentFormat.OpenXml.Packaging;

namespace BcWordLayout.Domain;

/// <summary>
/// The ONE way this codebase opens a <c>.docx</c> from a path: every open of an untrusted layout, template,
/// schema source, edit working copy and merge staging copy goes through <see cref="Open"/>, which checks the
/// package against <see cref="ResourceLimits.EnsurePackageWithinLimits"/> before the OpenXml SDK parses any
/// part. A crafted file therefore fails the call with a <see cref="ResourceLimitExceededException"/> instead
/// of overflowing the stack inside the SDK's own recursive load and killing the server (an uncatchable
/// <see cref="StackOverflowException"/>). <c>PackageOpenGatewayTests</c> fails the build if any source file
/// calls <c>WordprocessingDocument.Open</c> directly, so the gate holds structurally rather than by habit.
/// </summary>
public static class LayoutPackage
{
    /// <summary>
    /// Opens <paramref name="path"/> after checking it against the package-wide resource caps.
    /// <paramref name="autoSave"/> maps to <see cref="OpenSettings.AutoSave"/> (the SDK's default is true).
    /// </summary>
    /// <remarks>
    /// A writable open is checked through a separate READ-ONLY open first: in a writable package,
    /// System.IO.Packaging decompresses a whole part into memory as soon as its stream is opened, so a
    /// compressed "bomb" part would be inflated before any byte cap could apply. The extra open only parses
    /// the package's relationships, a few milliseconds for a real layout.
    /// <see cref="OpenSettings.MaxCharactersInPart"/> is set as a second, SDK-side ceiling.
    /// </remarks>
    public static WordprocessingDocument Open(string path, bool isEditable, bool autoSave = true)
    {
        var settings = new OpenSettings
        {
            AutoSave = autoSave,
            MaxCharactersInPart = ResourceLimits.MaxPackageXmlPartBytes,
        };

        if (isEditable)
        {
            using (var probe = WordprocessingDocument.Open(path, false, settings))
            {
                ResourceLimits.EnsurePackageWithinLimits(probe);
            }

            return WordprocessingDocument.Open(path, true, settings);
        }

        var doc = WordprocessingDocument.Open(path, false, settings);
        try
        {
            ResourceLimits.EnsurePackageWithinLimits(doc);
            return doc;
        }
        catch
        {
            doc.Dispose();
            throw;
        }
    }
}
