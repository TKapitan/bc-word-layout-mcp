using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Domain;

/// <summary>
/// Gives a layout the stock Office theme part, for <c>create_layout</c>'s <c>theme: "office"</c> (GitHub issue #150).
/// A blank build ships styles but no theme, so it had no colour scheme for a theme colour (<c>accent1</c>, #145) to
/// resolve against, and nothing for <c>set_layout_theme</c> to edit.
/// </summary>
/// <remarks>
/// <para>
/// THE PART is <c>StandardSalesQuote.docx</c>'s <c>word/theme/theme1.xml</c>, copied unchanged (embedded as
/// <c>Resources/OfficeTheme.xml</c>; the SDK may move a namespace declaration when it saves the package): the
/// Office theme of Office 2013 — colour scheme <c>Office</c> (<c>accent1</c>
/// <c>5B9BD5</c>), fonts Calibri Light and Calibri — which five stock corpus layouts share. Its body font is the one
/// <see cref="DefaultStylesScaffold"/> pins, so the theme changes nothing a blank build already prints.
/// </para>
/// <para>
/// THE MAPPING. Every corpus layout's settings part carries the default <c>w:clrSchemeMapping</c> (<c>t1</c> to
/// <c>dark1</c> and so on); one is added beside the theme when the settings part has none. Word applies the same
/// mapping when it is absent, so it changes no colour, but a layout then looks like every stock one.
/// </para>
/// </remarks>
public static class ThemeScaffold
{
    /// <summary>The one theme <c>create_layout</c> can add.</summary>
    public const string Office = "office";

    private const string ResourceName = "BcWordLayout.Domain.OfficeTheme.xml";

    /// <summary>True when <paramref name="theme"/> names a theme <see cref="AddOfficeTheme"/> can add (case-insensitive).</summary>
    public static bool IsKnown(string? theme) => string.Equals(theme?.Trim(), Office, StringComparison.OrdinalIgnoreCase);

    /// <summary>Adds the Office theme part to <paramref name="main"/>, which must have none.</summary>
    /// <exception cref="InvalidOperationException">The layout already has a theme part.</exception>
    public static void AddOfficeTheme(MainDocumentPart main)
    {
        ArgumentNullException.ThrowIfNull(main);

        if (main.ThemePart is not null)
        {
            throw new InvalidOperationException("The layout already has a theme part.");
        }

        var part = main.AddNewPart<ThemePart>();
        using (var stream = typeof(ThemeScaffold).Assembly.GetManifestResourceStream(ResourceName)
                            ?? throw new InvalidOperationException($"Embedded resource '{ResourceName}' is missing."))
        {
            part.FeedData(stream);
        }

        if (main.DocumentSettingsPart?.Settings is { } settings && settings.GetFirstChild<ColorSchemeMapping>() is null)
        {
            settings.AddChild(
                new ColorSchemeMapping
                {
                    Background1 = ColorSchemeIndexValues.Light1,
                    Text1 = ColorSchemeIndexValues.Dark1,
                    Background2 = ColorSchemeIndexValues.Light2,
                    Text2 = ColorSchemeIndexValues.Dark2,
                    Accent1 = ColorSchemeIndexValues.Accent1,
                    Accent2 = ColorSchemeIndexValues.Accent2,
                    Accent3 = ColorSchemeIndexValues.Accent3,
                    Accent4 = ColorSchemeIndexValues.Accent4,
                    Accent5 = ColorSchemeIndexValues.Accent5,
                    Accent6 = ColorSchemeIndexValues.Accent6,
                    Hyperlink = ColorSchemeIndexValues.Hyperlink,
                    FollowedHyperlink = ColorSchemeIndexValues.FollowedHyperlink,
                },
                throwOnError: false);
            settings.Save();
        }
    }
}
