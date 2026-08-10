using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Domain;

/// <summary>
/// Creates the <see cref="DocumentSettingsPart"/> a BLANK <see cref="LayoutBuilder.Create"/> output needs
/// before Word will admit that its own repeaters exist. A from-scratch layout used to ship no
/// <c>word/settings.xml</c> at all, and a document that declares no compatibility mode is treated by Word as
/// mode 12 — Word 2007 (measured: <c>Document.CompatibilityMode</c> reports <c>12</c> for a blank build,
/// GitHub issue #51). Repeating-section content controls are a Word 2013 (<c>w15</c>) feature that does not
/// exist in that mode, so the first time a human opened a tool-authored layout in Word and saved it, the
/// Compatibility Checker offered to "convert repeating section content controls to rich text content
/// controls" and continuing through it stripped <c>w15:repeatingSection</c>,
/// <c>w15:repeatingSectionItem</c> AND the repeater's own <c>w15:dataBinding</c> — leaving a plain rich-text
/// control that still looks like a layout but no longer repeats or binds. The scaffold declares the mode
/// every stock corpus layout declares, so a blank build survives the Word round trip that
/// <c>al-word-layout</c> §3 itself recommends for anything the tools cannot style.
/// </summary>
/// <remarks>
/// <para>
/// WHAT IS EMITTED: a settings part holding exactly <c>w:compat</c> with the single
/// <c>compatibilityMode</c> = 15 <c>w:compatSetting</c> — corpus-observed
/// (<c>StandardPurchaseOrder.docx</c> and every other capture carry that element with that value) and
/// nothing else. Word writes a dozen further settings of its own (zoom, proof state, rsids, math
/// properties, theme font language); none of them affect what BC renders or what the tools emit, so per the
/// observed-OOXML-only rule this scaffold stays the minimum that fixes the defect rather than a
/// reconstruction of Word's boilerplate. Word fills the rest in itself the first time it saves.
/// </para>
/// <para>
/// BLANK BUILDS ONLY, AND ONLY AT CREATE TIME — the same contract as
/// <see cref="DefaultStylesScaffold"/>, and for a sharper reason than typography: compatibility mode is not
/// merely a feature gate, it also selects Word's layout metrics (table and text measurement differ between
/// 12 and 15), so retrofitting a mode onto a document that already renders SOMEHOW could silently move its
/// pagination. A <c>templatePath</c> build and any pre-existing layout are therefore left exactly as found,
/// and the risk is REPORTED instead: <see cref="LayoutValidator"/>'s <c>compatibility-mode</c> check warns
/// whenever a layout containing repeaters declares a mode below 15, which surfaces in
/// <c>create_layout</c>'s own returned <c>quickValidation</c> at the moment a template introduces it.
/// </para>
/// </remarks>
public static class DocumentSettingsScaffold
{
    /// <summary>
    /// The compatibility mode every stock corpus layout declares — Word 2013 and later, the first mode in
    /// which repeating-section content controls exist.
    /// </summary>
    public const int Word2013CompatibilityMode = 15;

    /// <summary>
    /// The mode Word applies to a document that declares none, i.e. what a layout with no settings part (or
    /// no <c>compatibilityMode</c> setting) effectively is. Measured, not assumed — see this type's summary.
    /// </summary>
    public const int ImpliedCompatibilityMode = 12;

    /// <summary>The vendor URI a <c>w:compatSetting</c> carries; corpus-observed, identical in every capture.</summary>
    private const string CompatibilitySettingUri = "http://schemas.microsoft.com/office/word";

    /// <summary>
    /// Ensures <paramref name="main"/> has a <see cref="DocumentSettingsPart"/> declaring
    /// <c>compatibilityMode</c> = <see cref="Word2013CompatibilityMode"/>, adding the minimal part described
    /// in this type's remarks when it has none. Returns <c>true</c> only when the part was actually added —
    /// a document that already has a settings part is a no-op returning <c>false</c>, its own settings left
    /// untouched even when they declare a lower mode (see the remarks on why an existing document is never
    /// retrofitted; <see cref="ReadCompatibilityMode"/> is how that case gets reported instead).
    /// </summary>
    public static bool EnsureCompatibilityMode(MainDocumentPart main)
    {
        ArgumentNullException.ThrowIfNull(main);
        if (main.DocumentSettingsPart is not null)
        {
            return false;
        }

        var settingsPart = main.AddNewPart<DocumentSettingsPart>();
        settingsPart.Settings = new Settings(
            new Compatibility(
                new CompatibilitySetting
                {
                    Name = CompatSettingNameValues.CompatibilityMode,
                    Uri = CompatibilitySettingUri,
                    Val = Word2013CompatibilityMode.ToString(System.Globalization.CultureInfo.InvariantCulture),
                }));
        settingsPart.Settings.Save();
        return true;
    }

    /// <summary>
    /// The compatibility mode <paramref name="main"/> effectively has: the value of its
    /// <c>compatibilityMode</c> <c>w:compatSetting</c>, or <see cref="ImpliedCompatibilityMode"/> when the
    /// document has no settings part, no <c>w:compat</c>, or no such setting. Returns <c>null</c> only when
    /// the setting is present but carries a value that is not an integer — a shape no observed layout has,
    /// reported by the caller as "cannot be determined" rather than guessed at in either direction.
    /// </summary>
    public static int? ReadCompatibilityMode(MainDocumentPart main)
    {
        ArgumentNullException.ThrowIfNull(main);

        var setting = main.DocumentSettingsPart?.Settings
            ?.GetFirstChild<Compatibility>()
            ?.Elements<CompatibilitySetting>()
            .FirstOrDefault(s => s.Name is not null && s.Name.Value == CompatSettingNameValues.CompatibilityMode);

        if (setting is null)
        {
            return ImpliedCompatibilityMode;
        }

        return int.TryParse(
            setting.Val?.Value,
            System.Globalization.NumberStyles.Integer,
            System.Globalization.CultureInfo.InvariantCulture,
            out var mode)
            ? mode
            : null;
    }
}
