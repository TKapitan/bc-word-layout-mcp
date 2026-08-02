namespace BcWordLayout.Domain;

/// <summary>
/// Well-known OOXML namespace URIs and BC-specific markers used when reading Word layouts.
/// Attribute values (xpath, storeItemID, prefixMappings) live in the <c>w:</c> namespace even
/// on the <c>w15:dataBinding</c> element, so the same attribute-reading logic works for both.
/// </summary>
public static class OoxmlNames
{
    /// <summary>wordprocessingml main namespace (prefix <c>w</c>).</summary>
    public const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>Word 2012 namespace (prefix <c>w15</c>) — repeating sections live here.</summary>
    public const string W15 = "http://schemas.microsoft.com/office/word/2012/wordml";

    /// <summary>
    /// Namespace prefix that identifies the Business Central report dataset custom XML part.
    /// Real BC24+ layouts use the NAV form: <c>urn:microsoft-dynamics-nav/reports/&lt;Name&gt;/&lt;id&gt;/</c>.
    /// Used to distinguish the BC part from the unrelated Office bibliography custom XML part.
    /// </summary>
    public const string BcNamespacePrefix = "urn:microsoft-dynamics-nav/reports/";

    /// <summary>Root element name of the BC dataset custom XML part.</summary>
    public const string RootElementName = "NavWordReportXmlPart";

    /// <summary>System metadata element (first child of the root); not part of the business dataset.</summary>
    public const string SystemNodeName = "BCReportInformation";
}
