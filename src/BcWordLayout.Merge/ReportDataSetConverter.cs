using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;

namespace BcWordLayout.Merge;

/// <summary>
/// Converts what Business Central's report UI actually exports (*Send to → XML*: a namespace-less
/// <c>ReportDataSet</c> document with <c>Labels/Label[@name]</c> and
/// <c>DataItems/DataItem[@name]/Columns/Column[@name]</c>, one <c>DataItem</c> sibling per row) into the
/// layout's own data-store part shape (a <c>NavWordReportXmlPart</c> root in the report's
/// <c>urn:microsoft-dynamics-nav</c> namespace, one ELEMENT per column) — the only shape
/// <c>SampleDataGenerator.LoadOverrides</c> resolves bindings against. Both encodings carry the same
/// information; this is the in-product version of the bridge that previously lived only in
/// <c>tools/e2e/bc_compare.py</c> (GitHub issue #4).
/// </summary>
internal static class ReportDataSetConverter
{
    /// <summary>Root element local name of the *Send to → XML* export shape.</summary>
    internal const string ExportRootElementName = "ReportDataSet";

    /// <summary>
    /// Converts <paramref name="export"/> into the <c>NavWordReportXmlPart</c> shape in
    /// <paramref name="report"/>'s namespace. The export root's <c>id</c> is cross-checked against the
    /// report id the layout's dataset namespace encodes, so feeding the wrong report's export fails with
    /// an actionable message instead of a preview where every binding is silently unresolved. The root's
    /// <c>wordMergeDataItem</c> attribute is deliberately dropped — it names the data item Word's own
    /// mail-merge would iterate, which the merge engine derives from the layout's repeating sections
    /// instead.
    /// </summary>
    internal static XDocument ToNavWordReportXmlPart(XDocument export, ReportIdentity report)
    {
        var src = export.Root
            ?? throw new InvalidDataException("Data overrides XML has no root element.");

        var exportId = (string?)src.Attribute("id") ?? "";
        if (exportId.Length > 0 && report.ReportId.Length > 0
            && !string.Equals(exportId, report.ReportId, StringComparison.Ordinal))
        {
            throw new InvalidDataException(
                $"Data overrides export is from report {exportId} ('{(string?)src.Attribute("name")}'), but "
                + $"the layout's dataset belongs to report {report.ReportId} ('{report.ReportName}') — its "
                + "column bindings would all come out unresolved. Export the dataset from the report this "
                + "layout is attached to (run the report, then Send to → XML).");
        }

        var ns = XNamespace.Get(report.Namespace);
        var culture = ResolveCulture(src);
        var root = new XElement(ns + OoxmlNames.RootElementName);

        // Report metadata, when present, mirrors the layout's own /BCReportInformation subtree verbatim.
        var info = src.Element("BCReportInformation");
        if (info != null)
        {
            root.Add(CopyVerbatim(info, ns));
        }

        var labels = src.Element("Labels");
        if (labels != null)
        {
            var target = new XElement(ns + "Labels");
            foreach (var label in labels.Elements("Label"))
            {
                var name = (string?)label.Attribute("name");
                if (!string.IsNullOrEmpty(name))
                {
                    target.Add(new XElement(ColumnName(ns, name), label.Value));
                }
            }

            root.Add(target);
        }

        EmitDataItems(src, root, ns, culture);
        return new XDocument(root);
    }

    /// <summary>
    /// One element per <c>DataItem</c> occurrence: BC emits a sibling per ROW, which is exactly the shape
    /// the merge engine expands a repeater over — order is preserved, nesting recurses.
    /// </summary>
    private static void EmitDataItems(XElement container, XElement target, XNamespace ns, CultureInfo culture)
    {
        foreach (var item in container.Elements("DataItems").Elements("DataItem"))
        {
            var name = (string?)item.Attribute("name");
            if (string.IsNullOrEmpty(name))
            {
                continue;
            }

            var el = new XElement(ColumnName(ns, name));
            foreach (var column in item.Elements("Columns").Elements("Column"))
            {
                var columnName = (string?)column.Attribute("name");
                if (string.IsNullOrEmpty(columnName))
                {
                    continue;
                }

                el.Add(new XElement(ColumnName(ns, columnName), FormatColumnValue(column, culture)));
            }

            EmitDataItems(item, el, ns, culture);
            target.Add(el);
        }
    }

    /// <summary>
    /// A column carrying a <c>decimalformatter</c> attribute (e.g. <c>#,##0.00</c>, or the sectioned
    /// <c>$#,##0.00;$-#,##0.00</c> real exports contain) holds a RAW invariant number in its text, and BC's
    /// render applies that formatter — so this must too, or every such amount in the preview differs from
    /// the real render (<c>100</c> vs <c>100.00</c>). Other columns arrive pre-formatted and are copied
    /// verbatim: the rule is strictly per-column, never global. A raw value that does not parse as an
    /// invariant decimal is also copied verbatim — mangling a value is worse than leaving it unformatted.
    /// </summary>
    private static string FormatColumnValue(XElement column, CultureInfo culture)
    {
        var text = column.Value;
        var formatter = (string?)column.Attribute("decimalformatter");
        if (string.IsNullOrEmpty(formatter)
            || !decimal.TryParse(text, NumberStyles.Number, CultureInfo.InvariantCulture, out var value))
        {
            return text;
        }

        return value.ToString(formatter, culture);
    }

    /// <summary>
    /// The culture whose separator glyphs <c>decimalformatter</c> patterns are rendered with. The export
    /// root's <c>formatRegion</c> is BC's regional-format setting — the same one BC's own render uses — so
    /// it wins; <c>language</c> (the UI language) is only a fallback for an export that omits it, and a
    /// missing or unrecognized name falls back to the invariant culture rather than failing the preview.
    /// </summary>
    private static CultureInfo ResolveCulture(XElement exportRoot)
    {
        foreach (var attribute in new[] { "formatRegion", "language" })
        {
            var name = (string?)exportRoot.Attribute(attribute);
            if (string.IsNullOrEmpty(name))
            {
                continue;
            }

            try
            {
                return CultureInfo.GetCultureInfo(name);
            }
            catch (CultureNotFoundException)
            {
                // Fall through to the next candidate.
            }
        }

        return CultureInfo.InvariantCulture;
    }

    /// <summary>Copies an element subtree (names and text only) into <paramref name="ns"/>.</summary>
    private static XElement CopyVerbatim(XElement source, XNamespace ns)
    {
        var el = new XElement(ColumnName(ns, source.Name.LocalName));
        var children = source.Elements().ToList();
        if (children.Count == 0)
        {
            el.Value = source.Value;
            return el;
        }

        foreach (var child in children)
        {
            el.Add(CopyVerbatim(child, ns));
        }

        return el;
    }

    /// <summary>
    /// Builds the qualified name for an export-supplied element/column name, turning an invalid XML name
    /// into an actionable error naming the offender instead of a bare <see cref="XmlException"/>.
    /// </summary>
    private static XName ColumnName(XNamespace ns, string name)
    {
        try
        {
            XmlConvert.VerifyNCName(name);
        }
        catch (XmlException ex)
        {
            throw new InvalidDataException(
                $"Data overrides export contains a name that is not a valid XML element name: '{name}'.", ex);
        }

        return ns + name;
    }
}
