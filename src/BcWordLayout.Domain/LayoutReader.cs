using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Domain;

/// <summary>
/// Walks a layout's main document plus every header and footer part, recursively descending through
/// <c>w:sdt</c> content controls (including nested repeaters) to produce a flat control inventory.
/// Controls are classified from their <c>w:sdtPr</c> markers; the enclosing repeater is tracked so
/// nesting is visible to callers.
/// </summary>
public static class LayoutReader
{
    public static LayoutInventory Read(string docxPath)
    {
        if (!File.Exists(docxPath))
        {
            throw new FileNotFoundException("Layout file not found.", docxPath);
        }

        using var doc = WordprocessingDocument.Open(docxPath, false);
        return Read(doc);
    }

    /// <summary>Reads the inventory from an already-open document (lets callers share one package handle).</summary>
    public static LayoutInventory Read(WordprocessingDocument doc)
    {
        var main = doc.MainDocumentPart
            ?? throw new InvalidDataException("Layout has no main document part.");

        // Read the table structures once up front; the same pass yields a reference-keyed map from every
        // in-table sdt to its (table, row, column) coordinates, which each control below is stamped with.
        var (tables, coords) = TableStructureReader.Read(doc);

        var controls = new List<LayoutControl>();
        var parts = new List<string>();

        foreach (var (root, partName) in PartWalker.ContentParts(main))
        {
            parts.Add(partName);
            Walk(root, partName, parentRepeater: null, coords, controls, depth: 0);
        }

        return new LayoutInventory { Controls = controls, Parts = parts, Tables = tables };
    }

    /// <summary>
    /// <paramref name="depth"/> (0 at the top of each part's walk) enforces
    /// <see cref="ResourceLimits.MaxElementNestingDepth"/>. Unlike the schema tree — bounded once at
    /// construction by <see cref="SchemaProvider.BuildNode"/> — a document.xml/header/footer's own element
    /// nesting has no upstream ceiling, so this hand-rolled recursive walk must bound itself: a crafted file
    /// nested tens of thousands of levels deep would overflow the stack, and a
    /// <see cref="StackOverflowException"/> cannot be caught, so it would take the whole server down instead
    /// of failing one call.
    /// </summary>
    private static void Walk(
        OpenXmlElement element,
        string partName,
        LayoutControl? parentRepeater,
        IReadOnlyDictionary<SdtElement, TableStructureReader.Coord> coords,
        List<LayoutControl> sink,
        int depth)
    {
        if (depth > ResourceLimits.MaxElementNestingDepth)
        {
            throw ResourceLimits.DepthExceeded("Document element", ResourceLimits.MaxElementNestingDepth);
        }

        foreach (var child in element.ChildElements)
        {
            if (child is SdtElement sdt)
            {
                var pr = sdt.GetFirstChild<SdtProperties>();
                var classification = SdtInspector.Classify(pr);

                if (classification == SdtInspector.Classification.RepeaterItem)
                {
                    // Structural wrapper only: not surfaced as an inventory entry, but the enclosing
                    // repeater context flows unchanged to the controls inside its content.
                    Walk(sdt, partName, parentRepeater, coords, sink, depth + 1);
                    continue;
                }

                var control = BuildControl(sdt, partName, parentRepeater, LevelOf(sdt), Coords(sdt, coords));
                sink.Add(control);

                var childRepeater = control.Kind == ControlKind.Repeater ? control : parentRepeater;
                Walk(sdt, partName, childRepeater, coords, sink, depth + 1);
            }
            else
            {
                Walk(child, partName, parentRepeater, coords, sink, depth + 1);
            }
        }
    }

    private static SdtLevel LevelOf(SdtElement sdt) => sdt switch
    {
        SdtRun => SdtLevel.Run,
        SdtBlock => SdtLevel.Block,
        SdtCell => SdtLevel.Cell,
        SdtRow => SdtLevel.Row,
        SdtRunRuby => SdtLevel.RunRuby,
        _ => SdtLevel.Unknown,
    };

    private static TableStructureReader.Coord? Coords(SdtElement sdt, IReadOnlyDictionary<SdtElement, TableStructureReader.Coord> coords) =>
        coords.TryGetValue(sdt, out var c) ? c : null;

    private static LayoutControl BuildControl(
        SdtElement sdt,
        string partName,
        LayoutControl? parentRepeater,
        SdtLevel level,
        TableStructureReader.Coord? coord)
    {
        return new LayoutControl
        {
            Kind = SdtInspector.ClassifyControlKind(sdt),
            Alias = SdtInspector.ReadAlias(sdt),
            Tag = SdtInspector.ReadTag(sdt),
            XPath = SdtInspector.ReadXPath(sdt),
            StoreItemId = SdtInspector.ReadStoreItemId(sdt),
            BindingNamespace = SdtInspector.ReadBindingNamespace(sdt),
            Part = partName,
            SdtId = SdtInspector.ReadControlId(sdt),
            UsesW15Binding = SdtInspector.UsesW15Binding(sdt),
            ParentRepeater = parentRepeater,
            Level = level,
            TableIndex = coord?.TableIndex,
            RowIndex = coord?.RowIndex,
            ColIndex = coord?.ColIndex,
        };
    }
}
