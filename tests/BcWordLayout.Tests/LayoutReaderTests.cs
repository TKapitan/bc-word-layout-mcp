using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Tests;

/// <remarks>Joins the label-convention-seam collection: <c>InventoryOrderDetails_controls</c> pins
/// <c>InventoryOrderDetails.docx</c>'s zero-<see cref="ControlKind.Label"/>-control count, which a
/// concurrently-swapped <c>LabelConvention.Current</c> could disturb (see
/// <see cref="LabelConventionSeamCollection"/>).</remarks>
[Collection("label-convention-seam")]
public class LayoutReaderTests
{
    private static int RepeaterDepth(LayoutControl repeater)
    {
        var depth = 1;
        var parent = repeater.ParentRepeater;
        while (parent is not null)
        {
            depth++;
            parent = parent.ParentRepeater;
        }

        return depth;
    }

    [Fact]
    public void SalesInvoice_controls()
    {
        var inv = LayoutReader.Read(Corpus.Path(Corpus.SalesInvoice));

        Assert.Equal(9, inv.RepeaterCount);

        // Field- and label-kind controls in the main document (95 controls, split Field/Label by name).
        var docFieldsAndLabels = inv.Controls
            .Count(c => c.Part == "document.xml" && (c.Kind == ControlKind.Field || c.Kind == ControlKind.Label));
        Assert.Equal(95, docFieldsAndLabels);

        // Bound field controls also present in the two headers (company picture + address fields).
        var headerBound = inv.Controls.Count(c =>
            c.Part.StartsWith("header", StringComparison.Ordinal) &&
            (c.Kind == ControlKind.Field || c.Kind == ControlKind.Label || c.Kind == ControlKind.Picture));
        Assert.True(headerBound >= 5, $"expected header bindings, found {headerBound}");

        // Two picture controls (one per header).
        Assert.Equal(2, inv.Controls.Count(c => c.Kind == ControlKind.Picture));

        // Nesting reaches 3 levels deep (ContractBillingDetailsMapping > ContractBillingDetailsGrouping >
        // ContractBillingDetails).
        var maxDepth = inv.Controls.Where(c => c.Kind == ControlKind.Repeater).Max(RepeaterDepth);
        Assert.Equal(3, maxDepth);
    }

    [Fact]
    public void InventoryOrderDetails_controls()
    {
        var inv = LayoutReader.Read(Corpus.Path(Corpus.InventoryOrderDetails));

        Assert.Equal(2, inv.RepeaterCount);
        Assert.Equal(0, inv.Controls.Count(c => c.Kind == ControlKind.Label));
    }

    [Fact]
    public void StandardStatement_controls_and_deep_nesting()
    {
        var inv = LayoutReader.Read(Corpus.Path(Corpus.StandardStatement));

        Assert.Equal(10, inv.RepeaterCount);

        // Nesting reaches 4 levels deep (Customer > CurrencyLoop > OverdueVisible > CustLedgEntry2).
        var maxDepth = inv.Controls.Where(c => c.Kind == ControlKind.Repeater).Max(RepeaterDepth);
        Assert.Equal(4, maxDepth);
    }

    [Fact]
    public void StandardStatement_has_a_paragraph_wrapped_repeater()
    {
        // The top-level /Customer repeater's repeatingSectionItem wraps block content (w:p/w:tbl),
        // not a table row — verify at least one such repeater exists.
        using var doc = WordprocessingDocument.Open(Corpus.Path(Corpus.StandardStatement), false);
        var body = doc.MainDocumentPart!.Document!;

        var repeaterItems = body.Descendants<SdtElement>()
            .Where(s => s.GetFirstChild<SdtProperties>()?
                .Elements()
                .Any(e => e.LocalName == "repeatingSectionItem" &&
                          e.NamespaceUri == "http://schemas.microsoft.com/office/word/2012/wordml") == true);

        var paragraphWrapped = repeaterItems
            .Select(s => s.GetFirstChild<SdtContentBlock>())
            .Where(content => content is not null)
            .Any(content => content!.Elements().Any(e => e is Paragraph));

        Assert.True(paragraphWrapped, "expected at least one repeatingSectionItem wrapping a paragraph");
    }

    [Fact]
    public void Controls_inside_repeaters_link_to_their_enclosing_repeater()
    {
        var inv = LayoutReader.Read(Corpus.Path(Corpus.SalesInvoice));

        Assert.Contains(inv.Controls, c => c.ParentRepeater is not null);

        // Every control with a parent repeater points at a control that is actually a repeater.
        Assert.All(inv.Controls.Where(c => c.ParentRepeater is not null),
            c => Assert.Equal(ControlKind.Repeater, c.ParentRepeater!.Kind));
    }

    // ---- table structure + per-control level/coordinates ----

    [Fact]
    public void SalesInvoice_marks_address_fields_as_cell_level_and_locates_them_in_a_table()
    {
        var inv = LayoutReader.Read(Corpus.Path(Corpus.SalesInvoice));

        Assert.NotEmpty(inv.Tables);

        // #Nav: /Header/CustomerAddress6 — a cell-level control (SdtCell wrapping a whole w:tc).
        var addr6 = inv.Controls.Single(c => c.SdtId == -2064325541);
        Assert.Equal(SdtLevel.Cell, addr6.Level);
        Assert.NotNull(addr6.TableIndex);
        Assert.NotNull(addr6.RowIndex);
        Assert.NotNull(addr6.ColIndex);

        // The same control appears in the table structure as the owner of a control cell.
        var owningCell = inv.Tables
            .SelectMany(t => t.Rows)
            .SelectMany(r => r.Cells)
            .SingleOrDefault(cell => cell.ControlId == -2064325541);
        Assert.NotNull(owningCell);
        Assert.True(owningCell!.IsControlCell);
        Assert.Equal(addr6.ColIndex, owningCell.ColIndex);
    }

    [Fact]
    public void TableStructure_describes_grid_rows_cells_and_control_ownership()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.TableWithCellLevelControl(id: 99, text: "cell-text"));
        try
        {
            var inv = LayoutReader.Read(path);

            var table = Assert.Single(inv.Tables);
            Assert.Equal("document.xml", table.Part);
            Assert.Equal(0, table.TableIndex);
            Assert.Equal(2, table.ColumnCount);
            Assert.Equal(1, table.RowCount);

            var row = Assert.Single(table.Rows);
            Assert.False(row.IsControlRow);
            Assert.Equal(2, row.Cells.Count);

            Assert.True(row.Cells[0].IsControlCell);
            Assert.Equal(99, row.Cells[0].ControlId);
            Assert.Equal(0, row.Cells[0].ColIndex);
            Assert.Contains("cell-text", row.Cells[0].Text);
            Assert.False(row.Cells[1].IsControlCell);
            Assert.Equal(1, row.Cells[1].ColIndex);

            var control = inv.Controls.Single(c => c.SdtId == 99);
            Assert.Equal(SdtLevel.Cell, control.Level);
            Assert.Equal(0, control.TableIndex);
            Assert.Equal(0, control.RowIndex);
            Assert.Equal(0, control.ColIndex);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void TableStructure_ColumnCount_counts_every_gridCol_even_one_with_no_declared_width()
    {
        // A w:gridCol without w:w is schema-legal. ColumnCount must count it too (matching
        // TableGridNavigator.GridColumnCount, which every editor already uses), not just the width-bearing
        // ones — otherwise get_layout_info's documented "pass columnCount from here" contract breaks for
        // this table shape.
        var body =
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/><w:gridCol/></w:tblGrid>"
            + "<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>c1</w:t></w:r></w:p></w:tc>"
            + "<w:tc><w:tcPr/><w:p><w:r><w:t>c2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";
        var path = SyntheticLayout.Create(body);
        try
        {
            var inv = LayoutReader.Read(path);
            var table = Assert.Single(inv.Tables);

            Assert.Equal(2, table.ColumnCount);
            // GridColumnWidths stays filtered to width-bearing gridCols only (documented behaviour) —
            // deliberately NOT index-aligned with ColumnCount for this shape.
            Assert.Single(table.GridColumnWidths);
            Assert.Equal(2000, table.GridColumnWidths[0]);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RowLevel_control_is_reported_as_a_control_row_with_no_column_index()
    {
        var path = SyntheticLayout.Create(SyntheticLayout.TableWithRowLevelControl(id: 111, text: "row-text"));
        try
        {
            var inv = LayoutReader.Read(path);

            var table = Assert.Single(inv.Tables);
            var row = Assert.Single(table.Rows);
            Assert.True(row.IsControlRow);
            Assert.Equal(111, row.ControlId);
            Assert.Single(row.Cells);

            var control = inv.Controls.Single(c => c.SdtId == 111);
            Assert.Equal(SdtLevel.Row, control.Level);
            Assert.Equal(0, control.TableIndex);
            Assert.Equal(0, control.RowIndex);
            Assert.Null(control.ColIndex); // a row-level control spans the whole row, not one column
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Non_table_controls_report_their_level_but_carry_no_coordinates()
    {
        var path = SyntheticLayout.Create(
            SyntheticLayout.InlineControlWithId(55, "x") + SyntheticLayout.BlockControlWithId(77, "y"));
        try
        {
            var inv = LayoutReader.Read(path);

            var inline = inv.Controls.Single(c => c.SdtId == 55);
            Assert.Equal(SdtLevel.Run, inline.Level);
            Assert.Null(inline.TableIndex);
            Assert.Null(inline.RowIndex);
            Assert.Null(inline.ColIndex);

            var block = inv.Controls.Single(c => c.SdtId == 77);
            Assert.Equal(SdtLevel.Block, block.Level);
            Assert.Null(block.TableIndex);

            Assert.Empty(inv.Tables);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
