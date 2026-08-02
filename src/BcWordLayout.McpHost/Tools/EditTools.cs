using System.ComponentModel;
using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;
using DocumentFormat.OpenXml.Packaging;
using ModelContextProtocol.Server;
using static BcWordLayout.McpHost.Tools.ToolGuards;

namespace BcWordLayout.McpHost.Tools;

/// <summary>
/// MCP tools for editing an EXISTING layout's content: inserting a bound field/label control
/// (<c>insert_field</c>, <c>insert_label</c>), removing a control (<c>remove_control</c>), and setting/
/// clearing a PLAIN-TEXT table cell (<c>set_cell_text</c>, <c>clear_cell_text</c>). Table-STRUCTURE edits
/// (adding/removing/resizing/merging/splitting columns, and the repeater-table author tool) live in
/// <see cref="TableTools"/> instead. Every mutating tool here routes through <see cref="ToolGuards"/>'s
/// shared open/validate/save-or-reject choreography: <see cref="ToolGuards.GuardEdit"/> is the ONE AND ONLY
/// place an <see cref="EditResultDto"/> is constructed, so every tool that routes through it is structurally
/// incapable of reporting success without a freshly computed post-edit <c>QuickValidation</c>; likewise
/// <see cref="ToolGuards.GuardCellEdit"/> for <see cref="CellEditResultDto"/>. Both wrappers (and the
/// per-path edit lock they take — see <see cref="ToolGuards.EditLockFor"/>'s own remarks on why that
/// dictionary must stay a single shared instance across every tool family) live in <see cref="ToolGuards"/>,
/// not here, precisely so <see cref="TableTools"/>'s mutating tools serialize against these ones on the
/// same layout path too.
/// </summary>
[McpServerToolType]
public static class EditTools
{
    [McpServerTool(Name = "insert_field")]
    [Description("Insert a plain-text FIELD content control bound to a dataset path (e.g. "
                 + "'/Header/CustomerAddress1', as reported by list_dataset_fields) at the given location. "
                 + "Targets the main document body by default; pass layoutPart='header'/'footer' (optionally "
                 + "with partName) to target a header/footer part instead - with locationType='documentEnd' "
                 + "and no partName, a layout that has NO header/footer part yet (e.g. one built by an older "
                 + "create_layout) gets an empty one scaffolded and wired into the page setup automatically, "
                 + "so authoring a real per-page footer never dead-ends. Writes the layout in place: the "
                 + "edit is checked with OpenXmlValidator before saving — if it would produce structurally "
                 + "invalid OOXML, nothing is written and a structured failure is returned instead, leaving "
                 + "the file untouched. Returns the inserted control's details plus a post-edit quick "
                 + "validation summary.")]
    public static ToolResponse InsertField(
        [Description("Absolute path to the .docx layout file.")] string layoutPath,
        [Description("Dataset path to a non-label leaf column, e.g. '/Header/CustomerAddress1' (see "
                     + "list_dataset_fields).")] string field,
        [Description("Where to insert: 'documentEnd', 'afterControl', 'tableCell', or 'atText' "
                     + "(case-insensitive).")] string locationType,
        [Description("Required for 'afterControl': the w:id of the control to insert after.")] int? controlId = null,
        [Description("Required for 'tableCell': 0-based table index in document order.")] int? tableIndex = null,
        [Description("Required for 'tableCell': 0-based row index within the table.")] int? row = null,
        [Description("Required for 'tableCell': 0-based cell/column index within the row.")] int? col = null,
        [Description("Required for 'atText': substring to search for in existing run text (ordinal match).")]
        string? searchText = null,
        [Description("Which OOXML part locationType resolves within: 'body' (default), 'header', or "
                     + "'footer' (case-insensitive).")] string layoutPart = "body",
        [Description("Only used when layoutPart is 'header'/'footer': a specific part file name, e.g. "
                     + "'header2.xml' (see get_layout_info's controls[].part for the names actually present). "
                     + "Omit to use the FIRST header/footer part.")] string? partName = null,
        [Description("Optional: true makes the control's text bold, false strips bold; omit to leave the "
                     + "control's runs unstyled (a control in a fresh plain-table cell has nothing to "
                     + "inherit).")] bool? bold = null,
        [Description("Optional font size in points (4-96, halves allowed) for the control's text; omit to "
                     + "leave it unstyled.")] double? fontSizePoints = null)
    {
        return GuardEdit(layoutPath, doc =>
        {
            var location = BuildLocation(locationType, controlId, tableIndex, row, col, searchText, layoutPart, partName);
            return LayoutEditor.InsertField(doc, field, location, ControlFormat(bold, fontSizePoints));
        });
    }

    [McpServerTool(Name = "insert_label")]
    [Description("Insert a plain-text LABEL content control bound to a label-shaped dataset path (e.g. "
                 + "'/Header/YourReference_Lbl'; by default a name ending 'Lbl'/'_Lbl' - see "
                 + "list_dataset_fields's isLabel flag, which always reflects the server's active convention) "
                 + "at the given location. Same write/validate/save-or-reject safety and response shape as "
                 + "insert_field.")]
    public static ToolResponse InsertLabel(
        [Description("Absolute path to the .docx layout file.")] string layoutPath,
        [Description("Dataset path to a label column (by default a name ending in 'Lbl'/'_Lbl' - see "
                     + "list_dataset_fields's isLabel flag), e.g. '/Header/YourReference_Lbl'.")] string label,
        [Description("Where to insert: 'documentEnd', 'afterControl', 'tableCell', or 'atText' "
                     + "(case-insensitive).")] string locationType,
        [Description("Required for 'afterControl': the w:id of the control to insert after.")] int? controlId = null,
        [Description("Required for 'tableCell': 0-based table index in document order.")] int? tableIndex = null,
        [Description("Required for 'tableCell': 0-based row index within the table.")] int? row = null,
        [Description("Required for 'tableCell': 0-based cell/column index within the row.")] int? col = null,
        [Description("Required for 'atText': substring to search for in existing run text (ordinal match).")]
        string? searchText = null,
        [Description("Which OOXML part locationType resolves within: 'body' (default), 'header', or "
                     + "'footer' (case-insensitive).")] string layoutPart = "body",
        [Description("Only used when layoutPart is 'header'/'footer': a specific part file name, e.g. "
                     + "'header2.xml' (see get_layout_info's controls[].part for the names actually present). "
                     + "Omit to use the FIRST header/footer part.")] string? partName = null,
        [Description("Optional: true makes the control's text bold, false strips bold; omit to leave the "
                     + "control's runs unstyled (a control in a fresh plain-table cell has nothing to "
                     + "inherit).")] bool? bold = null,
        [Description("Optional font size in points (4-96, halves allowed) for the control's text; omit to "
                     + "leave it unstyled.")] double? fontSizePoints = null)
    {
        return GuardEdit(layoutPath, doc =>
        {
            var location = BuildLocation(locationType, controlId, tableIndex, row, col, searchText, layoutPart, partName);
            return LayoutEditor.InsertLabel(doc, label, location, ControlFormat(bold, fontSizePoints));
        });
    }

    [McpServerTool(Name = "insert_text")]
    [Description("Insert plain STATIC text - a literal run, NOT a content control and NOT bound to any "
                 + "dataset field. This is the glue between controls: the separator space, the colon, the "
                 + "' / ' between a date and a page number that every real BC header/footer uses. Reach for "
                 + "it when two inline controls would otherwise run together with nothing between them "
                 + "('Document NoDOCU-0150'), or when a layout needs a fixed caption that is not a dataset "
                 + "label (use insert_label for anything the dataset provides, so it stays translatable). "
                 + "Same locations as insert_field, including header/footer (a header/footer-targeted "
                 + "'documentEnd' insert scaffolds an empty part first if the layout has none). IMPORTANT: "
                 + "because it creates no content control there is no controlId in the response (it is 0), "
                 + "and the run CANNOT afterwards be targeted by remove_control or used as an afterControl "
                 + "anchor - reach it with an 'atText' location or set_cell_text instead. Whitespace is "
                 + "preserved exactly, so text=' ' inserts a real separator space. Same "
                 + "write/validate/save-or-reject safety as insert_field.")]
    public static ToolResponse InsertText(
        [Description("Absolute path to the .docx layout file.")] string layoutPath,
        [Description("The literal text to insert. Whitespace is preserved verbatim, so ' ' is a valid "
                     + "separator space; an EMPTY string is rejected (nothing to insert).")] string text,
        [Description("Where to insert: 'documentEnd', 'afterControl', 'tableCell', or 'atText' "
                     + "(case-insensitive).")] string locationType,
        [Description("Required for 'afterControl': the w:id of the control to insert after.")] int? controlId = null,
        [Description("Required for 'tableCell': 0-based table index in document order.")] int? tableIndex = null,
        [Description("Required for 'tableCell': 0-based row index within the table.")] int? row = null,
        [Description("Required for 'tableCell': 0-based cell/column index within the row.")] int? col = null,
        [Description("Required for 'atText': substring to search for in existing run text (ordinal match).")]
        string? searchText = null,
        [Description("Which OOXML part locationType resolves within: 'body' (default), 'header', or "
                     + "'footer' (case-insensitive).")] string layoutPart = "body",
        [Description("Only used when layoutPart is 'header'/'footer': a specific part file name, e.g. "
                     + "'header2.xml' (see get_layout_info's controls[].part for the names actually present). "
                     + "Omit to use the FIRST header/footer part.")] string? partName = null,
        [Description("Optional: true makes the text bold, false strips bold; omit to leave it unstyled so it "
                     + "inherits whatever surrounds it.")] bool? bold = null,
        [Description("Optional font size in points (4-96, halves allowed); omit to leave it unstyled.")]
        double? fontSizePoints = null)
    {
        return GuardEdit(layoutPath, doc =>
        {
            var location = BuildLocation(locationType, controlId, tableIndex, row, col, searchText, layoutPart, partName);
            return LayoutEditor.InsertText(doc, text, location, ControlFormat(bold, fontSizePoints));
        });
    }

    [McpServerTool(Name = "insert_picture")]
    [Description("Insert a PICTURE content control bound to a picture dataset path (e.g. "
                 + "'/Header/CompanyPicture' - the company logo every real BC document header carries), at "
                 + "the given location. This is the tool for putting a logo into a layout authored from "
                 + "scratch: merge/preview already FILL picture placeholders, but nothing could create one. "
                 + "The control is inserted with a gray placeholder image embedded (a picture frame whose "
                 + "image reference dangles is a corrupt document to Word) sized widthMm x heightMm - "
                 + "default 30x30 mm, the size the real BC corpus logo uses; BC replaces the image itself at "
                 + "render time, so the placeholder is only what you see in preview_layout. Targets the body "
                 + "by default; pass layoutPart='header'/'footer' (a logo usually belongs in the header) - "
                 + "with locationType='documentEnd' a layout that has no such part yet gets an empty one "
                 + "scaffolded first. Same write/validate/save-or-reject safety and response shape as "
                 + "insert_field.")]
    public static ToolResponse InsertPicture(
        [Description("Absolute path to the .docx layout file.")] string layoutPath,
        [Description("Dataset path to the picture column, e.g. '/Header/CompanyPicture' (see "
                     + "list_dataset_fields).")] string picture,
        [Description("Where to insert: 'documentEnd', 'afterControl', 'tableCell', or 'atText' "
                     + "(case-insensitive).")] string locationType,
        [Description("Required for 'afterControl': the w:id of the control to insert after.")] int? controlId = null,
        [Description("Required for 'tableCell': 0-based table index in document order.")] int? tableIndex = null,
        [Description("Required for 'tableCell': 0-based row index within the table.")] int? row = null,
        [Description("Required for 'tableCell': 0-based cell/column index within the row.")] int? col = null,
        [Description("Required for 'atText': substring to search for in existing run text (ordinal match).")]
        string? searchText = null,
        [Description("Picture frame width in millimetres (1-500); default 30, the corpus logo's own size.")]
        double widthMm = 30,
        [Description("Picture frame height in millimetres (1-500); default 30.")] double heightMm = 30,
        [Description("Which OOXML part locationType resolves within: 'body' (default), 'header', or "
                     + "'footer' (case-insensitive).")] string layoutPart = "body",
        [Description("Only used when layoutPart is 'header'/'footer': a specific part file name, e.g. "
                     + "'header2.xml'. Omit to use the FIRST header/footer part.")] string? partName = null)
    {
        return GuardEdit(layoutPath, doc =>
        {
            var location = BuildLocation(locationType, controlId, tableIndex, row, col, searchText, layoutPart, partName);
            return LayoutEditor.InsertPicture(doc, picture, location, widthMm, heightMm);
        });
    }

    /// <summary>Maps insert_field/insert_label's optional run-formatting knobs to the domain record (null when both omitted).</summary>
    private static CellTextFormat? ControlFormat(bool? bold, double? fontSizePoints) =>
        bold is null && fontSizePoints is null ? null : new CellTextFormat { Bold = bold, FontSizePoints = fontSizePoints };

    [McpServerTool(Name = "remove_control")]
    [Description("Remove a content control (field, label, repeater, picture, or unbound) identified by its "
                 + "w:id, searching the main document body plus every header and footer. When keepText is "
                 + "true, the control wrapper is removed but its content (visible text/row/cell) is kept in "
                 + "place; otherwise the control and its content are both removed. CELL-LEVEL CONTROLS (level="
                 + "'cell' in get_layout_info - e.g. BC header address fields like CustomerAddress1..6, which "
                 + "each wrap a whole table cell = one column) ARE HANDLED SPECIALLY: the table cell is ALWAYS "
                 + "preserved so the table grid is never broken. keepText=false empties that cell (removes the "
                 + "field AND its text, leaving a blank column); keepText=true keeps the cell's text too "
                 + "(only the binding is dropped). Either way the column stays - removing an address field no "
                 + "longer deletes its column. (To truly drop a column you must remove its cell from every "
                 + "row and adjust the grid, which is out of scope for a single-control edit.) Same write/"
                 + "validate/save-or-reject safety as insert_field.")]
    public static ToolResponse RemoveControl(
        [Description("Absolute path to the .docx layout file.")] string layoutPath,
        [Description("The w:id of the control to remove (see get_layout_info's control inventory).")] int controlId,
        [Description("When true, keep the control's content in place and only remove the content-control "
                     + "wrapper. Default false (removes the control and its content). For a cell-level control "
                     + "the cell/column is preserved either way; keepText only decides whether the cell keeps "
                     + "its text (true) or is emptied (false).")] bool keepText = false)
    {
        return GuardEdit(layoutPath, doc => LayoutEditor.RemoveControl(doc, controlId, keepText));
    }

    [McpServerTool(Name = "set_cell_text")]
    [Description("Set (replace) the PLAIN TEXT of a table cell, addressed by its 0-based table/row/column "
                 + "index (the SAME indices get_layout_info reports in its tables[] section). This is for "
                 + "cells that are ordinary text, NOT content controls - e.g. a line-items table's column-"
                 + "HEADER labels (like 'GST Amount' or 'Amount (ex. GST)'), which BC authors as plain cell "
                 + "text. Any existing paragraphs/runs in the cell are collapsed into one run carrying the "
                 + "given text; the cell's width/span/borders (w:tcPr) and its paragraph/run styling (pStyle, "
                 + "font) are preserved, so a re-labelled header keeps its look. The cell - and its table "
                 + "column - is always preserved. If the cell contains a content control (a bound field/"
                 + "label), this is REJECTED (use remove_control / insert_field / insert_label for those). "
                 + "To edit a cell in a header/footer, pass layoutPart='header'/'footer' (optionally partName). "
                 + "OPTIONAL FORMATTING: bold, alignment ('left'/'center'/'right'), and fontSizePoints (4-96) "
                 + "each apply only when passed - omit them all and the cell's existing look is preserved "
                 + "exactly as before. They exist for cells in a freshly authored plain table (insert_table), "
                 + "which have no styling to inherit: caption cells want bold, amount columns want right "
                 + "alignment, a document title wants a bigger size. "
                 + "Same write/validate/save-or-reject safety as insert_field.")]
    public static ToolResponse SetCellText(
        [Description("Absolute path to the .docx layout file.")] string layoutPath,
        [Description("0-based table index in document order (see get_layout_info's tables[]).")] int tableIndex,
        [Description("0-based row index within the table.")] int row,
        [Description("0-based cell/column index within the row.")] int col,
        [Description("The plain text to put in the cell (replaces whatever text was there).")] string text,
        [Description("Which OOXML part the cell is in: 'body' (default), 'header', or 'footer'.")] string layoutPart = "body",
        [Description("Only used when layoutPart is 'header'/'footer': a specific part file name (e.g. "
                     + "'header2.xml'); omit to use the first header/footer part.")] string? partName = null,
        [Description("Optional: true makes the text bold, false strips bold; omit to keep the cell's "
                     + "existing weight.")] bool? bold = null,
        [Description("Optional paragraph alignment 'left'/'center'/'right'; omit to keep the cell's "
                     + "existing alignment.")] string? alignment = null,
        [Description("Optional font size in points (4-96, halves allowed); omit to keep the cell's "
                     + "existing size.")] double? fontSizePoints = null)
    {
        return GuardCellEdit(layoutPath, doc =>
        {
            var location = BuildLocation("tableCell", null, tableIndex, row, col, null, layoutPart, partName);
            var format = bold is null && alignment is null && fontSizePoints is null
                ? null
                : new CellTextFormat { Bold = bold, Alignment = alignment, FontSizePoints = fontSizePoints };
            return CellTextEditor.SetCellText(doc, location, text, format);
        });
    }

    [McpServerTool(Name = "clear_cell_text")]
    [Description("Remove all PLAIN TEXT from a table cell, addressed by its 0-based table/row/column index "
                 + "(the SAME indices get_layout_info reports), leaving a valid EMPTY cell. This is the "
                 + "counterpart to set_cell_text - use it to blank a plain-text cell such as a now-unwanted "
                 + "column-header label (e.g. after remove_control emptied the matching data-field cells in a "
                 + "line-items table). The cell's width/span/borders (w:tcPr) and paragraph style are kept and "
                 + "the cell - and its table column - is preserved (only its text is removed). If the cell "
                 + "contains a content control, this is REJECTED (use remove_control instead). NOTE: this "
                 + "does NOT delete the column; removing a whole column (its cell in every row + the table "
                 + "grid) is a table-structure change that is out of scope (see the project backlog). Same "
                 + "write/validate/save-or-reject safety as insert_field.")]
    public static ToolResponse ClearCellText(
        [Description("Absolute path to the .docx layout file.")] string layoutPath,
        [Description("0-based table index in document order (see get_layout_info's tables[]).")] int tableIndex,
        [Description("0-based row index within the table.")] int row,
        [Description("0-based cell/column index within the row.")] int col,
        [Description("Which OOXML part the cell is in: 'body' (default), 'header', or 'footer'.")] string layoutPart = "body",
        [Description("Only used when layoutPart is 'header'/'footer': a specific part file name (e.g. "
                     + "'header2.xml'); omit to use the first header/footer part.")] string? partName = null)
    {
        return GuardCellEdit(layoutPath, doc =>
        {
            var location = BuildLocation("tableCell", null, tableIndex, row, col, null, layoutPart, partName);
            return CellTextEditor.ClearCellText(doc, location);
        });
    }
}
