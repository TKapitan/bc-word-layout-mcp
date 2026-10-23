using System.ComponentModel.DataAnnotations;

namespace BcWordLayout.McpHost.Tools;

// The values each enumerated string argument takes, written once and published in the tool's input schema: the SDK
// turns an [AllowedValues] into the schema's "enum", so an agent reads the choices from tools/list instead of from
// prose (GitHub issue #133). Schema only — the SDK does not validate against it. Each parser keeps reading its
// argument case-insensitively and refuses an unknown value with its own hint; AllowedValueAttributesTests holds
// every value listed here to the parser that reads it, so the two cannot drift apart. No list holds null: the SDK
// leaves it out of the enum, and an optional argument is omitted, not sent as null, to take its default.

/// <summary>Every insert tool's <c>locationType</c>.</summary>
internal sealed class LocationTypeValuesAttribute() : AllowedValuesAttribute("documentEnd", "afterControl", "tableCell", "atText");

/// <summary>Every <c>layoutPart</c>.</summary>
internal sealed class LayoutPartValuesAttribute() : AllowedValuesAttribute("body", "header", "footer");

/// <summary>A paragraph alignment: <c>set_cell_text</c>'s and <c>set_text_format</c>'s, and a column's or cell's.</summary>
internal sealed class AlignmentValuesAttribute() : AllowedValuesAttribute("left", "center", "right");

/// <summary>The <c>borders</c> of <c>insert_table</c> and <c>insert_repeater_table</c>.</summary>
internal sealed class BordersValuesAttribute() : AllowedValuesAttribute("none", "headerRule", "grid");

/// <summary><c>set_cell_format</c>'s <c>borderStyle</c>.</summary>
internal sealed class BorderStyleValuesAttribute() : AllowedValuesAttribute("single", "none");

/// <summary><c>set_cell_format</c>'s <c>edges</c>; on that list the values describe its items (see <c>ToolInputSchemas</c>).</summary>
internal sealed class EdgeValuesAttribute() : AllowedValuesAttribute("top", "bottom", "left", "right", "all");

/// <summary><c>insert_column</c>'s <c>mode</c>.</summary>
internal sealed class ColumnModeValuesAttribute() : AllowedValuesAttribute("field", "label", "plainText");

/// <summary><c>set_text_format</c>'s <c>target</c>.</summary>
internal sealed class TargetValuesAttribute() : AllowedValuesAttribute("control", "cell", "row", "text");

/// <summary><c>validate_layout</c>'s <c>level</c>.</summary>
internal sealed class LevelValuesAttribute() : AllowedValuesAttribute("quick", "full");

/// <summary><c>preview_layout</c>'s <c>converter</c>.</summary>
internal sealed class ConverterValuesAttribute() : AllowedValuesAttribute("auto", "word", "libreoffice");
