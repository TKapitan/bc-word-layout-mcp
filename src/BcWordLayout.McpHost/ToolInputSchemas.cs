using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.Extensions.DependencyInjection;
using ModelContextProtocol.Server;

namespace BcWordLayout.McpHost;

/// <summary>
/// Corrects the one thing the SDK's input-schema generation gets wrong for this server's tools: the allowed values
/// and range of a LIST parameter's items.
/// </summary>
/// <remarks>
/// <para>
/// The SDK turns DataAnnotations into schema keywords — <c>[AllowedValues]</c> into <c>enum</c>, <c>[Range]</c> into
/// <c>minimum</c>/<c>maximum</c> — which is how an agent learns, from tools/list alone, that an alignment is
/// <c>left</c>, <c>center</c> or <c>right</c> and that a width is 1 to 31680 twips (GitHub issue #133). On a list
/// parameter (<c>set_cell_format</c>'s <c>edges</c>, <c>set_column_widths</c>' <c>widths</c>) it writes those
/// keywords onto the array node itself, where they can never hold: an array is never equal to <c>"top"</c>, and
/// has no minimum. On a list they mean its items, so this moves them onto <c>items</c>.
/// </para>
/// <para>
/// WHY AFTER THE FACT: ModelContextProtocol 2.2.0 does not apply <c>McpServerToolCreateOptions.SchemaCreateOptions</c>'
/// <c>TransformSchemaNode</c> to a tool's input schema (measured), so the generated schema is corrected once, when
/// the server options are built, on the same <see cref="ModelContextProtocol.Protocol.Tool"/> that tools/list
/// returns and <see cref="ToolCallBoundary"/> reads.
/// </para>
/// </remarks>
internal static class ToolInputSchemas
{
    private static readonly string[] ItemKeywords = ["enum", "minimum", "maximum"];

    /// <summary>Registers the correction for every tool the server has, whichever registration added it.</summary>
    internal static IMcpServerBuilder WithListConstraintsOnItems(this IMcpServerBuilder builder)
    {
        builder.Services.PostConfigure<McpServerOptions>(options =>
        {
            foreach (var tool in options.ToolCollection ?? [])
            {
                tool.ProtocolTool.InputSchema = MoveListConstraintsToItems(tool.ProtocolTool.InputSchema);
            }
        });

        return builder;
    }

    /// <summary>
    /// <paramref name="schema"/> with every <c>enum</c>/<c>minimum</c>/<c>maximum</c> that sits on an array node
    /// moved onto that node's <c>items</c>; unchanged when there is none.
    /// </summary>
    internal static JsonElement MoveListConstraintsToItems(JsonElement schema)
    {
        var root = JsonNode.Parse(schema.GetRawText());
        if (root is null || !Move(root))
        {
            return schema;
        }

        using var document = JsonDocument.Parse(root.ToJsonString());
        return document.RootElement.Clone();
    }

    private static bool Move(JsonNode node)
    {
        if (node is not JsonObject schema)
        {
            return false;
        }

        var moved = false;
        if (IsArraySchema(schema) && schema["items"] is JsonObject items)
        {
            foreach (var keyword in ItemKeywords)
            {
                if (schema[keyword] is { } value)
                {
                    schema.Remove(keyword);
                    items[keyword] = value;
                    moved = true;
                }
            }
        }

        foreach (var child in schema.Select(p => p.Value).OfType<JsonObject>().ToList())
        {
            moved |= Move(child);
        }

        return moved;
    }

    private static bool IsArraySchema(JsonObject schema) => schema["type"] switch
    {
        JsonValue type => type.TryGetValue<string>(out var name) && name == "array",
        JsonArray types => types.Any(t => t is JsonValue v && v.TryGetValue<string>(out var name) && name == "array"),
        _ => false,
    };
}
