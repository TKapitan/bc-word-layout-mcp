using System.Collections.Frozen;
using System.Reflection;
using System.Text.Json;
using Microsoft.Extensions.DependencyInjection;
using ModelContextProtocol;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace BcWordLayout.McpHost;

/// <summary>
/// The call boundary: the one place between the MCP SDK receiving a <c>tools/call</c> request and a tool
/// body running. Every tool BODY already returns the <see cref="ToolResponse"/> envelope (its exceptions are
/// translated by <c>Tools.ToolGuards.Guard</c>), but three things happen before, or around, a body: the SDK
/// looks the tool up by name, binds the JSON arguments onto the method's typed parameters, and serializes
/// the result. A failure in any of those surfaces here as an exception thrown out of the SDK's handler, and
/// <see cref="Filter"/> turns it into the same envelope shape every other error takes, so an agent sees
/// <c>{ok:false,error:{code,message,hint}}</c> for a malformed call too, never the SDK's bare
/// <c>"An error occurred invoking '…'."</c> text.
/// </summary>
/// <remarks>
/// <para>
/// WHAT REACHES THE FILTER, measured against <c>ModelContextProtocol</c> 2.2.0 (and identical on the 1.4.1
/// the server shipped with at v1.0.0 — the 1.x→2.x upgrade changed none of it; pinned by
/// <c>McpHostPipelineTests</c>, which drives this exact pipeline over the SDK's stream transport):
/// <list type="bullet">
/// <item>An unknown tool NAME — <see cref="RequestContext{TParams}.MatchedPrimitive"/> is null and the
/// SDK's handler would throw <see cref="McpProtocolException"/> ("Unknown tool"). Answered here, before
/// <c>next</c>, as <c>invalid_argument</c> naming the tool: it is a mistake in the call, and keeping the
/// code list small matters more than a dedicated code for it.</item>
/// <item>A missing required argument — the SDK binder throws <see cref="ArgumentException"/> (ParamName
/// <c>arguments</c>, message naming the parameter); mapped to <c>invalid_argument</c>, with the missing
/// names recomputed from the tool's own input schema rather than parsed out of the SDK's message.</item>
/// <item>A wrong-typed argument (<c>"abc"</c> or <c>1.5</c> for an integer, <c>"true"</c> for a boolean,
/// <c>null</c> for a non-nullable number/boolean) — the binder throws <see cref="JsonException"/> whose
/// path is just <c>$</c>, with no parameter name. <see cref="DescribeTypeMismatch"/> recovers the parameter
/// by re-reading each supplied argument into its parameter's CLR type with the SAME serializer options the
/// binder uses (<see cref="McpJsonUtilities.DefaultOptions"/>), so the diagnosis cannot disagree with the
/// binding it explains.</item>
/// <item>An explicit <c>null</c> for a string parameter is NOT a binding failure: the binder passes it
/// through, so each tool's own parser handles it (an optional enum-like string's null means its documented
/// default; a required one is refused with <c>invalid_argument</c>).</item>
/// </list>
/// A request whose <c>arguments</c> is not a JSON object, or that has no <c>name</c>, fails while the SDK
/// deserializes the JSON-RPC params — before any filter exists to wrap it — and stays a JSON-RPC error
/// with no tool result. That is a protocol-level malformation no MCP client library produces.
/// </para>
/// <para>
/// NO SEPARATE <c>invalid_request</c> CODE: an earlier version of this filter rewrapped an
/// <c>isError:true</c> result with a non-JSON first block as <c>invalid_request</c>, a code no document
/// listed. That result shape can only come from an SDK layer below the filter swallowing a binding failure
/// into its generic text (every tool body returns an envelope, and no tool here sets <c>isError</c> itself),
/// so it is the same mistake the filter reports as <c>invalid_argument</c> when it catches the exception
/// directly. Folding it into that code keeps the documented list complete; the branch stays as
/// belt-and-braces against a future SDK that swallows instead of throws.
/// </para>
/// </remarks>
internal static class ToolCallBoundary
{
    /// <summary>
    /// Registers this server's tools and the call-tool filter. Shared by <c>Program.cs</c> and the host-level
    /// tests, so the tests exercise the pipeline the shipped server runs rather than a copy of it. The
    /// assembly is passed explicitly: the parameterless <c>WithToolsFromAssembly()</c> scans the CALLING
    /// assembly, which for a test host is the test assembly.
    /// </summary>
    internal static IMcpServerBuilder WithBcWordLayoutTools(this IMcpServerBuilder builder) =>
        builder
            .WithToolsFromAssembly(typeof(ToolCallBoundary).Assembly)
            .WithRequestFilters(filters => filters.AddCallToolFilter(Filter));

    /// <summary>
    /// Every <c>[McpServerTool]</c> method in this assembly, by tool name — the same set
    /// <see cref="WithBcWordLayoutTools"/> registers. Used only to recover a wrong-typed argument's CLR
    /// parameter type (see <see cref="DescribeTypeMismatch"/>); the tool's NAMES as an agent sees them come
    /// from its input schema instead.
    /// </summary>
    private static readonly FrozenDictionary<string, MethodInfo> ToolMethods =
        typeof(ToolCallBoundary).Assembly.GetTypes()
            .Where(t => t.GetCustomAttribute<McpServerToolTypeAttribute>() is not null)
            .SelectMany(t => t.GetMethods(BindingFlags.Public | BindingFlags.Static))
            .Select(m => (Method: m, m.GetCustomAttribute<McpServerToolAttribute>()?.Name))
            .Where(x => x.Name is not null)
            .ToFrozenDictionary(x => x.Name!, x => x.Method, StringComparer.Ordinal);

    /// <summary>The call-tool filter — see this class's remarks for what reaches it and how each case maps.</summary>
    internal static McpRequestHandler<CallToolRequestParams, CallToolResult> Filter(
        McpRequestHandler<CallToolRequestParams, CallToolResult> next) => async (context, cancellationToken) =>
    {
        var toolName = context.Params?.Name ?? string.Empty;
        var tool = context.MatchedPrimitive as McpServerTool
            ?? (context.Server.ServerOptions.ToolCollection is { } tools && tools.TryGetPrimitive(toolName, out var found)
                ? found
                : null);

        if (tool is null)
        {
            return EnvelopeError(
                "invalid_argument",
                $"Unknown tool '{toolName}'.",
                "This server has no tool by that name. Call tools/list for the exact names it exposes - they "
                + "are snake_case, e.g. get_layout_info, list_dataset_fields, insert_field - and re-send the "
                + "call with one of them.");
        }

        var arguments = context.Params?.Arguments;

        try
        {
            var result = await next(context, cancellationToken);

            // Belt-and-braces (see the class remarks): an isError result whose first block is not JSON can
            // only be an SDK layer below this filter having swallowed a binding failure into its generic text.
            if (result.IsError == true
                && result.Content.OfType<TextContentBlock>().FirstOrDefault() is { } text
                && !text.Text.TrimStart().StartsWith('{'))
            {
                return EnvelopeError("invalid_argument", text.Text, SchemaHint(tool));
            }

            return result;
        }
        catch (JsonException ex)
        {
            // Only the binder reads the caller's JSON before a body runs, so a JsonException is normally a
            // wrong-typed argument. If every supplied argument nevertheless reads cleanly into its parameter's
            // type, the exception came from somewhere else (e.g. serializing the result) and is a real fault.
            return DescribeTypeMismatch(tool, arguments)
                ?? EnvelopeError("internal_error", ex.Message, InternalErrorHint);
        }
        catch (ArgumentException ex)
        {
            // Every tool body is Guard-wrapped, so an ArgumentException escaping to here is the binder's own:
            // in practice a required argument that is absent from the call.
            var missing = MissingRequiredArguments(tool, arguments);
            return EnvelopeError(
                "invalid_argument",
                missing.Count > 0
                    ? $"{tool.ProtocolTool.Name} is missing required argument(s): "
                      + string.Join(", ", missing.Select(m => $"'{m}'")) + "."
                    : ex.Message,
                SchemaHint(tool));
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            return EnvelopeError("internal_error", ex.Message, InternalErrorHint);
        }
    };

    private const string InternalErrorHint =
        "This is an unexpected failure outside the tool body, not a problem with your arguments; retrying the "
        + "same call is unlikely to help. If it persists, it likely indicates a bug in the server and is worth "
        + "reporting together with the full error message above.";

    /// <summary>
    /// The hint for a call whose argument SET is wrong (missing required ones): lists the tool's required
    /// parameters straight from its input schema, so the agent does not need a second tools/list round trip.
    /// </summary>
    private static string SchemaHint(McpServerTool tool)
    {
        var required = RequiredParameters(tool);
        return $"Re-send the call with every required argument. {tool.ProtocolTool.Name} requires: "
            + (required.Count > 0 ? string.Join(", ", required) : "(none)")
            + " - see its input schema (tools/list) for each parameter's type and the optional ones.";
    }

    /// <summary>The tool's <c>required</c> list, read from the same input schema tools/list advertises.</summary>
    private static List<string> RequiredParameters(McpServerTool tool) =>
        tool.ProtocolTool.InputSchema.TryGetProperty("required", out var required)
        && required.ValueKind == JsonValueKind.Array
            ? required.EnumerateArray()
                .Where(e => e.ValueKind == JsonValueKind.String)
                .Select(e => e.GetString()!)
                .ToList()
            : [];

    private static List<string> MissingRequiredArguments(McpServerTool tool, IDictionary<string, JsonElement>? arguments) =>
        RequiredParameters(tool).Where(name => arguments is null || !arguments.ContainsKey(name)).ToList();

    /// <summary>
    /// Recovers WHICH argument failed binding: re-reads every supplied argument into its parameter's CLR type
    /// with the binder's own serializer options and reports the first that does not convert, naming the
    /// parameter, the JSON shape it needs, and what the call actually sent. Null when no supplied argument
    /// fails — the caller then treats the <see cref="JsonException"/> as a genuine internal fault.
    /// </summary>
    private static CallToolResult? DescribeTypeMismatch(McpServerTool tool, IDictionary<string, JsonElement>? arguments)
    {
        if (arguments is null || !ToolMethods.TryGetValue(tool.ProtocolTool.Name, out var method))
        {
            return null;
        }

        foreach (var parameter in method.GetParameters())
        {
            if (parameter.Name is null || !arguments.TryGetValue(parameter.Name, out var value))
            {
                continue;
            }

            try
            {
                JsonSerializer.Deserialize(value, parameter.ParameterType, McpJsonUtilities.DefaultOptions);
            }
            catch (JsonException)
            {
                var expected = DescribeExpectedJson(parameter.ParameterType);
                return EnvelopeError(
                    "invalid_argument",
                    $"Argument '{parameter.Name}' of {tool.ProtocolTool.Name} must be {expected}, but the call "
                    + $"sent {DescribeSentJson(value)}.",
                    $"Re-send the call with '{parameter.Name}' as {expected} (see the tool's input schema via "
                    + "tools/list). Numbers and booleans go unquoted (5, true - not \"5\", \"true\"), an integer "
                    + "parameter takes whole numbers only, and to get an optional parameter's default, omit it "
                    + "instead of sending null.");
            }
        }

        return null;
    }

    private static string DescribeExpectedJson(Type type)
    {
        var underlying = Nullable.GetUnderlyingType(type);
        var core = (underlying ?? type) switch
        {
            var t when t == typeof(int) || t == typeof(long) => "an integer",
            var t when t == typeof(double) || t == typeof(float) || t == typeof(decimal) => "a number",
            var t when t == typeof(bool) => "a boolean (true or false)",
            var t when t == typeof(string) => "a string",
            var t => $"a JSON value of type {t.Name}",
        };

        return underlying is null ? core : core + " or null";
    }

    private static string DescribeSentJson(JsonElement value)
    {
        var raw = value.GetRawText();
        var shown = raw.Length <= 60 ? raw : raw[..57] + "...";
        return value.ValueKind switch
        {
            JsonValueKind.String => $"the string {shown}",
            JsonValueKind.Number => $"the number {shown}",
            JsonValueKind.True or JsonValueKind.False => $"the boolean {shown}",
            JsonValueKind.Null => "null",
            JsonValueKind.Array => "an array",
            JsonValueKind.Object => "an object",
            _ => shown,
        };
    }

    /// <summary>
    /// A failing tool result carrying the envelope, serialized with the same options the SDK applies to every
    /// tool's returned <see cref="ToolResponse"/> — byte-for-byte the shape an agent parses everywhere else.
    /// <c>isError</c> is set because the call never produced a tool result at all.
    /// </summary>
    private static CallToolResult EnvelopeError(string code, string message, string hint) => new()
    {
        IsError = true,
        Content =
        [
            new TextContentBlock
            {
                Text = JsonSerializer.Serialize(ToolResponse.Failure(code, message, hint), McpJsonUtilities.DefaultOptions),
            },
        ],
    };
}
