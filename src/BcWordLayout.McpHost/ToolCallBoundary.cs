using System.Collections.Frozen;
using System.Reflection;
using System.Runtime.CompilerServices;
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
/// — or, inside an object or array argument, the member or item (<c>columns[1].width</c>) — by re-reading each
/// supplied argument into its parameter's CLR type with the SAME serializer options the binder uses
/// (<see cref="McpJsonUtilities.DefaultOptions"/>), so the diagnosis cannot disagree with the binding it
/// explains.</item>
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
/// <para>
/// TWO CHECKS THE SDK DOES NOT MAKE, both run before <c>next</c> so a refused call never reaches the binder
/// or a tool body (ADR-0003, refuse rather than guess; GitHub issue #90):
/// <list type="bullet">
/// <item>UNKNOWN ARGUMENT NAMES. The generated input schemas do not set <c>additionalProperties:false</c> and
/// the binder ignores keys it has no parameter for, so a misspelled OPTIONAL argument (<c>layoutpart</c>,
/// <c>atcolumn</c>, remove_column's <c>column</c> sent to set_cell_format, whose parameter is <c>col</c>)
/// silently fell back to its default and the call "succeeded" doing something the agent did not ask for.
/// Every key must now be a property of the tool's input schema; the refusal names the closest real
/// parameter (<see cref="ClosestName"/>) and lists them all. The same holds for the members of an object
/// argument and of the objects in an array argument.</item>
/// <item>RELATIVE PATHS. Every path parameter is documented as absolute, but a relative one used to resolve
/// against the server process's working directory, which the agent cannot see (and <c>create_layout</c>
/// happily wrote there). A non-blank string for a parameter in <see cref="FilePathParameters"/> must now be
/// <see cref="Path.IsPathFullyQualified(string)"/>; a blank one passes through to the tool's own "must not
/// be empty"/file-not-found handling, which names it more precisely.</item>
/// </list>
/// </para>
/// </remarks>
internal static class ToolCallBoundary
{
    /// <summary>
    /// The parameters, across every tool, that take a FILE-SYSTEM path — keyed by name, because each name means
    /// the same thing on every tool that has it. An explicit set rather than a naming heuristic, because the
    /// obvious heuristic is wrong here: <c>insert_column</c>'s <c>dataPath</c>/<c>headerLabelPath</c> end in
    /// "Path" but are DATASET paths (<c>/Header/Line/Discount_Line</c>), which are never fully qualified file
    /// paths. <c>McpHostPipelineTests</c> fails if a tool gains a path-like string parameter (a name ending in
    /// Path/Source/Dir/File/Folder, or a description promising an absolute path) that is neither listed here
    /// nor explicitly classified as a dataset path there.
    /// </summary>
    internal static readonly FrozenSet<string> FilePathParameters = new[]
    {
        "layoutPath", "source", "schemaSource", "newSchemaSource", "templatePath", "outputPath",
        "dataOverridesPath", "outputDir", "pdfPath",
    }.ToFrozenSet(StringComparer.Ordinal);

    /// <summary>
    /// Registers this server's tools and the call-tool filter. Shared by <c>Program.cs</c> and the host-level
    /// tests, so the tests exercise the pipeline the shipped server runs rather than a copy of it. The
    /// assembly is passed explicitly: the parameterless <c>WithToolsFromAssembly()</c> scans the CALLING
    /// assembly, which for a test host is the test assembly.
    /// </summary>
    internal static IMcpServerBuilder WithBcWordLayoutTools(this IMcpServerBuilder builder) =>
        builder
            .WithToolsFromAssembly(typeof(ToolCallBoundary).Assembly)
            .WithListConstraintsOnItems()
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
            var toolNames = context.Server.ServerOptions.ToolCollection?.Select(t => t.ProtocolTool.Name).ToList() ?? [];
            return EnvelopeError(
                ErrorCodes.InvalidArgument,
                $"Unknown tool '{toolName}'."
                + (ClosestName(toolName, toolNames) is { } closestTool ? $" Did you mean '{closestTool}'?" : string.Empty),
                "This server has no tool by that name. Call tools/list for the exact names it exposes - they "
                + "are snake_case, e.g. get_layout_info, list_dataset_fields, insert_field - and re-send the "
                + "call with one of them.");
        }

        var arguments = context.Params?.Arguments;
        DecodeStructuredArgumentsSentAsStrings(tool, arguments);
        if (RejectUnknownArguments(tool, arguments) is { } unknownArguments)
        {
            return unknownArguments;
        }

        if (RejectRelativePaths(tool, arguments) is { } relativePaths)
        {
            return relativePaths;
        }

        try
        {
            var result = await next(context, cancellationToken);

            // Belt-and-braces (see the class remarks): an isError result whose first block is not JSON can
            // only be an SDK layer below this filter having swallowed a binding failure into its generic text.
            if (result.IsError == true
                && result.Content.OfType<TextContentBlock>().FirstOrDefault() is { } text
                && !text.Text.TrimStart().StartsWith('{'))
            {
                return EnvelopeError(ErrorCodes.InvalidArgument, text.Text, SchemaHint(tool));
            }

            return result;
        }
        catch (JsonException ex)
        {
            // Only the binder reads the caller's JSON before a body runs, so a JsonException is normally a
            // wrong-typed argument. If every supplied argument nevertheless reads cleanly into its parameter's
            // type, the exception came from somewhere else (e.g. serializing the result) and is a real fault.
            return DescribeTypeMismatch(tool, arguments)
                ?? EnvelopeError(ErrorCodes.InternalError, ex.Message, InternalErrorHint);
        }
        catch (ArgumentException ex)
        {
            // Every tool body is Guard-wrapped, so an ArgumentException escaping to here is the binder's own:
            // in practice a required argument that is absent from the call.
            var missing = MissingRequiredArguments(tool, arguments);
            return EnvelopeError(
                ErrorCodes.InvalidArgument,
                missing.Count > 0
                    ? $"{tool.ProtocolTool.Name} is missing required argument(s): "
                      + string.Join(", ", missing.Select(m => $"'{m}'")) + "."
                    : ex.Message,
                SchemaHint(tool));
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            return EnvelopeError(ErrorCodes.InternalError, ex.Message, InternalErrorHint);
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

    /// <summary>Every parameter name in the tool's input schema, in schema (= method signature) order.</summary>
    private static List<string> SchemaParameters(McpServerTool tool) =>
        tool.ProtocolTool.InputSchema.TryGetProperty("properties", out var properties)
        && properties.ValueKind == JsonValueKind.Object
            ? properties.EnumerateObject().Select(p => p.Name).ToList()
            : [];

    /// <summary>
    /// One OBJECT or ARRAY parameter, read from the same input schema tools/list advertises: whether it is an array,
    /// and the member names its objects take (the object's own, or each item's for an array of objects; none for an
    /// array of plain values). The binder reads these as structured JSON, so they get the checks a flat argument
    /// does not need: the <c>format</c> object (GitHub issue #15) and the table tools' arrays (GitHub issue #133).
    /// </summary>
    private sealed record StructuredParameter(string Name, bool IsArray, IReadOnlyList<string> Members);

    private static List<StructuredParameter> StructuredParameters(McpServerTool tool)
    {
        var structured = new List<StructuredParameter>();
        if (!tool.ProtocolTool.InputSchema.TryGetProperty("properties", out var properties)
            || properties.ValueKind != JsonValueKind.Object)
        {
            return structured;
        }

        foreach (var parameter in properties.EnumerateObject().Where(p => p.Value.ValueKind == JsonValueKind.Object))
        {
            if (HasType(parameter.Value, "array"))
            {
                var items = parameter.Value.TryGetProperty("items", out var itemSchema) ? itemSchema : default;
                structured.Add(new(
                    parameter.Name,
                    IsArray: true,
                    items.ValueKind == JsonValueKind.Object && HasType(items, "object") ? MemberNames(items) : []));
            }
            else if (HasType(parameter.Value, "object"))
            {
                structured.Add(new(parameter.Name, IsArray: false, MemberNames(parameter.Value)));
            }
        }

        return structured;
    }

    private static List<string> MemberNames(JsonElement schema) =>
        schema.TryGetProperty("properties", out var members) && members.ValueKind == JsonValueKind.Object
            ? members.EnumerateObject().Select(m => m.Name).ToList()
            : [];

    private static bool HasType(JsonElement schema, string type) =>
        schema.TryGetProperty("type", out var declared)
        && (declared.ValueKind == JsonValueKind.String
            ? declared.GetString() == type
            : declared.ValueKind == JsonValueKind.Array && declared.EnumerateArray().Any(t => t.GetString() == type));

    /// <summary>
    /// Replaces a structured argument that arrived as a STRING holding its JSON (<c>"{\"bold\":true}"</c> instead of
    /// <c>{"bold":true}</c>, <c>"[1500,6000]"</c> instead of <c>[1500,6000]</c>) with the object or array itself.
    /// Models do send nested JSON that way, and the string is unambiguous, so reading it is not a guess (ADR-0003);
    /// refusing it would only cost the agent a round trip. Any other string - a comma-separated list included - is
    /// left alone for the binder to reject, which <see cref="DescribeTypeMismatch"/> then explains.
    /// </summary>
    private static void DecodeStructuredArgumentsSentAsStrings(McpServerTool tool, IDictionary<string, JsonElement>? arguments)
    {
        if (arguments is null || arguments.Count == 0)
        {
            return;
        }

        foreach (var parameter in StructuredParameters(tool))
        {
            var (opener, kind) = parameter.IsArray ? ('[', JsonValueKind.Array) : ('{', JsonValueKind.Object);
            if (!arguments.TryGetValue(parameter.Name, out var value)
                || value.ValueKind != JsonValueKind.String
                || value.GetString() is not { } text
                || !text.TrimStart().StartsWith(opener))
            {
                continue;
            }

            try
            {
                using var parsed = JsonDocument.Parse(text);
                if (parsed.RootElement.ValueKind == kind)
                {
                    arguments[parameter.Name] = parsed.RootElement.Clone();
                }
            }
            catch (JsonException)
            {
                // Not JSON after all: the binder refuses it, and DescribeTypeMismatch names the argument.
            }
        }
    }

    /// <summary>
    /// Refuses a call carrying any argument key that is not a property of the tool's input schema (matched
    /// ordinally — argument names are case-sensitive, so <c>layoutpart</c> is unknown, not <c>layoutPart</c>),
    /// or an object — an object argument, or an object inside an array argument — carrying a member its schema
    /// does not list (<c>format.size</c>, <c>columns[1].widht</c>; the binder would drop it just as silently).
    /// Null when every key is known. See the class remarks for why an ignored key is worse than a refusal.
    /// </summary>
    private static CallToolResult? RejectUnknownArguments(McpServerTool tool, IDictionary<string, JsonElement>? arguments)
    {
        if (arguments is null || arguments.Count == 0)
        {
            return null;
        }

        var known = SchemaParameters(tool);
        var unknown = arguments.Keys.Where(key => !known.Contains(key, StringComparer.Ordinal)).ToList();
        if (unknown.Count > 0)
        {
            var described = unknown.Select(key =>
                ClosestName(key, known) is { } closest ? $"'{key}' (did you mean '{closest}'?)" : $"'{key}'");
            return EnvelopeError(
                ErrorCodes.InvalidArgument,
                $"{tool.ProtocolTool.Name} has no argument named {string.Join(", ", described)}; the call was refused "
                + "rather than run as if the argument had not been sent.",
                "Argument names are case-sensitive and must match the tool's input schema exactly. An unknown name is "
                + "refused rather than ignored, because a misspelled optional argument would otherwise silently fall "
                + $"back to its default. {tool.ProtocolTool.Name} takes: {string.Join(", ", known)}.");
        }

        foreach (var parameter in StructuredParameters(tool).Where(p => p.Members.Count > 0))
        {
            if (!arguments.TryGetValue(parameter.Name, out var value))
            {
                continue;
            }

            IEnumerable<(JsonElement Value, string Path)> objects = parameter.IsArray
                ? value.ValueKind == JsonValueKind.Array
                    ? value.EnumerateArray().Select((item, index) => (item, $"{parameter.Name}[{index}]"))
                    : []
                : [(value, parameter.Name)];
            foreach (var (candidate, path) in objects.Where(o => o.Value.ValueKind == JsonValueKind.Object))
            {
                var unknownMembers = candidate.EnumerateObject()
                    .Select(m => m.Name)
                    .Where(member => !parameter.Members.Contains(member, StringComparer.Ordinal))
                    .ToList();
                if (unknownMembers.Count == 0)
                {
                    continue;
                }

                var described = unknownMembers.Select(member =>
                    ClosestName(member, parameter.Members) is { } closest
                        ? $"'{member}' (did you mean '{closest}'?)"
                        : $"'{member}'");
                return EnvelopeError(
                    ErrorCodes.InvalidArgument,
                    $"{tool.ProtocolTool.Name}'s {path} has no member named {string.Join(", ", described)}; the call "
                    + "was refused rather than run as if the member had not been sent.",
                    "Member names are case-sensitive and must match the input schema exactly. "
                    + (parameter.IsArray ? $"Each item of {parameter.Name}" : parameter.Name)
                    + $" takes: {string.Join(", ", parameter.Members)}.");
            }
        }

        return null;
    }

    /// <summary>
    /// Refuses a call passing a relative (or drive-/root-relative, e.g. <c>C:Report.docx</c>,
    /// <c>\Report.docx</c>) path to any <see cref="FilePathParameters"/> parameter, naming each one. Blank and
    /// non-string values pass through: the tool's own validation (blank) or the binder (wrong type) reports
    /// those more precisely. Null when every path argument is fully qualified.
    /// </summary>
    private static CallToolResult? RejectRelativePaths(McpServerTool tool, IDictionary<string, JsonElement>? arguments)
    {
        if (arguments is null)
        {
            return null;
        }

        var relative = arguments
            .Where(a => FilePathParameters.Contains(a.Key)
                        && a.Value.ValueKind == JsonValueKind.String
                        && a.Value.GetString() is { } path
                        && !string.IsNullOrWhiteSpace(path)
                        && !Path.IsPathFullyQualified(path))
            .Select(a => $"{a.Key} ('{a.Value.GetString()}')")
            .ToList();
        if (relative.Count == 0)
        {
            return null;
        }

        return EnvelopeError(
            ErrorCodes.InvalidArgument,
            $"{tool.ProtocolTool.Name} needs absolute paths, but {string.Join(" and ", relative)} "
            + (relative.Count == 1 ? "is" : "are")
            + " relative: it would resolve against the server process's working directory, which the caller "
            + "cannot see.",
            "Pass a fully qualified path - a drive letter and root (C:\\Work\\Report.docx) or a UNC path "
            + "(\\\\server\\share\\Report.docx) - built from your own workspace root. A relative path, and a "
            + "drive- or root-relative one such as 'C:Report.docx' or '\\Report.docx', is refused rather than "
            + "resolved against the server's current directory.");
    }

    /// <summary>
    /// The candidate an agent most plausibly MEANT by <paramref name="name"/>, or null when nothing is close:
    /// first a case-insensitive match (<c>layoutpart</c> → <c>layoutPart</c>); then the nearest candidate by
    /// case-insensitive edit distance, if within a quarter of the name's length (at least 1 —
    /// <c>tabelIndex</c> → <c>tableIndex</c>); then the longest candidate of 3+ characters that is a prefix of
    /// the name or vice versa (<c>column</c> → <c>col</c>). Deliberately conservative: a far-fetched
    /// suggestion is worse than none, and every refusal lists the full valid set anyway.
    /// </summary>
    internal static string? ClosestName(string name, IReadOnlyCollection<string> candidates)
    {
        if (string.IsNullOrEmpty(name) || candidates.Count == 0)
        {
            return null;
        }

        if (candidates.FirstOrDefault(c => string.Equals(c, name, StringComparison.OrdinalIgnoreCase)) is { } exact)
        {
            return exact;
        }

        var lowered = name.ToLowerInvariant();
        var (nearest, distance) = candidates
            .Select(c => (Candidate: c, Distance: EditDistance(lowered, c.ToLowerInvariant())))
            .MinBy(x => x.Distance);
        if (distance <= Math.Max(1, name.Length / 4))
        {
            return nearest;
        }

        return candidates
            .Where(c => Math.Min(c.Length, name.Length) >= 3
                        && (name.StartsWith(c, StringComparison.OrdinalIgnoreCase)
                            || c.StartsWith(name, StringComparison.OrdinalIgnoreCase)))
            .MaxBy(c => c.Length);
    }

    /// <summary>Levenshtein distance (insert/delete/substitute, each cost 1) — two-row dynamic programming.</summary>
    private static int EditDistance(string a, string b)
    {
        var previous = new int[b.Length + 1];
        var current = new int[b.Length + 1];
        for (var j = 0; j <= b.Length; j++)
        {
            previous[j] = j;
        }

        for (var i = 1; i <= a.Length; i++)
        {
            current[0] = i;
            for (var j = 1; j <= b.Length; j++)
            {
                var substitution = previous[j - 1] + (a[i - 1] == b[j - 1] ? 0 : 1);
                current[j] = Math.Min(Math.Min(previous[j] + 1, current[j - 1] + 1), substitution);
            }

            (previous, current) = (current, previous);
        }

        return previous[b.Length];
    }

    /// <summary>
    /// Recovers WHICH argument failed binding: re-reads every supplied argument into its parameter's CLR type
    /// with the binder's own serializer options and reports the first that does not convert, naming the
    /// parameter, the JSON shape it needs, and what the call actually sent. For an object or array argument it
    /// names the member or item at fault instead (<see cref="DescribeInnerMismatch"/>). Null when no supplied
    /// argument fails — the caller then treats the <see cref="JsonException"/> as a genuine internal fault.
    /// </summary>
    private static CallToolResult? DescribeTypeMismatch(McpServerTool tool, IDictionary<string, JsonElement>? arguments)
    {
        if (arguments is null || !ToolMethods.TryGetValue(tool.ProtocolTool.Name, out var method))
        {
            return null;
        }

        foreach (var parameter in method.GetParameters())
        {
            if (parameter.Name is null
                || !arguments.TryGetValue(parameter.Name, out var value)
                || Reads(value, parameter.ParameterType))
            {
                continue;
            }

            if (DescribeInnerMismatch(tool, parameter.Name, parameter.ParameterType, value) is { } inner)
            {
                return inner;
            }

            var expected = DescribeExpectedJson(parameter.ParameterType, parameter.HasDefaultValue);
            return EnvelopeError(
                ErrorCodes.InvalidArgument,
                $"Argument '{parameter.Name}' of {tool.ProtocolTool.Name} must be {expected}, but the call "
                + $"sent {DescribeSentJson(value)}.",
                $"Re-send the call with '{parameter.Name}' as {expected} (see the tool's input schema via "
                + "tools/list). Numbers and booleans go unquoted (5, true - not \"5\", \"true\"), an integer "
                + "parameter takes whole numbers only, "
                + (ElementType(parameter.ParameterType) is not null && value.ValueKind == JsonValueKind.String
                    ? "a list is a JSON array ([1500, 6000]), not a comma-separated string, "
                    : string.Empty)
                + "and to get an optional parameter's default, omit it instead of sending null.");
        }

        return null;
    }

    /// <summary>
    /// For an object or array value that failed binding: the first member or item whose own value does not read,
    /// found recursively and reported by its path (<c>format.bold</c>, <c>widths[2]</c>, <c>columns[1].width</c>)
    /// with the JSON shape it needs — or, for an object, the first <c>required</c> member it lacks. Null when every
    /// part reads on its own, so the caller describes the value as a whole (it is the wrong shape altogether).
    /// </summary>
    private static CallToolResult? DescribeInnerMismatch(McpServerTool tool, string path, Type type, JsonElement value)
    {
        if (value.ValueKind == JsonValueKind.Array && ElementType(type) is { } element)
        {
            var index = 0;
            foreach (var item in value.EnumerateArray())
            {
                var itemPath = $"{path}[{index++}]";
                if (!Reads(item, element))
                {
                    return DescribeInnerMismatch(tool, itemPath, element, item)
                        ?? DescribePartMismatch(tool, $"Item '{itemPath}'", itemPath, DescribeExpectedJson(element, nullable: false), item);
                }
            }

            return null;
        }

        if (value.ValueKind != JsonValueKind.Object || !IsObjectType(type))
        {
            return null;
        }

        var naming = McpJsonUtilities.DefaultOptions.PropertyNamingPolicy;
        var properties = type.GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .ToDictionary(p => naming?.ConvertName(p.Name) ?? p.Name, StringComparer.Ordinal);

        foreach (var (name, property) in properties)
        {
            if (property.IsDefined(typeof(RequiredMemberAttribute)) && !value.TryGetProperty(name, out _))
            {
                return EnvelopeError(
                    ErrorCodes.InvalidArgument,
                    $"'{path}' of {tool.ProtocolTool.Name} has no '{name}', which it requires.",
                    $"Re-send the call with '{path}.{name}' set (see the tool's input schema via tools/list).");
            }
        }

        foreach (var member in value.EnumerateObject())
        {
            if (properties.TryGetValue(member.Name, out var property) && !Reads(member.Value, property.PropertyType))
            {
                var memberPath = $"{path}.{member.Name}";
                return DescribeInnerMismatch(tool, memberPath, property.PropertyType, member.Value)
                    ?? DescribePartMismatch(
                        tool,
                        $"Member '{memberPath}'",
                        memberPath,
                        DescribeExpectedJson(property.PropertyType, nullable: !property.IsDefined(typeof(RequiredMemberAttribute))),
                        member.Value);
            }
        }

        return null;
    }

    private static CallToolResult DescribePartMismatch(McpServerTool tool, string what, string path, string expected, JsonElement value) =>
        EnvelopeError(
            ErrorCodes.InvalidArgument,
            $"{what} of {tool.ProtocolTool.Name} must be {expected}, but the call sent {DescribeSentJson(value)}.",
            $"Re-send the call with '{path}' as {expected} (see the tool's input schema via tools/list). Numbers and "
            + "booleans go unquoted (8, true - not \"8\", \"true\"); omit an optional member to take its default.");

    /// <summary>Whether <paramref name="value"/> reads into <paramref name="type"/> with the binder's own options.</summary>
    private static bool Reads(JsonElement value, Type type)
    {
        try
        {
            JsonSerializer.Deserialize(value, type, McpJsonUtilities.DefaultOptions);
            return true;
        }
        catch (JsonException)
        {
            return false;
        }
    }

    /// <summary>The item type of a list-shaped parameter type (<c>IReadOnlyList&lt;int&gt;</c>, an array), else null.</summary>
    private static Type? ElementType(Type type) =>
        type == typeof(string)
            ? null
            : type.IsArray
                ? type.GetElementType()
                : type.IsGenericType && type.GetGenericArguments() is [var item]
                  && typeof(IEnumerable<>).MakeGenericType(item).IsAssignableFrom(type)
                    ? item
                    : null;

    private static bool IsObjectType(Type type) => type.IsClass && type != typeof(string) && ElementType(type) is null;

    /// <summary>
    /// The JSON shape <paramref name="type"/> binds from, in words. <paramref name="nullable"/> says whether null
    /// is accepted too (an optional argument or member); a nullable value type always says so.
    /// </summary>
    private static string DescribeExpectedJson(Type type, bool nullable)
    {
        var underlying = Nullable.GetUnderlyingType(type);
        var core = DescribeJsonShape(underlying ?? type);

        // A string is never described as taking null: an explicit null for one is the tool's documented default.
        return underlying is not null || (nullable && type != typeof(string) && !type.IsValueType) ? core + " or null" : core;
    }

    private static string DescribeJsonShape(Type type) => type switch
    {
        var t when t == typeof(int) || t == typeof(long) => "an integer",
        var t when t == typeof(double) || t == typeof(float) || t == typeof(decimal) => "a number",
        var t when t == typeof(bool) => "a boolean (true or false)",
        var t when t == typeof(string) => "a string",
        var t when ElementType(t) is { } item => "a JSON array of " + DescribeJsonShape(item) switch
        {
            "an integer" => "integers",
            "a number" => "numbers",
            "a string" => "strings",
            "a JSON object" => "objects",
            var shape => $"items that are each {shape}",
        },
        var t when t.IsClass => "a JSON object",
        var t => $"a JSON value of type {t.Name}",
    };

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
