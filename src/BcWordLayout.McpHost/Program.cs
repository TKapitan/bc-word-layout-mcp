using System.Text.Json;
using BcWordLayout.Domain;
using BcWordLayout.McpHost;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol;
using ModelContextProtocol.Protocol;

// BC Word Layout MCP server (v1: inspect, validate, preview, create/edit/refresh tools).
// stdio JSON-RPC transport via the official ModelContextProtocol C# SDK.
// IMPORTANT: stdout carries the MCP protocol, so ALL logging must go to stderr.

// Optional host-level override of the default label convention (suffix "Lbl" plus the self-scoping
// "Labels" data-item rule — see LabelConvention's remarks): BCWL_LABEL_SUFFIXES, comma-separated
// (e.g. "Lbl,Caption"), and/or BCWL_LABELS_DATA_ITEM (a data-item name retargeting the rule — every
// direct column of a data item with that name is a label regardless of suffix — or "-" to disable it).
// Read directly via Console.Error here rather than the DI logger configured below, since the host isn't
// built yet at this point in the file — still stderr-only, never stdout, so it cannot corrupt the MCP
// JSON-RPC stream.
var rawLabelSuffixes = Environment.GetEnvironmentVariable(EnvironmentConfig.LabelSuffixesVariable);
var rawLabelsDataItem = Environment.GetEnvironmentVariable(EnvironmentConfig.LabelsDataItemVariable);
if (!string.IsNullOrWhiteSpace(rawLabelSuffixes) || !string.IsNullOrWhiteSpace(rawLabelsDataItem))
{
    var parsedLabelSuffixes = EnvironmentConfig.ParseLabelSuffixes(rawLabelSuffixes);
    if (!string.IsNullOrWhiteSpace(rawLabelSuffixes) && parsedLabelSuffixes is null)
    {
        Console.Error.WriteLine(
            $"[bc-word-layout-mcp] Ignoring invalid {EnvironmentConfig.LabelSuffixesVariable} value "
            + $"'{rawLabelSuffixes}'; keeping the default label suffixes (Lbl).");
    }

    var parsedLabelsDataItem = EnvironmentConfig.ParseLabelsDataItemName(rawLabelsDataItem);
    if (!string.IsNullOrWhiteSpace(rawLabelsDataItem) && parsedLabelsDataItem is null)
    {
        Console.Error.WriteLine(
            $"[bc-word-layout-mcp] Ignoring invalid {EnvironmentConfig.LabelsDataItemVariable} value "
            + $"'{rawLabelsDataItem}' (must be a single data-item name, or '-' to disable the rule); "
            + "keeping the default labels data item (Labels).");
    }

    // parsedLabelsDataItem: null = keep the default rule; "" (the "-" sentinel) = disable it — the
    // LabelConvention constructor maps a blank name to a disabled rule; anything else retargets it.
    if (parsedLabelSuffixes is not null || parsedLabelsDataItem is not null)
    {
        LabelConvention.Current = new LabelConvention(
            parsedLabelSuffixes ?? LabelConvention.Default.Suffixes,
            parsedLabelsDataItem ?? LabelConvention.Default.LabelsDataItemName);
        Console.Error.WriteLine(
            "[bc-word-layout-mcp] Using custom label convention: suffixes "
            + string.Join(", ", LabelConvention.Current.Suffixes)
            + (LabelConvention.Current.LabelsDataItemName is { } item
                ? $"; labels data item '{item}'."
                : "; labels data-item rule disabled."));
    }
}

var builder = Host.CreateApplicationBuilder(args);

builder.Logging.AddConsole(options =>
{
    options.LogToStandardErrorThreshold = LogLevel.Trace;
});

builder.Services
    .AddMcpServer()
    .WithStdioServerTransport()
    .WithToolsFromAssembly()
    // Every tool body already returns the uniform {ok,data,error} envelope, but a request that never
    // REACHES the body — most commonly a missing/mistyped required argument, which the SDK's parameter
    // marshaller throws on — would otherwise surface as the SDK's bare "An error occurred invoking
    // '<tool>'." text: no envelope, and the actually-useful message (naming the missing parameter) lost
    // to stderr. This filter wraps the call-tool pipeline to convert such failures into the same
    // envelope shape every other error takes (found by the 2026-07-31 edit-scenario e2e pass).
    .WithRequestFilters(filters => filters.AddCallToolFilter(next => async (context, cancellationToken) =>
    {
        static CallToolResult EnvelopeError(string code, string message) => new()
        {
            IsError = true,
            Content =
            [
                new TextContentBlock
                {
                    Text = JsonSerializer.Serialize(
                        ToolResponse.Failure(
                            code,
                            message,
                            "Check the tool's parameter schema (tools/list) and re-send the call with the "
                            + "required arguments."),
                        McpJsonUtilities.DefaultOptions),
                },
            ],
        };

        try
        {
            var result = await next(context, cancellationToken);

            // Belt-and-braces: if an SDK-internal layer below this filter already swallowed the failure
            // into its generic non-envelope text (isError with a non-JSON first block — no tool in this
            // host ever sets isError itself), rewrap it so the envelope guarantee still holds.
            if (result.IsError == true
                && result.Content.OfType<TextContentBlock>().FirstOrDefault() is { } text
                && !text.Text.TrimStart().StartsWith('{'))
            {
                return EnvelopeError("invalid_request", text.Text);
            }

            return result;
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            return EnvelopeError(ex is ArgumentException ? "invalid_argument" : "internal_error", ex.Message);
        }
    }));

await builder.Build().RunAsync();
