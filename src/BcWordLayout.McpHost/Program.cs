using BcWordLayout.Domain;
using BcWordLayout.McpHost;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

// BC Word Layout MCP server: tools that inspect, validate, preview, create, edit and refresh BC Word layouts.
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

// Optional target Business Central version for layouts outside an AL project (GitHub issue #152): BCWL_BC_VERSION,
// e.g. "29". Every call resolves its own target - the nearest app.json above the layout wins over this value - so
// it is only stored here, never parsed into server state. An invalid value is kept as-is: it makes the target
// unknown, and the responses that report the target name the reason.
var rawBcVersion = Environment.GetEnvironmentVariable(EnvironmentConfig.BcVersionVariable);
BcWordLayout.McpHost.Tools.ToolGuards.BcVersionEnvironmentValue = rawBcVersion;
if (!string.IsNullOrWhiteSpace(rawBcVersion) && !BcVersion.TryParse(rawBcVersion, out _))
{
    Console.Error.WriteLine(
        $"[bc-word-layout-mcp] {EnvironmentConfig.BcVersionVariable}='{rawBcVersion}' is not a Business Central "
        + "version (expected e.g. 29 or 29.1); layouts outside an AL project get an unknown target.");
}

var builder = Host.CreateApplicationBuilder(args);

builder.Logging.AddConsole(options =>
{
    options.LogToStandardErrorThreshold = LogLevel.Trace;
});

builder.Services
    .AddMcpServer()
    .WithStdioServerTransport()
    // serverInfo: the server name plus the PRODUCT version (1.1.0, 1.2.0-rc.1) rather than the SDK's fallback, the
    // four-part assembly version an RC shares with its release - see ServerIdentity. Shared with the host-level tests.
    .WithBcWordLayoutServerInfo()
    // The tools plus the call-tool filter that keeps the {ok,data,error} envelope guarantee for calls that
    // never reach a tool body (unknown tool, missing or wrong-typed argument) — see ToolCallBoundary. Shared
    // with the host-level tests, which run the same pipeline over an in-memory stream transport.
    .WithBcWordLayoutTools();

await builder.Build().RunAsync();
