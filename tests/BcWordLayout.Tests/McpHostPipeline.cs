using System.ComponentModel;
using System.IO.Pipelines;
using System.Text.Json;
using BcWordLayout.McpHost;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace BcWordLayout.Tests;

/// <summary>
/// The REAL server pipeline, in process: a generic host configured exactly like <c>Program.cs</c>
/// (<see cref="ToolCallBoundary.WithBcWordLayoutTools"/> — the tools plus the call-tool filter), talking to an
/// SDK <see cref="McpClient"/> over a pair of in-memory pipes instead of stdio. Every call therefore goes
/// through the SDK's JSON-RPC framing, tool lookup and argument binding, and the filter, which the direct
/// static-method calls in <see cref="McpHostToolTests"/> skip entirely. No process is spawned and nothing is
/// built on the fly; the only extra registration is <see cref="SwallowedFailureTestTool"/>, a test-only tool
/// that stands in for an SDK layer swallowing a binding failure (see its own remarks).
/// </summary>
public sealed class McpHostPipeline : IAsyncLifetime
{
    private readonly Pipe clientToServer = new();
    private readonly Pipe serverToClient = new();
    private IHost? host;
    private McpClient? client;

    public McpClient Client => client ?? throw new InvalidOperationException("The pipeline has not started.");

    public async Task InitializeAsync()
    {
        var builder = Host.CreateEmptyApplicationBuilder(new HostApplicationBuilderSettings());
        builder.Services
            .AddMcpServer()
            .WithStreamServerTransport(clientToServer.Reader.AsStream(), serverToClient.Writer.AsStream())
            .WithBcWordLayoutTools()
            .WithTools([typeof(SwallowedFailureTestTool)]);
        host = builder.Build();
        await host.StartAsync();

        client = await McpClient.CreateAsync(
            new StreamClientTransport(clientToServer.Writer.AsStream(), serverToClient.Reader.AsStream()));
    }

    public async Task DisposeAsync()
    {
        if (client is not null)
        {
            await client.DisposeAsync();
        }

        await clientToServer.Writer.CompleteAsync();
        if (host is not null)
        {
            await host.StopAsync();
            host.Dispose();
        }
    }

    /// <summary>
    /// Calls <paramref name="tool"/> with the given raw JSON object as its <c>arguments</c> — raw JSON rather
    /// than typed values, because the point is to send exactly what a misbehaving agent sends (a quoted
    /// number, an explicit null, a misspelled key).
    /// </summary>
    public async Task<PipelineCallResult> CallAsync(string tool, string argumentsJson)
    {
        var arguments = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(argumentsJson)
            ?? throw new ArgumentException("argumentsJson must be a JSON object.", nameof(argumentsJson));
        var result = await Client.CallToolAsync(new CallToolRequestParams { Name = tool, Arguments = arguments });
        var text = result.Content.OfType<TextContentBlock>().FirstOrDefault()?.Text;
        Assert.True(text is not null, $"{tool}: the result carried no text block at all.");
        using var envelope = JsonDocument.Parse(text!);
        return new PipelineCallResult(result.IsError, envelope.RootElement.Clone(), text!);
    }

    /// <summary>JSON-encodes a path (or any string) for splicing into a raw arguments object.</summary>
    public static string Json(string value) => JsonSerializer.Serialize(value);
}

/// <summary>One call's outcome: the MCP-level <c>isError</c> flag plus the parsed envelope JSON.</summary>
public sealed record PipelineCallResult(bool? IsError, JsonElement Envelope, string RawText)
{
    public bool Ok => Envelope.GetProperty("ok").GetBoolean();

    public string? Code => Envelope.TryGetProperty("error", out var e) && e.ValueKind == JsonValueKind.Object
        ? e.GetProperty("code").GetString()
        : null;

    public string Message => Envelope.GetProperty("error").GetProperty("message").GetString() ?? string.Empty;

    public string Hint => Envelope.GetProperty("error").GetProperty("hint").GetString() ?? string.Empty;

    public JsonElement Data => Envelope.GetProperty("data");

    /// <summary>
    /// Asserts a failure envelope with <paramref name="code"/>, a message mentioning every
    /// <paramref name="messageMentions"/> fragment, and a non-empty hint.
    /// </summary>
    public void AssertFailure(string code, params string[] messageMentions)
    {
        Assert.False(Ok, $"Expected a '{code}' failure but the call succeeded: {RawText}");
        Assert.Equal(code, Code);
        foreach (var fragment in messageMentions)
        {
            Assert.Contains(fragment, Message, StringComparison.Ordinal);
        }

        Assert.False(string.IsNullOrWhiteSpace(Hint));
    }
}

/// <summary>
/// Test-only tool registered on <see cref="McpHostPipeline"/>'s server only (never on the shipped one): it
/// returns exactly what an SDK layer below the call-tool filter would produce if it swallowed a binding
/// failure instead of throwing it — <c>isError:true</c> with the SDK's generic, non-JSON text. No shipped
/// tool can produce that shape (every one returns an envelope), so this is the only way to drive the
/// filter's belt-and-braces branch through the real pipeline.
/// </summary>
[McpServerToolType]
public static class SwallowedFailureTestTool
{
    public const string Name = "test_swallowed_binding_failure";

    public const string SdkText = "An error occurred invoking 'test_swallowed_binding_failure'.";

    [McpServerTool(Name = Name)]
    [Description("Test-only: returns the SDK's generic non-envelope error text.")]
    public static CallToolResult Invoke() => new() { IsError = true, Content = [new TextContentBlock { Text = SdkText }] };
}
