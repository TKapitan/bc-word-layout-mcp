using System.Text.Json;

namespace BcWordLayout.Domain;

/// <summary>
/// Resolves the <see cref="TargetVersion"/> of a layout for one call (GitHub issue #152, ADR-0007). The first
/// match wins:
/// <list type="number">
/// <item>the nearest AL <c>app.json</c> above the layout: its <c>application</c> value, e.g. <c>29.0.0.0</c>. The
/// walk stops at the first <c>app.json</c>, at a folder containing <c>.git</c>, or at the drive root. An
/// <c>app.json</c> that cannot be read, or names something that is not a version, makes the target unknown;
/// one with no <c>application</c> falls through to the environment variable;</item>
/// <item><see cref="EnvironmentVariable"/> (e.g. <c>29</c> or <c>29.1</c>), for layouts outside an AL project;</item>
/// <item>unknown, with the reason.</item>
/// </list>
/// There is no per-call parameter and no stored state: every call that needs the target resolves it again.
/// </summary>
public static class TargetVersionResolver
{
    /// <summary>The environment variable naming the target for layouts outside an AL project.</summary>
    public const string EnvironmentVariable = "BCWL_BC_VERSION";

    private const string AppJsonName = "app.json";

    private static readonly JsonDocumentOptions AppJsonOptions = new()
    {
        // AL projects keep comments and trailing commas in app.json; the AL tooling accepts both.
        CommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true,
        MaxDepth = 64,
    };

    /// <summary>
    /// The target for the layout at <paramref name="layoutPath"/> (for <c>create_layout</c>, the output path,
    /// which need not exist yet), given the raw value of <see cref="EnvironmentVariable"/>. Never throws for a
    /// file-system or JSON problem: that makes the target unknown, with the reason in
    /// <see cref="TargetVersion.Detail"/>.
    /// </summary>
    public static TargetVersion Resolve(string layoutPath, string? environmentValue)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(layoutPath);

        var appJson = FindAppJson(layoutPath, out var searchedTo);
        string? appJsonWithoutApplication = null;
        if (appJson is not null)
        {
            if (ReadAppJson(appJson) is { } fromAppJson)
            {
                return fromAppJson;
            }

            appJsonWithoutApplication = appJson;
        }

        if (string.IsNullOrWhiteSpace(environmentValue))
        {
            return TargetVersion.Unknown(
                (appJsonWithoutApplication is not null
                    ? $"{appJsonWithoutApplication} has no application property"
                    : $"no app.json above the layout, searched up to {searchedTo}")
                + $", and {EnvironmentVariable} is not set");
        }

        if (!BcVersion.TryParse(environmentValue, out var version))
        {
            return TargetVersion.Unknown(
                $"{EnvironmentVariable}='{environmentValue.Trim()}' is not a Business Central version; give a major "
                + "or major.minor version such as 29 or 29.1");
        }

        return TargetVersion.FromEnvironment(
            version,
            environmentValue,
            appJsonWithoutApplication is null ? null : $"{appJsonWithoutApplication} has no application property");
    }

    /// <summary>
    /// The nearest <c>app.json</c> in <paramref name="layoutPath"/>'s folder or above it, or null. The walk
    /// stops at a folder containing <c>.git</c> (a folder, or a file in a worktree): an AL project inside a
    /// repository never takes its target from an <c>app.json</c> outside that repository.
    /// <paramref name="searchedTo"/> names the folder the walk stopped at.
    /// </summary>
    private static string? FindAppJson(string layoutPath, out string searchedTo)
    {
        var folder = Path.GetDirectoryName(Path.GetFullPath(layoutPath));
        searchedTo = folder ?? layoutPath;
        while (folder is not null)
        {
            searchedTo = folder;
            var candidate = Path.Combine(folder, AppJsonName);
            if (File.Exists(candidate))
            {
                return candidate;
            }

            var git = Path.Combine(folder, ".git");
            if (Directory.Exists(git) || File.Exists(git))
            {
                searchedTo = $"{folder} (the repository root)";
                return null;
            }

            folder = Path.GetDirectoryName(folder);
        }

        return null;
    }

    /// <summary>
    /// The target <paramref name="appJsonPath"/> states: a known one, an unknown one when the file cannot be read
    /// or names something that is not a version, or null when it has no <c>application</c> property at all.
    /// </summary>
    private static TargetVersion? ReadAppJson(string appJsonPath)
    {
        try
        {
            var length = new FileInfo(appJsonPath).Length;
            if (length > ResourceLimits.MaxAppJsonBytes)
            {
                return TargetVersion.Unknown(
                    $"{appJsonPath} could not be read: it is {length} bytes, more than the "
                    + $"{ResourceLimits.MaxAppJsonBytes} an app.json may have");
            }

            using var stream = File.OpenRead(appJsonPath);
            using var json = JsonDocument.Parse(stream, AppJsonOptions);
            if (json.RootElement.ValueKind != JsonValueKind.Object)
            {
                return TargetVersion.Unknown($"{appJsonPath} could not be read: it is not a JSON object");
            }

            JsonElement? application = null;
            foreach (var property in json.RootElement.EnumerateObject())
            {
                if (string.Equals(property.Name, "application", StringComparison.OrdinalIgnoreCase))
                {
                    application = property.Value;
                    break;
                }
            }

            if (application is not { } value)
            {
                return null;
            }

            var text = value.ValueKind == JsonValueKind.String ? value.GetString() : value.GetRawText();
            return BcVersion.TryParse(text, out var version)
                ? TargetVersion.FromAppJson(version, appJsonPath, text!.Trim())
                : TargetVersion.Unknown($"{appJsonPath}: application '{text}' is not a Business Central version");
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException)
        {
            return TargetVersion.Unknown($"{appJsonPath} could not be read: {ex.Message}");
        }
    }
}
