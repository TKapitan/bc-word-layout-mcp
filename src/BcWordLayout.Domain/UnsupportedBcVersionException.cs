namespace BcWordLayout.Domain;

/// <summary>
/// Thrown by <see cref="BcFeatures.Require"/> when a call would author a feature the target Business Central
/// version does not have, or the target is unknown (GitHub issue #152). The host reports it as
/// <c>bc_version_unsupported</c>, with a hint built from <see cref="Feature"/> and <see cref="Target"/>.
/// </summary>
public sealed class UnsupportedBcVersionException : Exception
{
    public UnsupportedBcVersionException(BcFeature feature, TargetVersion target)
        : base(BuildMessage(feature, target))
    {
        Feature = feature;
        Target = target;
    }

    /// <summary>The feature the call needs.</summary>
    public BcFeature Feature { get; }

    /// <summary>The target in use, and where it came from.</summary>
    public TargetVersion Target { get; }

    private static string BuildMessage(BcFeature feature, TargetVersion target)
    {
        ArgumentNullException.ThrowIfNull(feature);
        ArgumentNullException.ThrowIfNull(target);

        var needs = $"This call needs {feature.What} ('{feature.Id}'), which Business Central has from "
                    + $"{feature.Since.Describe()}";
        if (feature.Until is { } until)
        {
            needs += $" until {until}";
        }

        return target.Version is null
            ? $"{needs}, and the target BC version is {target.Describe()}, so the tools cannot tell whether "
              + "Business Central can use it. Nothing was changed."
            : $"{needs}; the target is {target.Describe()}. Nothing was changed.";
    }
}
