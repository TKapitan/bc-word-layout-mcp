using System.Globalization;

namespace BcWordLayout.Domain;

/// <summary>
/// A Business Central version as major.minor: <c>29.0</c>, <c>28.3</c> (GitHub issue #152). An AL
/// <c>app.json</c> states four parts (<c>29.0.0.0</c>); build and revision are dropped. Minor matters, because
/// BC behaviour can change within a major version: BC 28.0 and 28.3 accepted an external attached template, and
/// a later 28.x refused it.
/// </summary>
public readonly record struct BcVersion(int Major, int Minor) : IComparable<BcVersion>
{
    /// <summary>
    /// Reads <c>29</c>, <c>29.1</c> or a four-part <c>29.0.0.0</c> as major.minor. False for anything else:
    /// blank, a sign or stray whitespace inside, an empty or non-numeric part, more than four parts, or a
    /// major version below 1.
    /// </summary>
    public static bool TryParse(string? text, out BcVersion version)
    {
        version = default;
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        var parts = text.Trim().Split('.');
        if (parts.Length > 4)
        {
            return false;
        }

        var numbers = new int[parts.Length];
        for (var i = 0; i < parts.Length; i++)
        {
            if (!int.TryParse(parts[i], NumberStyles.None, CultureInfo.InvariantCulture, out numbers[i]))
            {
                return false;
            }
        }

        if (numbers[0] < 1)
        {
            return false;
        }

        version = new BcVersion(numbers[0], parts.Length > 1 ? numbers[1] : 0);
        return true;
    }

    /// <summary>
    /// The release wave Microsoft shipped this major version in, e.g. <c>2026 release wave 2</c> for 29. Null
    /// before 14 (2019 release wave 1), where the naming does not follow the twice-yearly pattern.
    /// </summary>
    public string? ReleaseWave =>
        Major < 14 ? null : $"{2019 + ((Major - 14) / 2)} release wave {(Major % 2 == 0 ? 1 : 2)}";

    /// <summary><c>BC 29.0 (2026 release wave 2)</c>: the version with its release wave, for messages.</summary>
    public string Describe() => ReleaseWave is { } wave ? $"BC {this} ({wave})" : $"BC {this}";

    public int CompareTo(BcVersion other) =>
        Major != other.Major ? Major.CompareTo(other.Major) : Minor.CompareTo(other.Minor);

    public override string ToString() => $"{Major}.{Minor}";

    public static bool operator <(BcVersion left, BcVersion right) => left.CompareTo(right) < 0;

    public static bool operator >(BcVersion left, BcVersion right) => left.CompareTo(right) > 0;

    public static bool operator <=(BcVersion left, BcVersion right) => left.CompareTo(right) <= 0;

    public static bool operator >=(BcVersion left, BcVersion right) => left.CompareTo(right) >= 0;
}
