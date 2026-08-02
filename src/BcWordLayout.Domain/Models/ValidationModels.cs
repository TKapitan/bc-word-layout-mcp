namespace BcWordLayout.Domain.Models;

public enum FindingSeverity
{
    Warning,
    Error,
}

/// <summary>A single validation finding with a stable check id, severity and human-actionable message.</summary>
public sealed class ValidationFinding
{
    /// <summary>Stable identifier of the check that produced this finding (e.g. <c>openxml-structure</c>).</summary>
    public required string Check { get; init; }

    public required FindingSeverity Severity { get; init; }

    public required string Message { get; init; }

    /// <summary>Optional location context (part name, xpath, control alias) to help the caller act.</summary>
    public string? Location { get; init; }
}

/// <summary>The outcome of a validation run: a flat list of findings and an overall pass/fail.</summary>
public sealed class ValidationResult
{
    public required string Level { get; init; }

    public required IReadOnlyList<ValidationFinding> Findings { get; init; }

    /// <summary>True when there are no error-severity findings (warnings do not fail validation).</summary>
    public bool Passed => !Findings.Any(f => f.Severity == FindingSeverity.Error);

    public int ErrorCount => Findings.Count(f => f.Severity == FindingSeverity.Error);

    public int WarningCount => Findings.Count(f => f.Severity == FindingSeverity.Warning);
}
