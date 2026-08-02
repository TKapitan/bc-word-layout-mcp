using BcWordLayout.Domain;
using BcWordLayout.Domain.Models;

namespace BcWordLayout.Merge;

/// <summary>
/// Full-level layout validation: combines <see cref="LayoutValidator.Quick(string)"/>'s structural/binding
/// findings with a real dry-run merge (<see cref="MergeEngine.Merge(string, string, MergeOptions?)"/>)
/// against a throwaway temp output, surfacing every <see cref="MergeWarning"/> as a <see cref="ValidationFinding"/>.
/// This lives in the Merge project (not Domain) because Domain must not depend on Merge (Merge depends on
/// Domain) — <c>BcWordLayout.McpHost.Tools.ReadTools</c> is the only caller, for <c>validate_layout</c>'s
/// <c>level == "full"</c>.
/// </summary>
public static class FullValidator
{
    /// <summary>
    /// Runs quick validation, then a dry-run merge with default <see cref="MergeOptions"/>, against
    /// <paramref name="layoutPath"/>. The dry-run's merged .docx is written to a temp file that is always
    /// deleted afterward (try/finally), whether or not the merge succeeds. Only the dry-run MERGE step is
    /// wrapped in a try/catch: a merge-time exception is captured as a single <c>dry-run-merge</c> error
    /// finding rather than propagating. <see cref="LayoutValidator.Quick(string)"/> runs outside that
    /// try/catch, on purpose — a missing file or an unopenable/malformed package still throws straight out
    /// of this method, exactly as it does at quick level, so the caller (<c>ToolGuards</c>'s <c>Guard</c>)
    /// turns it into the same structured <c>file_not_found</c>/<c>invalid_layout</c> error envelope that
    /// <c>validate_layout level=quick</c> would produce for that same broken file, instead of it being
    /// buried as an ordinary validation finding.
    /// <para>
    /// SECURITY: the dry-run merge below deliberately leaves
    /// <see cref="MergeOptions.StripExternalRelationships"/> at its default <c>false</c>. Unlike
    /// <c>preview_layout</c>, this merge's output (<c>tempOutput</c>) is never handed to a PDF converter and
    /// is always deleted in the <c>finally</c> block below regardless of outcome - so no renderer could ever
    /// dereference a poisoned layout's external relationship here, and stripping them would only add cost
    /// with no security benefit. <see cref="LayoutValidator.Quick(string)"/> (invoked above, and included in
    /// this method's own findings either way) already surfaces a real <c>attachedTemplate</c> relationship
    /// as its own <c>attached-template</c> warning finding.
    /// </para>
    /// </summary>
    public static ValidationResult Full(string layoutPath)
    {
        var findings = new List<ValidationFinding>();

        var quick = LayoutValidator.Quick(layoutPath);
        findings.AddRange(quick.Findings);

        var tempOutput = Path.Combine(Path.GetTempPath(), $"bcwl-full-validate-{Guid.NewGuid():N}.docx");
        try
        {
            var mergeResult = MergeEngine.Merge(layoutPath, tempOutput, new MergeOptions());
            foreach (var warning in mergeResult.Warnings)
            {
                findings.Add(new ValidationFinding
                {
                    Check = warning.Kind,
                    Severity = SeverityFor(warning.Kind),
                    Message = warning.Message,
                    Location = warning.Location,
                });
            }
        }
        catch (Exception ex)
        {
            findings.Add(new ValidationFinding
            {
                Check = "dry-run-merge",
                Severity = FindingSeverity.Error,
                Message = $"Dry-run merge failed: {ex.Message}",
            });
        }
        finally
        {
            if (File.Exists(tempOutput))
            {
                File.Delete(tempOutput);
            }
        }

        return new ValidationResult { Level = "full", Findings = findings };
    }

    /// <summary>
    /// Maps a <see cref="MergeWarning.Kind"/> to a validation severity. <c>unresolved-binding</c> (a
    /// binding that did not resolve against sample data) and <c>xpath-error</c> (a malformed XPath) are
    /// Errors — both indicate the layout will render visibly wrong in BC. <c>xpath-fallback</c> (re-anchor
    /// fell back to document-root evaluation), <c>content-write-failed</c> (resolved a value but had
    /// nowhere to write it), <c>picture-no-blip</c> (picture control with no repointable image), and
    /// <c>row-cap</c> (a repeating section exceeded <see cref="MergeOptions.MaxRowsPerRepeater"/> and was
    /// capped) are Warnings — degraded but not necessarily broken; a capped preview in particular is a
    /// robustness measure on the MOCK render, not evidence the real layout is wrong.
    /// </summary>
    private static FindingSeverity SeverityFor(string kind) => kind switch
    {
        "unresolved-binding" => FindingSeverity.Error,
        "xpath-error" => FindingSeverity.Error,
        _ => FindingSeverity.Warning,
    };
}
