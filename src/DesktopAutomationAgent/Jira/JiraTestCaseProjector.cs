using System.Text;
using System.Text.Json;
using DesktopAutomationAgent.Compilation;
using DesktopAutomationAgent.Driver;

namespace DesktopAutomationAgent.Jira;

/// <summary>
/// Whitelist-projects mapped documents onto the Phase 4D schema surface only
/// (no extension data / unrestricted vendor fields), then validates + hashes
/// the sanitized bytes used for jira-cache.
/// </summary>
internal static class JiraTestCaseProjector
{
    private static readonly JiraTestCaseValidator Validator = new();

    public static JiraTestCaseDocument ProjectWhitelist(JiraTestCaseDocument source)
    {
        ArgumentNullException.ThrowIfNull(source);

        return new JiraTestCaseDocument
        {
            Schema = null,
            SchemaVersion = 1,
            JiraKey = source.JiraKey?.Trim() ?? string.Empty,
            Revision = string.IsNullOrWhiteSpace(source.Revision) ? null : source.Revision.Trim(),
            UpdatedAt = source.UpdatedAt?.ToUniversalTime(),
            Summary = source.Summary?.Trim() ?? string.Empty,
            Description = null, // never persist unrestricted description dumps
            Feature = source.Feature,
            Background = ProjectBackground(source.Background),
            Scenarios = ProjectScenarios(source.Scenarios),
            Tags = ProjectTags(source.Tags),
            TestDataRefs = ProjectTags(source.TestDataRefs),
            ExtensionData = null
        };
    }

    public static bool TrySanitizeAndHash(
        JiraTestCaseDocument source,
        string displayPath,
        out JiraTestCaseDocument sanitized,
        out byte[] testCaseBytes,
        out string contentSha256,
        out List<string> errors)
    {
        sanitized = ProjectWhitelist(source);
        errors = [];

        var json = SecretRedactor.Redact(JsonSerializer.Serialize(sanitized, CompilationJson.Indented));
        if (!json.EndsWith('\n'))
            json += "\n";
        testCaseBytes = Encoding.UTF8.GetBytes(json);

        var validation = Validator.ValidateBytes(testCaseBytes, displayPath);
        if (!validation.IsValid || validation.Sha256 is null || validation.Document is null)
        {
            errors.AddRange(validation.Errors);
            contentSha256 = string.Empty;
            return false;
        }

        // Prefer the schema-validated document (already whitelist-shaped).
        sanitized = validation.Document;
        contentSha256 = validation.Sha256;
        return true;
    }

    public static bool TryValidateCachedBytes(
        byte[] testCaseBytes,
        string expectedKey,
        string? expectedContentSha256,
        string displayPath,
        out JiraTestCaseDocument? document,
        out List<string> errors)
    {
        document = null;
        errors = [];
        var validation = Validator.ValidateBytes(testCaseBytes, displayPath);
        if (!validation.IsValid || validation.Document is null || validation.Sha256 is null)
        {
            errors.AddRange(validation.Errors);
            return false;
        }

        if (!string.Equals(validation.Document.JiraKey, expectedKey, StringComparison.Ordinal))
        {
            errors.Add($"{displayPath}: jiraKey '{validation.Document.JiraKey}' does not match '{expectedKey}'.");
            return false;
        }

        if (!string.IsNullOrWhiteSpace(expectedContentSha256)
            && !string.Equals(validation.Sha256, expectedContentSha256, StringComparison.OrdinalIgnoreCase))
        {
            errors.Add(
                $"{displayPath}: contentSha256 mismatch (metadata '{expectedContentSha256}' vs bytes '{validation.Sha256}').");
            return false;
        }

        document = validation.Document;
        return true;
    }

    private static StructuredBackground? ProjectBackground(StructuredBackground? background)
    {
        if (background is null)
            return null;
        return new StructuredBackground
        {
            Name = background.Name,
            Steps = background.Steps.Select(s => new StructuredStep
            {
                Keyword = s.Keyword,
                Text = s.Text
            }).ToList()
        };
    }

    private static List<StructuredScenario>? ProjectScenarios(List<StructuredScenario>? scenarios)
    {
        if (scenarios is null)
            return null;
        return scenarios.Select(s => new StructuredScenario
        {
            Name = s.Name,
            Tags = ProjectTags(s.Tags),
            Steps = s.Steps.Select(st => new StructuredStep
            {
                Keyword = st.Keyword,
                Text = st.Text
            }).ToList()
        }).ToList();
    }

    private static List<string>? ProjectTags(List<string>? tags)
    {
        if (tags is null || tags.Count == 0)
            return null;
        return tags
            .Where(t => !string.IsNullOrWhiteSpace(t))
            .Select(t => t.Trim())
            .Distinct(StringComparer.Ordinal)
            .ToList();
    }
}
