using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
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
    private static readonly Regex ContentSha256Pattern = new(
        "^[0-9a-fA-F]{64}$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    public static bool IsValidContentSha256(string? value) =>
        !string.IsNullOrWhiteSpace(value) && ContentSha256Pattern.IsMatch(value);

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
        sanitized = new JiraTestCaseDocument();
        testCaseBytes = [];
        contentSha256 = string.Empty;
        errors = [];

        // Validate nested structure before projection so malformed arrays never reach serializers.
        if (!TryValidateNestedStructure(source, displayPath, errors))
            return false;

        sanitized = ProjectWhitelist(source);

        var json = SecretRedactor.Redact(JsonSerializer.Serialize(sanitized, CompilationJson.Indented));
        if (!json.EndsWith('\n'))
            json += "\n";
        testCaseBytes = Encoding.UTF8.GetBytes(json);

        var validation = Validator.ValidateBytes(testCaseBytes, displayPath);
        if (!validation.IsValid || validation.Sha256 is null || validation.Document is null)
        {
            errors.AddRange(validation.Errors);
            return false;
        }

        if (!IsValidContentSha256(validation.Sha256))
        {
            errors.Add($"{displayPath}: validator produced invalid contentSha256.");
            return false;
        }

        sanitized = validation.Document;
        contentSha256 = validation.Sha256.ToLowerInvariant();
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

        if (!IsValidContentSha256(expectedContentSha256))
        {
            errors.Add(
                $"{displayPath}: metadata contentSha256 is required and must be a 64-char hex digest " +
                $"(got '{expectedContentSha256 ?? "<null>"}').");
            return false;
        }

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

        if (!string.Equals(validation.Sha256, expectedContentSha256, StringComparison.OrdinalIgnoreCase))
        {
            errors.Add(
                $"{displayPath}: contentSha256 mismatch (metadata '{expectedContentSha256}' vs bytes '{validation.Sha256}').");
            return false;
        }

        document = validation.Document;
        return true;
    }

    public static bool TryValidateNestedStructure(
        JiraTestCaseDocument source,
        string displayPath,
        List<string> errors)
    {
        ArgumentNullException.ThrowIfNull(source);
        ArgumentNullException.ThrowIfNull(errors);
        var before = errors.Count;

        if (source.Background is { } background)
        {
            if (background.Steps is null)
            {
                errors.Add($"{displayPath}: background.steps must not be null.");
            }
            else
            {
                for (var i = 0; i < background.Steps.Count; i++)
                {
                    if (background.Steps[i] is null)
                        errors.Add($"{displayPath}: background.steps[{i}] must not be null.");
                }
            }
        }

        if (source.Scenarios is { } scenarios)
        {
            for (var i = 0; i < scenarios.Count; i++)
            {
                var scenario = scenarios[i];
                if (scenario is null)
                {
                    errors.Add($"{displayPath}: scenarios[{i}] must not be null.");
                    continue;
                }

                if (scenario.Steps is null)
                {
                    errors.Add($"{displayPath}: scenarios[{i}].steps must not be null.");
                    continue;
                }

                for (var j = 0; j < scenario.Steps.Count; j++)
                {
                    if (scenario.Steps[j] is null)
                        errors.Add($"{displayPath}: scenarios[{i}].steps[{j}] must not be null.");
                }

                if (scenario.Tags is { } tags)
                {
                    for (var t = 0; t < tags.Count; t++)
                    {
                        if (tags[t] is null)
                            errors.Add($"{displayPath}: scenarios[{i}].tags[{t}] must not be null.");
                    }
                }
            }
        }

        if (source.Tags is { } rootTags)
        {
            for (var i = 0; i < rootTags.Count; i++)
            {
                if (rootTags[i] is null)
                    errors.Add($"{displayPath}: tags[{i}] must not be null.");
            }
        }

        if (source.TestDataRefs is { } refs)
        {
            for (var i = 0; i < refs.Count; i++)
            {
                if (refs[i] is null)
                    errors.Add($"{displayPath}: testDataRefs[{i}] must not be null.");
            }
        }

        return errors.Count == before;
    }

    private static StructuredBackground? ProjectBackground(StructuredBackground? background)
    {
        if (background is null)
            return null;

        var steps = background.Steps ?? [];
        return new StructuredBackground
        {
            Name = background.Name,
            Steps = steps
                .Where(s => s is not null)
                .Select(s => new StructuredStep
                {
                    Keyword = s.Keyword ?? string.Empty,
                    Text = s.Text ?? string.Empty
                })
                .ToList()
        };
    }

    private static List<StructuredScenario>? ProjectScenarios(List<StructuredScenario>? scenarios)
    {
        if (scenarios is null)
            return null;

        return scenarios
            .Where(s => s is not null)
            .Select(s => new StructuredScenario
            {
                Name = s.Name ?? string.Empty,
                Tags = ProjectTags(s.Tags),
                Steps = (s.Steps ?? [])
                    .Where(st => st is not null)
                    .Select(st => new StructuredStep
                    {
                        Keyword = st.Keyword ?? string.Empty,
                        Text = st.Text ?? string.Empty
                    })
                    .ToList()
            })
            .ToList();
    }

    private static List<string>? ProjectTags(List<string>? tags)
    {
        if (tags is null || tags.Count == 0)
            return null;
        return tags
            .Where(t => t is not null && !string.IsNullOrWhiteSpace(t))
            .Select(t => t.Trim())
            .Distinct(StringComparer.Ordinal)
            .ToList();
    }
}
