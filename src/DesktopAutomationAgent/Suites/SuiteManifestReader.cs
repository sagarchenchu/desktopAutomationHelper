using System.Text.Json;
using System.Text.RegularExpressions;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Workspace;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Suites;

public sealed class SuiteManifestReader : ISuiteManifestReader
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true
    };

    private readonly AgentOptions _options;
    private readonly IWorkspaceManager _workspace;
    private readonly Regex _jiraKeyRegex;

    public SuiteManifestReader(IOptions<AgentOptions> options, IWorkspaceManager workspace)
    {
        _options = options.Value;
        _workspace = workspace;
        AgentOptionsValidator.Validate(_options, OptionsValidationScope.Suites);
        _jiraKeyRegex = new Regex(
            _options.Suites.JiraKeyPattern,
            RegexOptions.CultureInvariant | RegexOptions.Compiled);
    }

    public SuiteValidationResult ValidateFile(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return new SuiteValidationResult
            {
                FilePath = path ?? string.Empty,
                SuiteName = string.Empty,
                Errors = ["Suite file path is required."]
            };
        }

        string fullPath;
        try
        {
            fullPath = _workspace.ResolveSafePath(path);
        }
        catch (WorkspaceException ex)
        {
            return new SuiteValidationResult
            {
                FilePath = path,
                SuiteName = string.Empty,
                Errors = [ex.Message]
            };
        }

        if (!File.Exists(fullPath))
        {
            return new SuiteValidationResult
            {
                FilePath = fullPath,
                SuiteName = string.Empty,
                Errors = [$"Suite file not found: '{fullPath}'."]
            };
        }

        SuiteManifest? manifest;
        try
        {
            var json = File.ReadAllText(fullPath);
            manifest = JsonSerializer.Deserialize<SuiteManifest>(json, JsonOptions);
        }
        catch (JsonException ex)
        {
            return new SuiteValidationResult
            {
                FilePath = fullPath,
                SuiteName = string.Empty,
                Errors = [$"{fullPath}: invalid JSON ({ex.Message})."]
            };
        }

        if (manifest is null)
        {
            return new SuiteValidationResult
            {
                FilePath = fullPath,
                SuiteName = string.Empty,
                Errors = [$"{fullPath}: suite manifest was empty."]
            };
        }

        var errors = new List<string>();
        if (manifest.SchemaVersion != 1)
        {
            errors.Add($"{fullPath}: unsupported schemaVersion {manifest.SchemaVersion}; expected 1.");
        }

        if (string.IsNullOrWhiteSpace(manifest.Name))
            errors.Add($"{fullPath}: 'name' is required.");

        if (manifest.TestCases is null)
        {
            errors.Add($"{fullPath}: 'testCases' is required.");
            manifest.TestCases = [];
        }

        var seen = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var enabledKeys = new List<string>();
        var disabledCount = 0;
        var duplicateCount = 0;

        for (var i = 0; i < manifest.TestCases.Count; i++)
        {
            var entry = manifest.TestCases[i];
            var location = $"{fullPath}: testCases[{i}]";
            var entryValid = true;

            if (string.IsNullOrWhiteSpace(entry.JiraKey))
            {
                errors.Add($"{location}: jiraKey is required.");
                continue;
            }

            var key = entry.JiraKey.Trim();
            if (!_jiraKeyRegex.IsMatch(key))
            {
                errors.Add($"{location}: invalid jiraKey '{key}' (pattern: {_options.Suites.JiraKeyPattern}).");
                entryValid = false;
            }

            if (seen.TryGetValue(key, out var previousIndex))
            {
                duplicateCount++;
                errors.Add($"{location}: duplicate jiraKey '{key}' (first seen at testCases[{previousIndex}]).");
                entryValid = false;
            }
            else
            {
                seen[key] = i;
            }

            if (!entry.Enabled || !manifest.Enabled)
            {
                disabledCount++;
            }
            else if (entryValid)
            {
                enabledKeys.Add(key);
            }
        }

        return new SuiteValidationResult
        {
            FilePath = fullPath,
            SuiteName = manifest.Name ?? string.Empty,
            SuiteEnabled = manifest.Enabled,
            TotalCount = manifest.TestCases.Count,
            EnabledCount = enabledKeys.Count,
            DisabledCount = disabledCount,
            DuplicateCount = duplicateCount,
            EnabledJiraKeys = enabledKeys,
            Errors = errors
        };
    }

    public KeyValidationResult ValidateKeys(IEnumerable<string> keys)
    {
        ArgumentNullException.ThrowIfNull(keys);

        var errors = new List<string>();
        var valid = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var index = 0;

        foreach (var raw in keys)
        {
            var location = $"keys[{index}]";
            index++;

            if (string.IsNullOrWhiteSpace(raw))
            {
                errors.Add($"{location}: jiraKey is required.");
                continue;
            }

            var key = raw.Trim();
            if (!_jiraKeyRegex.IsMatch(key))
            {
                errors.Add($"{location}: invalid jiraKey '{key}' (pattern: {_options.Suites.JiraKeyPattern}).");
                continue;
            }

            if (!seen.Add(key))
            {
                errors.Add($"{location}: duplicate jiraKey '{key}'.");
                continue;
            }

            valid.Add(key);
        }

        return new KeyValidationResult
        {
            ValidKeys = valid,
            Errors = errors
        };
    }
}
