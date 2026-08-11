using System.Text.Json;
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

    public SuiteManifestReader(IOptions<AgentOptions> options, IWorkspaceManager workspace)
    {
        _options = options.Value;
        _workspace = workspace;
        AgentOptionsValidator.Validate(_options, OptionsValidationScope.Suites);
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
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return new SuiteValidationResult
            {
                FilePath = fullPath,
                SuiteName = string.Empty,
                Errors = [$"{fullPath}: failed to read suite file ({ex.Message})."]
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
            return new SuiteValidationResult
            {
                FilePath = fullPath,
                SuiteName = manifest.Name ?? string.Empty,
                SuiteEnabled = manifest.Enabled,
                TotalCount = 0,
                EnabledCount = 0,
                DisabledCount = 0,
                DuplicateCount = 0,
                EnabledJiraKeys = [],
                Errors = errors
            };
        }

        var testCases = manifest.TestCases;
        var seen = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var enabledKeys = new List<string>();
        var disabledCount = 0;
        var duplicateCount = 0;
        var projectPattern = _options.Suites.JiraKeyPattern;

        for (var i = 0; i < testCases.Count; i++)
        {
            var entry = testCases[i];
            var location = $"{fullPath}: testCases[{i}]";

            if (string.IsNullOrWhiteSpace(entry.JiraKey))
            {
                errors.Add($"{location}: jiraKey is required.");
                continue;
            }

            var entryValid = JiraKeyContract.TryValidate(
                entry.JiraKey, projectPattern, out var key, out var keyError);
            if (!entryValid)
            {
                errors.Add($"{location}: {keyError}");
                key = entry.JiraKey.Trim();
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
            TotalCount = testCases.Count,
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

        var materialised = keys.ToArray();
        if (materialised.Length == 0)
        {
            return new KeyValidationResult
            {
                ValidKeys = [],
                Errors = ["At least one Jira key is required."]
            };
        }

        var errors = new List<string>();
        var valid = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var index = 0;
        var projectPattern = _options.Suites.JiraKeyPattern;

        foreach (var raw in materialised)
        {
            var location = $"keys[{index}]";
            index++;

            if (!JiraKeyContract.TryValidate(raw, projectPattern, out var key, out var keyError))
            {
                errors.Add($"{location}: {keyError}");
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
