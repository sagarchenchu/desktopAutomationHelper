using DesktopAutomationAgent.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Workspace;

public sealed class WorkspaceManager : IWorkspaceManager
{
    private readonly WorkspaceOptions _options;
    private readonly ILogger<WorkspaceManager> _logger;
    private readonly string _rootPath;

    public WorkspaceManager(IOptions<AgentOptions> options, ILogger<WorkspaceManager> logger)
    {
        _options = options.Value.Workspace;
        _logger = logger;
        _rootPath = Path.GetFullPath(_options.Root);
    }

    public string RootPath => _rootPath;

    public WorkspaceInitResult Initialize()
    {
        AgentOptionsValidator.Validate(new AgentOptions { Workspace = _options }, OptionsValidationScope.Workspace);

        try
        {
            var created = new List<string>();
            var skipped = new List<string>();

            Directory.CreateDirectory(_rootPath);

            foreach (var relativeDir in RelativeDirectories)
            {
                var full = Path.Combine(_rootPath, relativeDir);
                EnsureInsideRoot(full);
                if (Directory.Exists(full))
                {
                    skipped.Add(relativeDir + Path.DirectorySeparatorChar);
                    continue;
                }

                Directory.CreateDirectory(full);
                created.Add(relativeDir + Path.DirectorySeparatorChar);
            }

            foreach (var (relativePath, contents) in TemplateFiles)
            {
                var full = Path.Combine(_rootPath, relativePath);
                EnsureInsideRoot(full);
                Directory.CreateDirectory(Path.GetDirectoryName(full)!);

                if (File.Exists(full))
                {
                    skipped.Add(relativePath);
                    continue;
                }

                File.WriteAllText(full, contents);
                created.Add(relativePath);
            }

            _logger.LogInformation(
                "Workspace initialized at {Root}. created={CreatedCount} skippedExisting={SkippedCount}",
                _rootPath,
                created.Count,
                skipped.Count);

            return new WorkspaceInitResult
            {
                RootPath = _rootPath,
                CreatedPaths = created,
                SkippedExistingPaths = skipped
            };
        }
        catch (WorkspaceException)
        {
            throw;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Security.SecurityException)
        {
            throw new WorkspaceException(
                $"Failed to initialize workspace at '{_rootPath}': {ex.Message}",
                ex);
        }
    }

    public void EnsureInitialized()
    {
        try
        {
            if (!Directory.Exists(_rootPath))
            {
                throw new WorkspaceException(
                    $"Workspace root '{_rootPath}' does not exist. Run 'init' first.");
            }

            foreach (var relativeDir in RelativeDirectories)
            {
                var full = Path.Combine(_rootPath, relativeDir);
                if (!Directory.Exists(full))
                {
                    throw new WorkspaceException(
                        $"Workspace directory '{relativeDir}' is missing under '{_rootPath}'. Run 'init' first.");
                }
            }
        }
        catch (WorkspaceException)
        {
            throw;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Security.SecurityException)
        {
            throw new WorkspaceException(
                $"Failed to inspect workspace at '{_rootPath}': {ex.Message}",
                ex);
        }
    }

    public string ResolveSafePath(string relativeOrAbsolutePath)
    {
        if (string.IsNullOrWhiteSpace(relativeOrAbsolutePath))
            throw new WorkspaceException("Path is required.");

        if (relativeOrAbsolutePath.Contains("..", StringComparison.Ordinal))
            throw new WorkspaceException($"Path traversal is not allowed: '{relativeOrAbsolutePath}'.");

        string candidate;
        if (Path.IsPathRooted(relativeOrAbsolutePath))
        {
            candidate = Path.GetFullPath(relativeOrAbsolutePath);
        }
        else
        {
            // Prefer a path that already resolves inside the workspace from the
            // current directory (e.g. automation/suites/smoke.json), otherwise
            // treat the argument as workspace-relative (suites/smoke.json).
            var fromCwd = Path.GetFullPath(relativeOrAbsolutePath);
            candidate = IsInsideRoot(fromCwd)
                ? fromCwd
                : Path.GetFullPath(Path.Combine(_rootPath, relativeOrAbsolutePath));
        }

        EnsureInsideRoot(candidate);
        return candidate;
    }

    private void EnsureInsideRoot(string fullPath)
    {
        if (!IsInsideRoot(fullPath))
        {
            throw new WorkspaceException(
                $"Resolved path '{Path.GetFullPath(fullPath)}' is outside the workspace root '{_rootPath}'.");
        }
    }

    private bool IsInsideRoot(string fullPath)
    {
        var root = Path.GetFullPath(_rootPath)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var normalized = Path.GetFullPath(fullPath)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

        if (PathsEqual(normalized, root))
            return true;

        var relative = Path.GetRelativePath(root, normalized);
        if (string.IsNullOrEmpty(relative) || relative == ".")
            return true;

        // Path.GetRelativePath uses OS path rules. Reject anything that escapes
        // the root via ".." or an absolute/rooted relative result.
        return !Path.IsPathRooted(relative)
               && relative != ".."
               && !relative.StartsWith(".." + Path.DirectorySeparatorChar, StringComparison.Ordinal)
               && !relative.StartsWith(".." + Path.AltDirectorySeparatorChar, StringComparison.Ordinal);
    }

    internal static bool PathsEqual(string left, string right) =>
        string.Equals(
            left,
            right,
            OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);

    private static readonly string[] RelativeDirectories =
    [
        "config",
        "schemas",
        "suites",
        "plans",
        "object-repository",
        "runs"
    ];

    private static readonly IReadOnlyDictionary<string, string> TemplateFiles =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["config/agentsettings.example.json"] = AgentSettingsExample,
            ["schemas/suite.schema.json"] = SuiteSchema,
            ["suites/smoke.json"] = EmptySuite("smoke"),
            ["suites/regression.json"] = EmptySuite("regression"),
            ["plans/README.md"] = PlansReadme,
            ["object-repository/README.md"] = ObjectRepositoryReadme,
            ["runs/.gitignore"] = RunsGitIgnore
        };

    private const string AgentSettingsExample =
        """
        {
          "Driver": {
            "BaseUrl": "http://127.0.0.1:33201",
            "BearerToken": "REPLACE_WITH_TOKEN_FROM_VERIFY",
            "VerifyUrl": "http://localhost:9102/verify",
            "RequestTimeoutSeconds": 20,
            "ExpectedCatalogSchemaVersion": 2,
            "AllowRemoteDriver": false
          },
          "Workspace": {
            "Root": "automation"
          },
          "Suites": {
            "JiraKeyPattern": "^[A-Z][A-Z0-9_]*-[0-9]+$"
          }
        }
        """;

    private const string SuiteSchema =
        """
        {
          "$schema": "https://json-schema.org/draft/2020-12/schema",
          "$id": "https://local/desktop-automation-agent/suite.schema.json",
          "title": "DesktopAutomationAgent Suite Manifest",
          "type": "object",
          "required": [ "schemaVersion", "name", "testCases" ],
          "additionalProperties": true,
          "properties": {
            "schemaVersion": { "const": 1 },
            "name": { "type": "string", "minLength": 1 },
            "enabled": { "type": "boolean", "default": true },
            "testCases": {
              "type": "array",
              "items": {
                "type": "object",
                "required": [ "jiraKey" ],
                "additionalProperties": true,
                "properties": {
                  "jiraKey": {
                    "type": "string",
                    "pattern": "^[A-Z][A-Z0-9_]*-[0-9]+$"
                  },
                  "enabled": { "type": "boolean", "default": true }
                }
              }
            }
          }
        }
        """;

    private static string EmptySuite(string name) =>
        $$"""
        {
          "schemaVersion": 1,
          "name": "{{name}}",
          "enabled": true,
          "testCases": []
        }
        """;

    private const string PlansReadme =
        """
        # Plans

        Reserved for reusable compiled command plans in a later phase.
        Do not store secrets here.
        """;

    private const string ObjectRepositoryReadme =
        """
        # Object repository

        Reserved for page and element definitions in a later phase.
        Do not store secrets here.
        """;

    private const string RunsGitIgnore =
        """
        *
        !.gitignore
        """;
}

public sealed class WorkspaceException : Exception
{
    public WorkspaceException(string message) : base(message)
    {
    }

    public WorkspaceException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}
