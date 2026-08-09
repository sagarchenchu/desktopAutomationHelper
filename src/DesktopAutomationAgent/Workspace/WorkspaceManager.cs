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
            ["schemas/plan.schema.json"] = PlanSchema,
            ["suites/smoke.json"] = EmptySuite("smoke"),
            ["suites/regression.json"] = EmptySuite("regression"),
            ["plans/README.md"] = PlansReadme,
            ["plans/example.plan.json"] = ExamplePlan,
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
          },
          "Runner": {
            "StepTransportTimeoutSeconds": 60,
            "CleanupTimeoutSeconds": 15,
            "MaxPlanBytes": 1048576,
            "MaxResponseBytes": 10485760,
            "RegexTimeoutMilliseconds": 500
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

        Plans are compiled executable artifacts for the Desktop Automation Agent
        deterministic runner (Phase 2).

        ## Layout

        - `example.plan.json` — session-free smoke example that calls `listwindows` only.
          The driver returns a root JSON array of window descriptors; assert `path: ""` / `isNotNull`.
          Do not pass `limit` unless/until the driver supports it.
        - `../schemas/plan.schema.json` — JSON Schema (Draft 2020-12) for offline validation.

        ## Phase 2 authoring

        - Phase 2 supports **manually authored** plans.
        - Later phases will compile Jira BDD into the same format.
        - Existing valid plans execute without AI, Jira, object-repository, or database access.
        - `schemaVersion` must be `1` and `catalogSchemaVersion` must be `2`.
        - `planId` must match `^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$`.
        - `steps` is required and must contain at least one step. `onFailureSteps` is optional.
        - Combined `steps` + `onFailureSteps` count must not exceed `1000`.
        - Step `id` values must be unique (case-insensitive) across all step lists.
        - Each step requires `operation` (no leading/trailing whitespace) and an `arguments` object.
        - Do not place `operation`, `authorization`, or `bearerToken` inside `arguments`.
        - Cleanup steps must not define `assertions` or `captureResponse`.
        - Plans that call `launch` must end with `close` or `quit`, and `onFailureSteps` must include `close` or `quit`. `closewindow` does not end a session.
        - Plans must not contain credentials. Password-entry steps must use `sensitive: true`.

        ```bash
        dotnet run --project src/DesktopAutomationAgent -- validate-plan --file automation/plans/example.plan.json
        dotnet run --project src/DesktopAutomationAgent -- run-plan --file automation/plans/example.plan.json --dry-run
        ```
        """;

    private const string ExamplePlan =
        """
        {
          "$schema": "../schemas/plan.schema.json",
          "schemaVersion": 1,
          "catalogSchemaVersion": 2,
          "planId": "example.listwindows",
          "name": "List visible windows",
          "description": "Session-free smoke plan. listwindows returns a root JSON array of window descriptors.",
          "steps": [
            {
              "id": "list-windows",
              "operation": "listwindows",
              "arguments": {
                "includeDesktopDescendants": false
              },
              "captureResponse": true,
              "assertions": [
                {
                  "path": "",
                  "operator": "isNotNull"
                }
              ]
            }
          ]
        }
        """;

    private const string PlanSchema =
        """
        {
          "$schema": "https://json-schema.org/draft/2020-12/schema",
          "$id": "https://local/desktop-automation-agent/plan.schema.json",
          "title": "DesktopAutomationAgent Plan Manifest",
          "type": "object",
          "required": [
            "schemaVersion",
            "catalogSchemaVersion",
            "planId",
            "name",
            "steps"
          ],
          "additionalProperties": false,
          "properties": {
            "$schema": {
              "type": "string"
            },
            "schemaVersion": {
              "const": 1
            },
            "catalogSchemaVersion": {
              "const": 2
            },
            "planId": {
              "type": "string",
              "pattern": "^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$"
            },
            "name": {
              "type": "string",
              "minLength": 1
            },
            "description": {
              "type": "string"
            },
            "tags": {
              "type": "array",
              "items": {
                "type": "string"
              }
            },
            "metadata": {
              "type": "object",
              "additionalProperties": true
            },
            "steps": {
              "type": "array",
              "minItems": 1,
              "maxItems": 1000,
              "items": {
                "$ref": "#/$defs/mainStep"
              }
            },
            "onFailureSteps": {
              "type": "array",
              "maxItems": 1000,
              "items": {
                "$ref": "#/$defs/cleanupStep"
              }
            }
          },
          "$defs": {
            "mainStep": {
              "type": "object",
              "required": [
                "id",
                "operation",
                "arguments"
              ],
              "additionalProperties": false,
              "properties": {
                "id": {
                  "type": "string",
                  "minLength": 1
                },
                "operation": {
                  "type": "string",
                  "minLength": 1,
                  "pattern": "^\\S(.*\\S)?$"
                },
                "arguments": {
                  "type": "object",
                  "additionalProperties": true,
                  "propertyNames": {
                    "not": {
                      "enum": [
                        "operation",
                        "authorization",
                        "bearerToken",
                        "Operation",
                        "Authorization",
                        "BearerToken"
                      ]
                    }
                  }
                },
                "assertions": {
                  "type": "array",
                  "items": {
                    "$ref": "#/$defs/planAssertion"
                  }
                },
                "sensitive": {
                  "type": "boolean",
                  "default": false
                },
                "captureResponse": {
                  "type": "boolean",
                  "default": false
                }
              }
            },
            "cleanupStep": {
              "type": "object",
              "required": [
                "id",
                "operation",
                "arguments"
              ],
              "additionalProperties": false,
              "properties": {
                "id": {
                  "type": "string",
                  "minLength": 1
                },
                "operation": {
                  "type": "string",
                  "minLength": 1,
                  "pattern": "^\\S(.*\\S)?$"
                },
                "arguments": {
                  "type": "object",
                  "additionalProperties": true,
                  "propertyNames": {
                    "not": {
                      "enum": [
                        "operation",
                        "authorization",
                        "bearerToken",
                        "Operation",
                        "Authorization",
                        "BearerToken"
                      ]
                    }
                  }
                },
                "sensitive": {
                  "type": "boolean",
                  "default": false
                }
              }
            },
            "planAssertion": {
              "type": "object",
              "required": [
                "operator"
              ],
              "additionalProperties": false,
              "properties": {
                "path": {
                  "type": "string",
                  "default": ""
                },
                "operator": {
                  "type": "string",
                  "enum": [
                    "equals",
                    "notEquals",
                    "contains",
                    "matchesRegex",
                    "isTrue",
                    "isFalse",
                    "isNull",
                    "isNotNull"
                  ]
                },
                "expected": {},
                "ignoreCase": {
                  "type": "boolean",
                  "default": false
                }
              },
              "allOf": [
                {
                  "if": {
                    "properties": {
                      "operator": {
                        "enum": [
                          "equals",
                          "notEquals",
                          "contains",
                          "matchesRegex"
                        ]
                      }
                    },
                    "required": [
                      "operator"
                    ]
                  },
                  "then": {
                    "required": [
                      "operator",
                      "expected"
                    ]
                  }
                }
              ]
            }
          }
        }
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
