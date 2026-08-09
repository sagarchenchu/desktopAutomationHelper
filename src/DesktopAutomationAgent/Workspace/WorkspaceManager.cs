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

                File.WriteAllText(full, contents.EndsWith('\n') ? contents : contents + "\n");
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
        "object-repository/pages",
        "object-repository/candidates",
        "object-repository/captures",
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
            ["object-repository/repository.json"] = ObjectRepositoryManifest,
            ["object-repository/.gitignore"] = ObjectRepositoryGitIgnore,
            ["schemas/object-repository.schema.json"] = ObjectRepositorySchema,
            ["schemas/page-object.schema.json"] = PageObjectSchema,
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
          },
          "ObjectRepository": {
            "MaxFileBytes": 5242880,
            "MaxPages": 500,
            "MaxElementsPerPage": 5000,
            "MaxTotalElements": 50000,
            "DiagnosticTimeoutMilliseconds": 15000
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
          "$comment": "AUTHORITATIVE combined step limit: DesktopAutomationAgent PlanValidator enforces (steps + onFailureSteps) <= 1000. Per-array maxItems below are editor hints only and do not replace that combined rule. Reserved argument names are also rejected case-insensitively by PlanValidator; propertyNames patterns use portable ECMA-262 character classes.",
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
            "objectRepository": {
              "type": "string",
              "minLength": 1
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
            "reservedArgumentName": {
              "anyOf": [
                {
                  "pattern": "^[Oo][Pp][Ee][Rr][Aa][Tt][Ii][Oo][Nn]$"
                },
                {
                  "pattern": "^[Aa][Uu][Tt][Hh][Oo][Rr][Ii][Zz][Aa][Tt][Ii][Oo][Nn]$"
                },
                {
                  "pattern": "^[Bb][Ee][Aa][Rr][Ee][Rr][Tt][Oo][Kk][Ee][Nn]$"
                }
              ]
            },
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
                      "$ref": "#/$defs/reservedArgumentName"
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
                      "$ref": "#/$defs/reservedArgumentName"
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
        # Object repository (Phase 3)

        The object repository stores **approved, versioned UI locators** for deterministic plan execution.

        ## Layout

        ```text
        object-repository/
          repository.json          # manifest (tracked)
          pages/                   # active page-object documents (tracked)
          candidates/              # draft page objects awaiting promotion (gitignored)
          captures/                # raw capture output from tooling (gitignored)
        ```

        ## CLI workflow

        ```bash
        # Offline validation
        dotnet run --project src/DesktopAutomationAgent -- validate-object-repository \
          --file automation/object-repository/repository.json

        # Offline resolve one reference
        dotnet run --project src/DesktopAutomationAgent -- resolve-object \
          --file automation/object-repository/repository.json --ref login.submit

        # Live capture (dumpuia) — writes captures/ and candidates/
        dotnet run --project src/DesktopAutomationAgent -- capture-page \
          --file automation/object-repository/repository.json \
          --page login --name "Login page" \
          [--view control|content|raw] \
          [--root activeWindow|processWindows|desktopChildren] \
          [--max-depth 8] [--max-children 200] [--include-offscreen] \
          [--json]

        # Live verify (finduia) — all active objects, or filter
        dotnet run --project src/DesktopAutomationAgent -- verify-object-repository \
          --file automation/object-repository/repository.json \
          [--page login | --ref login.submit] [--json]
        ```

        Plans may reference repository objects via `$objectRef` in locator arguments when
        `objectRepository` is set on the plan. See `docs/phase3-object-repository.md`.

        ## Captures vs approved objects

        | Area | Purpose | Git |
        |------|---------|-----|
        | `captures/` | Raw, machine-generated locator snapshots from a capture session | Ignored |
        | `candidates/` | Human-reviewed drafts promoted from captures | Ignored |
        | `pages/` | Active page-object JSON referenced by `repository.json` | Tracked |

        **Captures are never executed directly.** Operators review captures, curate locators,
        and **manually promote** approved definitions into `pages/` with `state: "active"`.

        ## PII and secrets

        - Do **not** store passwords, tokens, customer data, or other PII in page objects.
        - Prefer stable `automationId` and structural locators over visible text that may contain names.
        - Review capture output before promotion; redact or generalize sensitive `name` values.

        ## Locator rules (enforced by the agent)

        Allowed locator fields:

        - `automationId`
        - `name` + `controlType`
        - `className` + `controlType`
        - `matchMode`: `exact`, `contains`, or `startswith`
        - `foundIndex` (discouraged; adds fragility warnings)

        Volatile fields (handles, coordinates, runtime IDs, bounding boxes, etc.) are **rejected**.

        `automationId` alone is sufficient. Without it, `name` and `controlType` **or** `className` and
        `controlType` are required.

        ## Manual promotion workflow

        1. Run `capture-page` to write artifacts under `captures/` and `candidates/`.
        2. Review and edit locators; iterate in `candidates/` if needed.
        3. Set `state` to `active`, `source.kind` to `manual` or `approved`, and add the page to
           `repository.json` under `pages/`.
        4. Run `verify-object-repository` against the promoted page.
        5. Commit only the manifest and `pages/` files. Never commit `captures/` or `candidates/`.

        Active pages must not contain `source.kind: "capture"` elements.

        ## No AI in Phase 3

        Phase 3 does **not** call AI providers, perform self-healing, or rewrite locators automatically.
        Validation, capture, verification, resolution, and plan expansion are fully deterministic.
        All approved objects are human-reviewed. GitLab/XAML/WinForms source extraction remains future work.

        ## Identifiers

        `repositoryId`, `pageId`, and element keys must match:

        ```text
        ^[a-z][a-z0-9-]{0,63}$
        ```

        Object references use `pageId.elementId` (for example `login.submit-button`).

        ## Schemas

        - `schemas/object-repository.schema.json` — manifest
        - `schemas/page-object.schema.json` — page documents

        Do not store secrets here.
        """;

    private const string ObjectRepositoryManifest =
        """
        {
          "$schema": "../schemas/object-repository.schema.json",
          "schemaVersion": 1,
          "repositoryId": "default",
          "name": "Default object repository",
          "pages": []
        }
        """;

    private const string ObjectRepositoryGitIgnore =
        """
        captures/**
        candidates/**
        """;

    private const string ObjectRepositorySchema =
        """
        {
          "$schema": "https://json-schema.org/draft/2020-12/schema",
          "$id": "https://local/desktop-automation-agent/object-repository.schema.json",
          "title": "DesktopAutomationAgent Object Repository Manifest",
          "$comment": "AUTHORITATIVE limits: DesktopAutomationAgent ObjectRepositoryValidator enforces MaxPages, MaxElementsPerPage, and MaxTotalElements from configuration. Identifier pattern is lowercase-only and enforced by ObjectRepositoryValidator.",
          "type": "object",
          "required": [
            "schemaVersion",
            "repositoryId",
            "name",
            "pages"
          ],
          "additionalProperties": false,
          "properties": {
            "$schema": {
              "type": "string"
            },
            "schemaVersion": {
              "const": 1
            },
            "repositoryId": {
              "type": "string",
              "pattern": "^[a-z][a-z0-9-]{0,63}$"
            },
            "name": {
              "type": "string",
              "minLength": 1
            },
            "pages": {
              "type": "array",
              "maxItems": 500,
              "items": {
                "$ref": "#/$defs/pageReference"
              }
            }
          },
          "$defs": {
            "pageReference": {
              "type": "object",
              "required": [
                "pageId",
                "file"
              ],
              "additionalProperties": false,
              "properties": {
                "pageId": {
                  "type": "string",
                  "pattern": "^[a-z][a-z0-9-]{0,63}$"
                },
                "file": {
                  "type": "string",
                  "minLength": 1
                }
              }
            }
          }
        }
        """;

    private const string PageObjectSchema =
        """
        {
          "$schema": "https://json-schema.org/draft/2020-12/schema",
          "$id": "https://local/desktop-automation-agent/page-object.schema.json",
          "title": "DesktopAutomationAgent Page Object Document",
          "$comment": "AUTHORITATIVE locator rules: DesktopAutomationAgent ObjectLocatorValidator enforces allowed locator fields, volatile property rejection, and combination rules. Active pages must not contain capture-sourced elements.",
          "type": "object",
          "required": [
            "schemaVersion",
            "pageId",
            "name",
            "state",
            "elements"
          ],
          "additionalProperties": false,
          "properties": {
            "$schema": {
              "type": "string"
            },
            "schemaVersion": {
              "const": 1
            },
            "pageId": {
              "type": "string",
              "pattern": "^[a-z][a-z0-9-]{0,63}$"
            },
            "name": {
              "type": "string",
              "minLength": 1
            },
            "state": {
              "type": "string",
              "enum": [
                "candidate",
                "active"
              ]
            },
            "elements": {
              "type": "object",
              "maxProperties": 5000,
              "propertyNames": {
                "pattern": "^[a-z][a-z0-9-]{0,63}$"
              },
              "additionalProperties": {
                "$ref": "#/$defs/element"
              }
            },
            "unresolved": {
              "type": "array",
              "items": {
                "type": "object"
              }
            }
          },
          "$defs": {
            "element": {
              "type": "object",
              "required": [
                "locator"
              ],
              "additionalProperties": false,
              "properties": {
                "description": {
                  "type": "string"
                },
                "locator": {
                  "$ref": "#/$defs/locator"
                },
                "quality": {
                  "$ref": "#/$defs/quality"
                },
                "source": {
                  "$ref": "#/$defs/source"
                }
              }
            },
            "locator": {
              "type": "object",
              "additionalProperties": false,
              "properties": {
                "automationId": {
                  "type": "string",
                  "minLength": 1
                },
                "name": {
                  "type": "string",
                  "minLength": 1
                },
                "className": {
                  "type": "string",
                  "minLength": 1
                },
                "controlType": {
                  "type": "string",
                  "minLength": 1
                },
                "matchMode": {
                  "type": "string",
                  "enum": [
                    "exact",
                    "contains",
                    "startswith"
                  ]
                },
                "foundIndex": {
                  "type": "integer",
                  "minimum": 0
                }
              }
            },
            "quality": {
              "type": "object",
              "additionalProperties": false,
              "properties": {
                "grade": {
                  "type": "string",
                  "enum": [
                    "strong",
                    "medium",
                    "weak"
                  ]
                },
                "warnings": {
                  "type": "array",
                  "items": {
                    "type": "string"
                  }
                }
              }
            },
            "source": {
              "type": "object",
              "required": [
                "kind"
              ],
              "additionalProperties": false,
              "properties": {
                "kind": {
                  "type": "string",
                  "enum": [
                    "capture",
                    "manual",
                    "approved"
                  ]
                },
                "path": {
                  "type": "string"
                },
                "metadata": {
                  "type": "object",
                  "additionalProperties": true
                }
              }
            }
          }
        }
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
