using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace DesktopAutomationAgent.Mcp;

/// <summary>
/// MCP protocol adapter tools. Each handler delegates once to <see cref="IDesktopMcpFacade"/>.
/// Categories are encoded in descriptions: observation / navigation / mutation / execution / validation.
/// There is intentionally no unrestricted desktop_execute tool.
/// </summary>
[McpServerToolType]
public sealed class DesktopMcpTools
{
    private static readonly HashSet<string> AllowedViews =
        new(StringComparer.OrdinalIgnoreCase) { "control", "content", "raw" };

    private static readonly HashSet<string> AllowedRoots =
        new(StringComparer.OrdinalIgnoreCase) { "activeWindow", "processWindows", "desktopChildren" };

    private readonly IDesktopMcpFacade _facade;

    public DesktopMcpTools(IDesktopMcpFacade facade)
    {
        _facade = facade;
    }

    [McpServerTool(Name = "desktop_doctor", Title = "Desktop doctor", ReadOnly = true, Destructive = false, Idempotent = true, OpenWorld = false)]
    [Description("Observation: run local readiness checks (workspace, driver discovery, status, catalog). Never returns driver bearer tokens or Authorization headers.")]
    public Task<CallToolResult> DesktopDoctor(CancellationToken cancellationToken) =>
        InvokeAsync(() => _facade.DoctorAsync(cancellationToken));

    [McpServerTool(Name = "desktop_observe", Title = "Desktop observe", ReadOnly = true, Destructive = false, Idempotent = true, OpenWorld = false)]
    [Description("Observation: capture a bounded semantic UI snapshot (no clicks/types). Returns compact nodes with opaque refs that expire; never a raw unlimited UIA tree.")]
    public Task<CallToolResult> DesktopObserve(
        [Description("control|content|raw")] string? view = null,
        [Description("activeWindow|processWindows|desktopChildren")] string? root = null,
        [Range(0, 12)] int? maxDepth = null,
        [Range(1, 200)] int? maxChildren = null,
        [Range(1, 500)] int? maxNodes = null,
        bool includeOffscreen = false,
        bool includeBounds = false,
        bool includeHelpText = false,
        bool? includeFocusState = null,
        bool? actionableOnly = null,
        [MaxLength(512)] string? outputPath = null,
        CancellationToken cancellationToken = default) =>
        InvokeAsync(() =>
        {
            ValidateOptionalEnum(view, nameof(view), AllowedViews);
            ValidateOptionalEnum(root, nameof(root), AllowedRoots);
            ValidateOptionalRange(maxDepth, nameof(maxDepth), 0, 12);
            ValidateOptionalRange(maxChildren, nameof(maxChildren), 1, 200);
            ValidateOptionalRange(maxNodes, nameof(maxNodes), 1, 500);
            return _facade.ObserveAsync(
                view,
                root,
                maxDepth,
                maxChildren,
                maxNodes,
                includeOffscreen,
                includeBounds,
                includeHelpText,
                includeFocusState,
                actionableOnly,
                OptionalArg(outputPath, nameof(outputPath), maxLength: 512),
                cancellationToken);
        });

    // Not ReadOnly: dry-run may write exploration artifact stubs. Not Idempotent: repeated calls can append artifacts.
    [McpServerTool(Name = "desktop_explore_dry_run", Title = "Desktop explore dry-run", ReadOnly = false, Destructive = false, Idempotent = false, OpenWorld = false)]
    [Description("Observation: list policy-permitted exploration candidates without UI mutation. Deny-by-default exploration policy required.")]
    public Task<CallToolResult> DesktopExploreDryRun(
        [Required, MaxLength(512), Description("Workspace-contained exploration policy path")] string policyPath,
        CancellationToken cancellationToken = default) =>
        InvokeAsync(() => _facade.ExploreAsync(RequireArg(policyPath, nameof(policyPath), maxLength: 512), dryRun: true, cancellationToken));

    [McpServerTool(Name = "desktop_explore", Title = "Desktop explore", ReadOnly = false, Destructive = false, Idempotent = false, OpenWorld = false)]
    [Description("Navigation: policy-bounded observe/navigate explorer. Deny-by-default; no typing or transactional actions. Not unrestricted execution.")]
    public Task<CallToolResult> DesktopExplore(
        [Required, MaxLength(512), Description("Workspace-contained exploration policy path")] string policyPath,
        CancellationToken cancellationToken = default) =>
        InvokeAsync(() => _facade.ExploreAsync(RequireArg(policyPath, nameof(policyPath), maxLength: 512), dryRun: false, cancellationToken));

    [McpServerTool(Name = "desktop_search_knowledge", Title = "Desktop search knowledge", ReadOnly = true, Destructive = false, Idempotent = true, OpenWorld = false)]
    [Description("Observation: deterministic lexical search over approved application knowledge (no embeddings).")]
    public Task<CallToolResult> DesktopSearchKnowledge(
        [Required, MaxLength(512)] string knowledgePath,
        [Required, MaxLength(256)] string query,
        [Range(1, 100)] int? limit = null,
        CancellationToken cancellationToken = default) =>
        InvokeAsync(() =>
        {
            cancellationToken.ThrowIfCancellationRequested();
            ValidateOptionalRange(limit, nameof(limit), 1, 100);
            return Task.FromResult(_facade.SearchKnowledge(
                RequireArg(knowledgePath, nameof(knowledgePath), maxLength: 512),
                RequireArg(query, nameof(query), maxLength: 256),
                limit));
        });

    [McpServerTool(Name = "desktop_validate_plan", Title = "Desktop validate plan", ReadOnly = true, Destructive = false, Idempotent = true, OpenWorld = false)]
    [Description("Validation: offline deterministic plan validation including optional object-repository expansion. No driver HTTP.")]
    public Task<CallToolResult> DesktopValidatePlan(
        [Required, MaxLength(512)] string planPath,
        CancellationToken cancellationToken = default) =>
        InvokeAsync(() =>
        {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult(_facade.ValidatePlan(RequireArg(planPath, nameof(planPath), maxLength: 512)));
        });

    // Not ReadOnly: compile writes plan-cache artifacts. OpenWorld=true: fromJira may contact remote Jira.
    [McpServerTool(Name = "desktop_compile_key", Title = "Desktop compile key", ReadOnly = false, Destructive = false, Idempotent = true, OpenWorld = true)]
    [Description("Validation: compile a Jira BDD key into a deterministic plan cache entry. fromJira performs read-only remote sync then compile (open-world); never POST /ui.")]
    public Task<CallToolResult> DesktopCompileKey(
        [Required, MaxLength(64)] string key,
        [MaxLength(512)] string? file = null,
        bool fromJira = false,
        [MaxLength(512)] string? registry = null,
        [MaxLength(512)] string? objectRepository = null,
        [MaxLength(512)] string? knowledge = null,
        [MaxLength(512)] string? catalog = null,
        bool explain = false,
        CancellationToken cancellationToken = default) =>
        InvokeAsync(async () =>
        {
            var safeKey = RequireArg(key, nameof(key), maxLength: 64);
            var safeFile = OptionalArg(file, nameof(file), maxLength: 512);
            var safeRegistry = OptionalArg(registry, nameof(registry), maxLength: 512);
            var safeObjectRepository = OptionalArg(objectRepository, nameof(objectRepository), maxLength: 512);
            var safeKnowledge = OptionalArg(knowledge, nameof(knowledge), maxLength: 512);
            var safeCatalog = OptionalArg(catalog, nameof(catalog), maxLength: 512);
            if (fromJira)
            {
                if (!string.IsNullOrWhiteSpace(safeFile))
                    throw new McpToolException(McpErrorCode.Validation, "fromJira cannot be combined with file.");
                return await _facade.CompileKeyFromJiraAsync(
                        safeKey, safeRegistry, safeObjectRepository, safeKnowledge, safeCatalog, explain, cancellationToken)
                    .ConfigureAwait(false);
            }

            cancellationToken.ThrowIfCancellationRequested();
            return _facade.CompileKey(
                safeKey, safeFile, fromJira: false, safeRegistry, safeObjectRepository, safeKnowledge, safeCatalog, explain);
        });

    [McpServerTool(Name = "desktop_run_plan", Title = "Desktop run plan", ReadOnly = false, Destructive = true, Idempotent = false, OpenWorld = false)]
    [Description("Execution: run a validated deterministic plan through the Phase 2 runner. dryRun defaults to true (omission never live-runs); dryRun=false executes against the live catalog/driver.")]
    public Task<CallToolResult> DesktopRunPlan(
        [Required, MaxLength(512)] string planPath,
        bool dryRun = true,
        CancellationToken cancellationToken = default) =>
        InvokeAsync(() => _facade.RunPlanAsync(RequireArg(planPath, nameof(planPath), maxLength: 512), dryRun, cancellationToken));

    // OpenWorld=true when fromJira syncs from remote Jira before execution.
    [McpServerTool(Name = "desktop_run_key", Title = "Desktop run key", ReadOnly = false, Destructive = true, Idempotent = false, OpenWorld = true)]
    [Description("Execution: compile (or reuse cache) then execute a Jira key via the Phase 2 runner. fromJira may contact remote Jira; dryRun defaults to true (omission never live-runs / never POST /ui).")]
    public Task<CallToolResult> DesktopRunKey(
        [Required, MaxLength(64)] string key,
        bool dryRun = true,
        bool fromJira = false,
        [MaxLength(512)] string? registry = null,
        [MaxLength(512)] string? objectRepository = null,
        [MaxLength(512)] string? knowledge = null,
        [MaxLength(512)] string? catalog = null,
        CancellationToken cancellationToken = default) =>
        InvokeAsync(() => _facade.RunKeyAsync(
            RequireArg(key, nameof(key), maxLength: 64),
            dryRun,
            fromJira,
            OptionalArg(registry, nameof(registry), maxLength: 512),
            OptionalArg(objectRepository, nameof(objectRepository), maxLength: 512),
            OptionalArg(knowledge, nameof(knowledge), maxLength: 512),
            OptionalArg(catalog, nameof(catalog), maxLength: 512),
            cancellationToken));

    [McpServerTool(Name = "desktop_run_suite", Title = "Desktop run suite", ReadOnly = false, Destructive = true, Idempotent = false, OpenWorld = true)]
    [Description("Execution: sequential suite orchestration. Fail-fast by default. fromJira may contact remote Jira; dryRun defaults to true (omission never live-runs / never POST /ui).")]
    public Task<CallToolResult> DesktopRunSuite(
        [Required, MaxLength(512)] string suitePath,
        bool dryRun = true,
        bool fromJira = false,
        bool continueOnFailure = false,
        [Range(1, 10_000)] int? maxFailures = null,
        [MaxLength(512)] string? registry = null,
        [MaxLength(512)] string? objectRepository = null,
        [MaxLength(512)] string? knowledge = null,
        [MaxLength(512)] string? catalog = null,
        CancellationToken cancellationToken = default) =>
        InvokeAsync(() =>
        {
            ValidateOptionalRange(maxFailures, nameof(maxFailures), 1, 10_000);
            return _facade.RunSuiteAsync(
                RequireArg(suitePath, nameof(suitePath), maxLength: 512),
                dryRun,
                fromJira,
                continueOnFailure,
                maxFailures,
                OptionalArg(registry, nameof(registry), maxLength: 512),
                OptionalArg(objectRepository, nameof(objectRepository), maxLength: 512),
                OptionalArg(knowledge, nameof(knowledge), maxLength: 512),
                OptionalArg(catalog, nameof(catalog), maxLength: 512),
                cancellationToken);
        });

    [McpServerTool(Name = "desktop_get_run_summary", Title = "Desktop get run summary", ReadOnly = true, Destructive = false, Idempotent = true, OpenWorld = false)]
    [Description("Observation: load a compact sanitized run or suite-run summary by runId. Secrets are redacted; unrestricted raw artifacts are not returned.")]
    public Task<CallToolResult> DesktopGetRunSummary(
        [Required, MaxLength(128)] string runId,
        CancellationToken cancellationToken = default) =>
        InvokeAsync(() =>
        {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult(_facade.GetRunSummary(RequireArg(runId, nameof(runId), maxLength: 128)));
        });

    private static string RequireArg(string? value, string name, int maxLength)
    {
        if (string.IsNullOrWhiteSpace(value))
            throw new McpToolException(McpErrorCode.Validation, $"{name} is required.");
        EnsureMaxLength(value, name, maxLength);
        return value;
    }

    private static string? OptionalArg(string? value, string name, int maxLength)
    {
        if (value is null)
            return null;
        if (string.IsNullOrWhiteSpace(value))
            return null;
        EnsureMaxLength(value, name, maxLength);
        return value;
    }

    private static void EnsureMaxLength(string value, string name, int maxLength)
    {
        if (value.Length > maxLength)
        {
            throw new McpToolException(
                McpErrorCode.Validation,
                $"{name} exceeds maximum length of {maxLength}.");
        }
    }

    private static void ValidateOptionalEnum(string? value, string name, HashSet<string> allowed)
    {
        if (value is null)
            return;
        if (string.IsNullOrWhiteSpace(value) || !allowed.Contains(value))
        {
            throw new McpToolException(
                McpErrorCode.Validation,
                $"{name} must be one of: {string.Join(", ", allowed.OrderBy(x => x, StringComparer.Ordinal))}.");
        }
    }

    private static void ValidateOptionalRange(int? value, string name, int min, int max)
    {
        if (value is null)
            return;
        if (value < min || value > max)
        {
            throw new McpToolException(
                McpErrorCode.Validation,
                $"{name} must be between {min} and {max}.");
        }
    }

    private static async Task<CallToolResult> InvokeAsync(Func<Task<object>> action)
    {
        try
        {
            var payload = await action().ConfigureAwait(false);
            return SuccessResult(payload);
        }
        catch (OperationCanceledException)
        {
            return ErrorResult(McpErrorCode.Cancelled, "Operation cancelled.");
        }
        catch (McpToolException ex)
        {
            return ErrorResult(ex.ErrorCode, Driver.SecretRedactor.Redact(ex.Message), Driver.SecretRedactor.Redact(ex.Detail));
        }
        catch (Exception ex)
        {
            return ErrorResult(McpErrorCode.Internal, Driver.SecretRedactor.Redact(ex.Message));
        }
    }

    private static CallToolResult SuccessResult(object payload)
    {
        // Redact first, then budget (or re-budget) so secrets cannot survive truncation.
        var json = McpResponseBudget.SerializeRedactedBounded(payload);
        EnsureWireJson(json);
        return new CallToolResult
        {
            Content = [new TextContentBlock { Text = json }],
            StructuredContent = JsonSerializer.Deserialize<JsonElement>(json)
        };
    }

    private static CallToolResult ErrorResult(string code, string message, string? detail = null)
    {
        var payload = new McpErrorPayload
        {
            Success = false,
            ErrorCode = code,
            Message = message,
            Detail = string.IsNullOrWhiteSpace(detail) ? null : detail
        };

        string json;
        try
        {
            json = McpResponseBudget.SerializeRedactedBounded(payload);
            EnsureWireJson(json);
        }
        catch (McpToolException)
        {
            json = Driver.SecretRedactor.Redact(
                JsonSerializer.Serialize(
                    new
                    {
                        success = false,
                        errorCode = McpErrorCode.Truncated,
                        message = "Error payload exceeded MCP response budget.",
                        truncated = true
                    },
                    McpJson.Options));
        }

        return new CallToolResult
        {
            IsError = true,
            Content = [new TextContentBlock { Text = json }],
            StructuredContent = JsonSerializer.Deserialize<JsonElement>(json)
        };
    }

    private static void EnsureWireJson(string json)
    {
        try
        {
            using var _ = JsonDocument.Parse(json);
        }
        catch (JsonException ex)
        {
            throw new McpToolException(
                McpErrorCode.Truncated,
                "Wire payload was not valid JSON after budgeting/redaction.",
                ex.Message);
        }
    }
}
