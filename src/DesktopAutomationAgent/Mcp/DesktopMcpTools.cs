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
        InvokeAsync(() => _facade.ObserveAsync(
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
            outputPath,
            cancellationToken));

    [McpServerTool(Name = "desktop_explore_dry_run", Title = "Desktop explore dry-run", ReadOnly = true, Destructive = false, Idempotent = true, OpenWorld = false)]
    [Description("Observation: list policy-permitted exploration candidates without UI mutation. Deny-by-default exploration policy required.")]
    public Task<CallToolResult> DesktopExploreDryRun(
        [Required, MaxLength(512), Description("Workspace-contained exploration policy path")] string policyPath,
        CancellationToken cancellationToken) =>
        InvokeAsync(() => _facade.ExploreAsync(policyPath, dryRun: true, cancellationToken));

    [McpServerTool(Name = "desktop_explore", Title = "Desktop explore", ReadOnly = false, Destructive = false, Idempotent = false, OpenWorld = false)]
    [Description("Navigation: policy-bounded observe/navigate explorer. Deny-by-default; no typing or transactional actions. Not unrestricted execution.")]
    public Task<CallToolResult> DesktopExplore(
        [Required, MaxLength(512), Description("Workspace-contained exploration policy path")] string policyPath,
        CancellationToken cancellationToken) =>
        InvokeAsync(() => _facade.ExploreAsync(policyPath, dryRun: false, cancellationToken));

    [McpServerTool(Name = "desktop_search_knowledge", Title = "Desktop search knowledge", ReadOnly = true, Destructive = false, Idempotent = true, OpenWorld = false)]
    [Description("Observation: deterministic lexical search over approved application knowledge (no embeddings).")]
    public Task<CallToolResult> DesktopSearchKnowledge(
        [Required, MaxLength(512)] string knowledgePath,
        [Required, MaxLength(256)] string query,
        [Range(1, 100)] int? limit,
        CancellationToken cancellationToken) =>
        InvokeAsync(() =>
        {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult(_facade.SearchKnowledge(knowledgePath, query, limit));
        });

    [McpServerTool(Name = "desktop_validate_plan", Title = "Desktop validate plan", ReadOnly = true, Destructive = false, Idempotent = true, OpenWorld = false)]
    [Description("Validation: offline deterministic plan validation including optional object-repository expansion. No driver HTTP.")]
    public Task<CallToolResult> DesktopValidatePlan(
        [Required, MaxLength(512)] string planPath,
        CancellationToken cancellationToken) =>
        InvokeAsync(() =>
        {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult(_facade.ValidatePlan(planPath));
        });

    // OpenWorld=true: fromJira may contact an external Jira host (read-only). Offline compiles stay local.
    [McpServerTool(Name = "desktop_compile_key", Title = "Desktop compile key", ReadOnly = true, Destructive = false, Idempotent = true, OpenWorld = true)]
    [Description("Validation: compile a Jira BDD key into a deterministic plan cache entry. fromJira performs read-only remote sync then compile (open-world); never POST /ui.")]
    public Task<CallToolResult> DesktopCompileKey(
        [Required, MaxLength(64)] string key,
        [MaxLength(512)] string? file,
        bool fromJira,
        [MaxLength(512)] string? registry,
        [MaxLength(512)] string? objectRepository,
        [MaxLength(512)] string? knowledge,
        [MaxLength(512)] string? catalog,
        bool explain,
        CancellationToken cancellationToken) =>
        InvokeAsync(async () =>
        {
            if (fromJira)
            {
                if (!string.IsNullOrWhiteSpace(file))
                    throw new McpToolException(McpErrorCode.Validation, "fromJira cannot be combined with file.");
                return await _facade.CompileKeyFromJiraAsync(
                        key, registry, objectRepository, knowledge, catalog, explain, cancellationToken)
                    .ConfigureAwait(false);
            }

            cancellationToken.ThrowIfCancellationRequested();
            return _facade.CompileKey(key, file, fromJira: false, registry, objectRepository, knowledge, catalog, explain);
        });

    [McpServerTool(Name = "desktop_run_plan", Title = "Desktop run plan", ReadOnly = false, Destructive = true, Idempotent = false, OpenWorld = false)]
    [Description("Execution: run a validated deterministic plan through the Phase 2 runner. dryRun=true validates against the live catalog without POST /ui.")]
    public Task<CallToolResult> DesktopRunPlan(
        [Required, MaxLength(512)] string planPath,
        bool dryRun = false,
        CancellationToken cancellationToken = default) =>
        InvokeAsync(() => _facade.RunPlanAsync(planPath, dryRun, cancellationToken));

    // OpenWorld=true when fromJira syncs from remote Jira before execution.
    [McpServerTool(Name = "desktop_run_key", Title = "Desktop run key", ReadOnly = false, Destructive = true, Idempotent = false, OpenWorld = true)]
    [Description("Execution: compile (or reuse cache) then execute a Jira key via the Phase 2 runner. fromJira may contact remote Jira; dryRun never calls POST /ui.")]
    public Task<CallToolResult> DesktopRunKey(
        [Required, MaxLength(64)] string key,
        bool dryRun,
        bool fromJira,
        [MaxLength(512)] string? registry,
        [MaxLength(512)] string? objectRepository,
        [MaxLength(512)] string? knowledge,
        [MaxLength(512)] string? catalog,
        CancellationToken cancellationToken) =>
        InvokeAsync(() => _facade.RunKeyAsync(
            key, dryRun, fromJira, registry, objectRepository, knowledge, catalog, cancellationToken));

    [McpServerTool(Name = "desktop_run_suite", Title = "Desktop run suite", ReadOnly = false, Destructive = true, Idempotent = false, OpenWorld = true)]
    [Description("Execution: sequential suite orchestration. Fail-fast by default. fromJira may contact remote Jira; dryRun never calls POST /ui.")]
    public Task<CallToolResult> DesktopRunSuite(
        [Required, MaxLength(512)] string suitePath,
        bool dryRun,
        bool fromJira,
        bool continueOnFailure,
        [Range(1, 10_000)] int? maxFailures,
        [MaxLength(512)] string? registry,
        [MaxLength(512)] string? objectRepository,
        [MaxLength(512)] string? knowledge,
        [MaxLength(512)] string? catalog,
        CancellationToken cancellationToken) =>
        InvokeAsync(() => _facade.RunSuiteAsync(
            suitePath,
            dryRun,
            fromJira,
            continueOnFailure,
            maxFailures,
            registry,
            objectRepository,
            knowledge,
            catalog,
            cancellationToken));

    [McpServerTool(Name = "desktop_get_run_summary", Title = "Desktop get run summary", ReadOnly = true, Destructive = false, Idempotent = true, OpenWorld = false)]
    [Description("Observation: load a compact sanitized run or suite-run summary by runId. Secrets are redacted; unrestricted raw artifacts are not returned.")]
    public Task<CallToolResult> DesktopGetRunSummary(
        [Required, MaxLength(128)] string runId,
        CancellationToken cancellationToken) =>
        InvokeAsync(() =>
        {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult(_facade.GetRunSummary(runId));
        });

    private static async Task<CallToolResult> InvokeAsync(Func<Task<object>> action)
    {
        try
        {
            var payload = await action().ConfigureAwait(false);
            var json = McpResponseBudget.SerializeBounded(payload);
            // Final redaction pass — never return bearer/Authorization material via MCP.
            json = Driver.SecretRedactor.Redact(json);
            return new CallToolResult
            {
                Content = [new TextContentBlock { Text = json }],
                StructuredContent = JsonSerializer.Deserialize<JsonElement>(json)
            };
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
            // Redact before placing on the wire — never leak bearer/Authorization via MCP errors.
            return ErrorResult(McpErrorCode.Internal, Driver.SecretRedactor.Redact(ex.Message));
        }
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
        var json = Driver.SecretRedactor.Redact(JsonSerializer.Serialize(payload, McpJson.Options));
        return new CallToolResult
        {
            IsError = true,
            Content = [new TextContentBlock { Text = json }],
            StructuredContent = JsonSerializer.Deserialize<JsonElement>(json)
        };
    }
}
