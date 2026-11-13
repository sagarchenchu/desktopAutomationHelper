namespace DesktopAutomationAgent.Mcp;

/// <summary>
/// Application-service façade for MCP handlers. Each method delegates once into existing
/// agent services — no duplicated driver clients, validators, compiler, policy, or runners.
/// </summary>
public interface IDesktopMcpFacade
{
    Task<object> DoctorAsync(CancellationToken cancellationToken);

    Task<object> ObserveAsync(
        string? view,
        string? root,
        int? maxDepth,
        int? maxChildren,
        int? maxNodes,
        bool includeOffscreen,
        bool includeBounds,
        bool includeHelpText,
        bool? includeFocusState,
        bool? actionableOnly,
        string? outputPath,
        CancellationToken cancellationToken);

    Task<object> ExploreAsync(string policyPath, bool dryRun, CancellationToken cancellationToken);

    object SearchKnowledge(string knowledgePath, string query, int? limit);

    object ValidatePlan(string planPath);

    object CompileKey(
        string key,
        string? file,
        bool fromJira,
        string? registry,
        string? objectRepository,
        string? knowledge,
        string? catalog,
        bool explain);

    Task<object> CompileKeyFromJiraAsync(
        string key,
        string? registry,
        string? objectRepository,
        string? knowledge,
        string? catalog,
        bool explain,
        CancellationToken cancellationToken);

    Task<object> RunPlanAsync(string planPath, bool dryRun, CancellationToken cancellationToken);

    Task<object> RunKeyAsync(
        string key,
        bool dryRun,
        bool fromJira,
        string? registry,
        string? objectRepository,
        string? knowledge,
        string? catalog,
        CancellationToken cancellationToken);

    Task<object> RunSuiteAsync(
        string suitePath,
        bool dryRun,
        bool fromJira,
        bool continueOnFailure,
        int? maxFailures,
        string? registry,
        string? objectRepository,
        string? knowledge,
        string? catalog,
        CancellationToken cancellationToken);

    object GetRunSummary(string runId);

    /// <summary>Test/support: resolve an opaque observe ref or throw stale_ref.</summary>
    object ResolveOpaqueRef(string snapshotId, string nodeRef);
}
