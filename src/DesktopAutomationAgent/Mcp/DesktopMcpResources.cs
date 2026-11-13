using System.ComponentModel;
using System.Text.Json;
using DesktopAutomationAgent.Compilation;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Knowledge;
using DesktopAutomationAgent.Workspace;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace DesktopAutomationAgent.Mcp;

/// <summary>
/// Optional MCP resources: operation catalog summary, sanitized knowledge summary,
/// compilation explanation, run summary. Never secrets or unrestricted raw artifacts.
/// </summary>
[McpServerResourceType]
public sealed class DesktopMcpResources
{
    private readonly IWorkspaceManager _workspace;
    private readonly KnowledgeService _knowledge;
    private readonly BddCompilationService _compiler;
    private readonly IDesktopMcpFacade _facade;

    public DesktopMcpResources(
        IWorkspaceManager workspace,
        KnowledgeService knowledge,
        BddCompilationService compiler,
        IDesktopMcpFacade facade)
    {
        _workspace = workspace;
        _knowledge = knowledge;
        _compiler = compiler;
        _facade = facade;
    }

    [McpServerResource(UriTemplate = "desktop://catalog/summary", Name = "operation_catalog_summary", MimeType = "application/json")]
    [Description("Sanitized summary of the local step registry / operation catalog (names only).")]
    public string OperationCatalogSummary()
    {
        try
        {
            var registry = _compiler.ValidateStepRegistry();
            var payload = new
            {
                success = registry.IsValid,
                path = "step-registry/registry.json",
                registryId = registry.Document?.RegistryId,
                mappingCount = registry.Document?.Mappings.Count,
                errors = registry.Errors.Select(SecretRedactor.Redact).Take(20).ToArray()
            };
            return SecretRedactor.Redact(JsonSerializer.Serialize(payload, McpJson.Options));
        }
        catch (Exception ex)
        {
            return SecretRedactor.Redact(JsonSerializer.Serialize(new
            {
                success = false,
                errorCode = McpErrorCode.Validation,
                message = ex.Message
            }, McpJson.Options));
        }
    }

    [McpServerResource(UriTemplate = "desktop://knowledge/{knowledgeId}/summary", Name = "knowledge_summary", MimeType = "application/json")]
    [Description("Sanitized application-knowledge summary (counts and fingerprint only).")]
    public string KnowledgeSummary(string knowledgeId)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(knowledgeId) || knowledgeId.Contains("..", StringComparison.Ordinal))
                throw new McpToolException(McpErrorCode.PathTraversal, "Invalid knowledgeId.");

            var path = McpPathGuard.RequireWorkspacePath(
                _workspace,
                Path.Combine("knowledge", knowledgeId),
                "knowledgeId");
            var doc = _knowledge.LoadDocument(path);
            var payload = new
            {
                success = true,
                applicationId = doc.ApplicationId,
                buildFingerprint = doc.BuildFingerprint,
                status = doc.Status,
                windowCount = doc.Graph.Windows.Count,
                stateCount = doc.Graph.States.Count,
                elementCount = doc.Graph.Elements.Count,
                transitionCount = doc.Graph.Transitions.Count,
                workflowCount = doc.Workflows.Count
            };
            return SecretRedactor.Redact(JsonSerializer.Serialize(payload, McpJson.Options));
        }
        catch (Exception ex)
        {
            return SecretRedactor.Redact(JsonSerializer.Serialize(new
            {
                success = false,
                errorCode = McpErrorCode.Validation,
                message = ex.Message
            }, McpJson.Options));
        }
    }

    [McpServerResource(UriTemplate = "desktop://runs/{runId}/summary", Name = "run_summary", MimeType = "application/json")]
    [Description("Compact sanitized run summary (same as desktop_get_run_summary).")]
    public string RunSummary(string runId)
    {
        try
        {
            var payload = _facade.GetRunSummary(runId);
            return McpResponseBudget.SerializeBounded(payload);
        }
        catch (Exception ex)
        {
            return SecretRedactor.Redact(JsonSerializer.Serialize(new
            {
                success = false,
                errorCode = McpErrorCode.Validation,
                message = ex.Message
            }, McpJson.Options));
        }
    }
}
