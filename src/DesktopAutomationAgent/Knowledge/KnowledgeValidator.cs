using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Json.Schema;

namespace DesktopAutomationAgent.Knowledge;

public sealed class KnowledgeValidator
{
    private static readonly Regex ApplicationIdPattern = new(
        @"^[a-z][a-z0-9-]{0,63}$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex Hex64 = new(
        @"^[a-f0-9]{64}$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private readonly string? _schemaDirectory;

    public KnowledgeValidator(string? schemaDirectory = null)
    {
        _schemaDirectory = schemaDirectory;
    }

    public KnowledgeValidationResult Validate(ApplicationKnowledgeDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);
        var result = new KnowledgeValidationResult();

        if (document.KnowledgeSchemaVersion != KnowledgeConstants.KnowledgeSchemaVersion)
            result.Errors.Add($"Unsupported knowledgeSchemaVersion '{document.KnowledgeSchemaVersion}'.");

        if (!ApplicationIdPattern.IsMatch(document.ApplicationId ?? string.Empty))
            result.Errors.Add("applicationId must match ^[a-z][a-z0-9-]{0,63}$.");

        if (string.IsNullOrWhiteSpace(document.KnowledgeId))
            result.Errors.Add("knowledgeId is required.");

        if (string.IsNullOrWhiteSpace(document.BuildFingerprint))
            result.Errors.Add("buildFingerprint is required.");

        if (document.Status is not ("candidate" or "approved"))
            result.Errors.Add("status must be candidate or approved.");

        if (document.SourceExplorationRuns.Count == 0)
            result.Errors.Add("sourceExplorationRuns must contain at least one entry.");

        ValidateFingerprint(document.Fingerprint, document.ApplicationId ?? string.Empty, result);
        ValidateGraph(document.Graph, result);
        ValidateWorkflows(document.Workflows, document.Graph, result);
        ValidateCoverage(document.Coverage, result);

        if (document.Status == "approved")
        {
            if (document.Approval is null)
                result.Errors.Add("approved knowledge requires an approval record.");
            else
            {
                ValidateApproval(document.Approval, result);
                VerifyApprovedContentHash(document, result);
            }
        }

        if (document.Approval is not null && document.Approval.PromotesPageObjects)
            result.Errors.Add("approval.promotesPageObjects must be false; knowledge never promotes Page Objects.");

        // Schema evaluation when schema files are available.
        if (!string.IsNullOrWhiteSpace(_schemaDirectory))
            EvaluateSchemas(document, result);

        return result;
    }

    public KnowledgeValidationResult ValidateDirectory(string knowledgeDirectory)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(knowledgeDirectory);
        var path = Path.Combine(knowledgeDirectory, "application-knowledge.json");
        if (!File.Exists(path))
            return new KnowledgeValidationResult { Errors = { $"Missing application-knowledge.json under '{knowledgeDirectory}'." } };

        var json = File.ReadAllText(path);
        var document = JsonSerializer.Deserialize<ApplicationKnowledgeDocument>(json, KnowledgeJson.Compact)
                       ?? throw new InvalidOperationException("Failed to deserialize application-knowledge.json.");

        var result = Validate(document);

        // Cross-check optional content-hash.txt when present.
        var hashFile = Path.Combine(knowledgeDirectory, "content-hash.txt");
        if (File.Exists(hashFile))
        {
            var recorded = File.ReadAllText(hashFile).Trim().ToLowerInvariant();
            var actual = KnowledgeArtifactWriter.ComputeContentHash(document);
            if (!string.Equals(recorded, actual, StringComparison.Ordinal))
            {
                result.Errors.Add(
                    $"content-hash.txt mismatch: recorded '{recorded}', recomputed '{actual}'.");
            }
        }

        return result;
    }

    private static void VerifyApprovedContentHash(
        ApplicationKnowledgeDocument document,
        KnowledgeValidationResult result)
    {
        if (document.Approval is null)
            return;

        var actual = KnowledgeArtifactWriter.ComputeContentHash(document);
        if (!string.Equals(
                actual,
                document.Approval.SourceKnowledgeHash.Trim().ToLowerInvariant(),
                StringComparison.Ordinal))
        {
            result.Errors.Add(
                $"approved knowledge sourceKnowledgeHash mismatch: approval has "
                + $"'{document.Approval.SourceKnowledgeHash}', recomputed canonical hash is '{actual}'. "
                + "Document may have been tampered with after review.");
        }
    }

    private void EvaluateSchemas(ApplicationKnowledgeDocument document, KnowledgeValidationResult result)
    {
        try
        {
            var knowledgeSchemaPath = Path.Combine(_schemaDirectory!, "application-knowledge.schema.json");
            if (!File.Exists(knowledgeSchemaPath))
            {
                result.Warnings.Add("application-knowledge.schema.json not found; skipped JSON Schema evaluation.");
                return;
            }

            var options = new EvaluationOptions { OutputFormat = OutputFormat.List };
            foreach (var sibling in new[]
                     {
                         "application-state-graph.schema.json",
                         "application-workflow.schema.json",
                         "knowledge-approval.schema.json",
                         "application-knowledge.schema.json"
                     })
            {
                var path = Path.Combine(_schemaDirectory!, sibling);
                if (!File.Exists(path))
                    continue;

                var siblingSchema = JsonSchema.FromText(File.ReadAllText(path));
                if (siblingSchema.BaseUri is not null)
                    options.SchemaRegistry.Register(siblingSchema);
            }

            var schema = JsonSchema.FromText(File.ReadAllText(knowledgeSchemaPath));
            var node = JsonNode.Parse(JsonSerializer.Serialize(document, KnowledgeJson.Compact));
            var eval = schema.Evaluate(node, options);
            if (!eval.IsValid)
            {
                foreach (var detail in eval.Details.Where(d => d.Errors is not null))
                {
                    foreach (var error in detail.Errors!.Values)
                        result.Errors.Add($"schema:{detail.InstanceLocation}: {error}");
                }

                if (result.Errors.All(e => !e.StartsWith("schema:", StringComparison.Ordinal)))
                    result.Errors.Add("application-knowledge.json failed JSON Schema validation.");
            }
        }
        catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException or RefResolutionException)
        {
            // Semantic validation already ran; unresolved external $refs must not block offline builds.
            result.Warnings.Add($"JSON Schema evaluation skipped: {ex.Message}");
        }
    }

    private static void ValidateFingerprint(
        ApplicationFingerprint fingerprint,
        string applicationId,
        KnowledgeValidationResult result)
    {
        if (!string.Equals(fingerprint.ApplicationId, applicationId, StringComparison.Ordinal))
            result.Errors.Add("fingerprint.applicationId must match document applicationId.");

        if (fingerprint.SnapshotSchemaVersion != 1)
            result.Errors.Add("fingerprint.snapshotSchemaVersion must be 1.");

        if (string.IsNullOrWhiteSpace(fingerprint.ExplorerVersion))
            result.Errors.Add("fingerprint.explorerVersion is required.");

        if (string.IsNullOrWhiteSpace(fingerprint.DriverVersion))
            result.Errors.Add("fingerprint.driverVersion is required.");

        if (!Hex64.IsMatch(fingerprint.ExplorationPolicyHash ?? string.Empty))
            result.Errors.Add("fingerprint.explorationPolicyHash must be a 64-char lowercase hex digest.");

        if (!fingerprint.ConfigIdentity.ContainsKey("processName")
            || string.IsNullOrWhiteSpace(fingerprint.ConfigIdentity["processName"]))
        {
            result.Errors.Add("fingerprint.configIdentity.processName is required.");
        }

        if (fingerprint.ExecutableSha256 is not null
            && fingerprint.ExecutableSha256.Length > 0
            && !Hex64.IsMatch(fingerprint.ExecutableSha256))
        {
            result.Errors.Add("fingerprint.executableSha256 must be a 64-char lowercase hex digest when set.");
        }

        foreach (var key in fingerprint.ConfigIdentity.Keys)
        {
            if (key.Contains("secret", StringComparison.OrdinalIgnoreCase)
                || key.Contains("token", StringComparison.OrdinalIgnoreCase)
                || key.Contains("password", StringComparison.OrdinalIgnoreCase))
            {
                result.Errors.Add($"fingerprint.configIdentity key '{key}' looks like a secret and is forbidden.");
            }
        }
    }

    private static void ValidateGraph(ApplicationStateGraph graph, KnowledgeValidationResult result)
    {
        if (graph.GraphSchemaVersion != KnowledgeConstants.GraphSchemaVersion)
            result.Errors.Add($"Unsupported graphSchemaVersion '{graph.GraphSchemaVersion}'.");

        var stateIds = new HashSet<string>(graph.States.Select(s => s.StateId), StringComparer.Ordinal);
        var windowIds = new HashSet<string>(graph.Windows.Select(w => w.WindowId), StringComparer.Ordinal);
        var elementIds = new HashSet<string>(graph.Elements.Select(e => e.ElementId), StringComparer.Ordinal);
        var transitionIds = new HashSet<string>(StringComparer.Ordinal);

        if (graph.Windows.Count != windowIds.Count)
            result.Errors.Add("graph.windows contains duplicate windowId values.");
        if (graph.States.Count != stateIds.Count)
            result.Errors.Add("graph.states contains duplicate stateId values.");
        if (graph.Elements.Count != elementIds.Count)
            result.Errors.Add("graph.elements contains duplicate elementId values.");

        foreach (var state in graph.States)
        {
            if (!windowIds.Contains(state.WindowId))
                result.Errors.Add($"state '{state.StateId}' references unknown windowId '{state.WindowId}'.");
            if (!Hex64.IsMatch(state.SemanticFingerprint ?? string.Empty))
                result.Errors.Add($"state '{state.StateId}' has invalid semanticFingerprint.");
        }

        foreach (var element in graph.Elements)
        {
            if (string.IsNullOrWhiteSpace(element.Identity.ControlType))
                result.Errors.Add($"element '{element.ElementId}' is missing controlType.");
            if (string.IsNullOrWhiteSpace(element.Identity.AutomationId)
                && string.IsNullOrWhiteSpace(element.Identity.ClassName)
                && string.IsNullOrWhiteSpace(element.Identity.Name))
            {
                result.Errors.Add($"element '{element.ElementId}' has no usable identity attributes.");
            }
        }

        foreach (var transition in graph.Transitions)
        {
            if (!transitionIds.Add(transition.TransitionId))
                result.Errors.Add($"duplicate transitionId '{transition.TransitionId}'.");
            if (!stateIds.Contains(transition.SourceStateId))
                result.Errors.Add($"transition '{transition.TransitionId}' references unknown sourceStateId.");
            if (transition.TargetStateId is not null && !stateIds.Contains(transition.TargetStateId))
                result.Errors.Add($"transition '{transition.TransitionId}' references unknown targetStateId.");
            if (transition.ElementId is not null && !elementIds.Contains(transition.ElementId))
                result.Errors.Add($"transition '{transition.TransitionId}' references unknown elementId.");
            if (transition.ObservationCount < 1)
                result.Errors.Add($"transition '{transition.TransitionId}' observationCount must be >= 1.");
            if (transition.Evidence.Count == 0)
                result.Errors.Add($"transition '{transition.TransitionId}' must include evidence.");
        }

        foreach (var entry in graph.EntryStateIds)
        {
            if (!stateIds.Contains(entry))
                result.Errors.Add($"entryStateId '{entry}' is unknown.");
        }

        foreach (var terminal in graph.TerminalStateIds)
        {
            if (!stateIds.Contains(terminal))
                result.Errors.Add($"terminalStateId '{terminal}' is unknown.");
        }
    }

    private static void ValidateWorkflows(
        List<ApplicationWorkflow> workflows,
        ApplicationStateGraph graph,
        KnowledgeValidationResult result)
    {
        var stateIds = new HashSet<string>(graph.States.Select(s => s.StateId), StringComparer.Ordinal);
        var transitionIds = new HashSet<string>(graph.Transitions.Select(t => t.TransitionId), StringComparer.Ordinal);
        var seen = new HashSet<string>(StringComparer.Ordinal);

        foreach (var workflow in workflows)
        {
            if (!seen.Add(workflow.WorkflowId))
                result.Errors.Add($"duplicate workflowId '{workflow.WorkflowId}'.");
            if (workflow.Steps.Count == 0)
            {
                result.Errors.Add($"workflow '{workflow.WorkflowId}' has no steps.");
                continue;
            }

            if (!stateIds.Contains(workflow.EntryStateId))
                result.Errors.Add($"workflow '{workflow.WorkflowId}' has unknown entryStateId.");
            if (workflow.TerminalStateId is not null && !stateIds.Contains(workflow.TerminalStateId))
                result.Errors.Add($"workflow '{workflow.WorkflowId}' has unknown terminalStateId.");

            if (!string.Equals(workflow.EntryStateId, workflow.Steps[0].SourceStateId, StringComparison.Ordinal))
            {
                result.Errors.Add(
                    $"workflow '{workflow.WorkflowId}' entryStateId does not match first step sourceStateId.");
            }

            var last = workflow.Steps[^1];
            var expectedTerminal = last.TargetStateId ?? last.SourceStateId;
            if (workflow.TerminalStateId is not null
                && !string.Equals(workflow.TerminalStateId, expectedTerminal, StringComparison.Ordinal))
            {
                result.Errors.Add(
                    $"workflow '{workflow.WorkflowId}' terminalStateId does not match final step target.");
            }

            string? previousTarget = null;
            for (var i = 0; i < workflow.Steps.Count; i++)
            {
                var step = workflow.Steps[i];
                if (step.Ordinal != i)
                    result.Errors.Add($"workflow '{workflow.WorkflowId}' step ordinal {step.Ordinal} expected {i}.");

                if (!transitionIds.Contains(step.TransitionId))
                {
                    result.Warnings.Add(
                        $"workflow '{workflow.WorkflowId}' step references transition '{step.TransitionId}' not present in graph (may be filtered).");
                }

                if (!stateIds.Contains(step.SourceStateId))
                    result.Errors.Add($"workflow '{workflow.WorkflowId}' step {i} has unknown sourceStateId.");
                if (step.TargetStateId is not null && !stateIds.Contains(step.TargetStateId))
                    result.Errors.Add($"workflow '{workflow.WorkflowId}' step {i} has unknown targetStateId.");

                if (previousTarget is not null
                    && !string.Equals(previousTarget, step.SourceStateId, StringComparison.Ordinal))
                {
                    result.Errors.Add(
                        $"workflow '{workflow.WorkflowId}' is disconnected at step {i}: "
                        + $"previous target '{previousTarget}' != source '{step.SourceStateId}'.");
                }

                previousTarget = step.TargetStateId ?? step.SourceStateId;
            }
        }
    }

    private static void ValidateCoverage(KnowledgeCoverageReport coverage, KnowledgeValidationResult result)
    {
        if (coverage.Windows < 0 || coverage.UniqueStates < 0 || coverage.Conflicts < 0)
            result.Errors.Add("coverage counts must be non-negative.");
    }

    private static void ValidateApproval(KnowledgeApproval approval, KnowledgeValidationResult result)
    {
        if (approval.ApprovalSchemaVersion != KnowledgeConstants.ApprovalSchemaVersion)
            result.Errors.Add($"Unsupported approvalSchemaVersion '{approval.ApprovalSchemaVersion}'.");
        if (string.IsNullOrWhiteSpace(approval.Reviewer))
            result.Errors.Add("approval.reviewer is required.");
        if (!Hex64.IsMatch(approval.SourceKnowledgeHash ?? string.Empty))
            result.Errors.Add("approval.sourceKnowledgeHash must be a 64-char lowercase hex digest.");
        if (approval.Decision is not ("approve" or "reject"))
            result.Errors.Add("approval.decision must be approve or reject.");
        if (approval.PromotesPageObjects)
            result.Errors.Add("approval.promotesPageObjects must be false.");
    }
}
