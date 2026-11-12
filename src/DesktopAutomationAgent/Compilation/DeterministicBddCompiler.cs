using System.Text.Json;
using DesktopAutomationAgent.Driver.Models;
using DesktopAutomationAgent.Knowledge;
using DesktopAutomationAgent.ObjectRepository;
using DesktopAutomationAgent.Plans;

namespace DesktopAutomationAgent.Compilation;

public sealed class DeterministicBddCompiler
{
    private readonly GherkinBddParser _parser = new();
    private readonly PlanValidator _planValidator = new();

    public CompilationResult Compile(DeterministicCompileRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        var explain = request.Explain ? new List<CompilationExplainEntry>() : null;
        var errors = new List<string>();
        var warnings = new List<string>();

        if (request.TestCase is null || request.Registry is null)
        {
            return Fail(5, ["test case and step registry are required."], explain);
        }

        var parse = _parser.ParseTestCase(request.TestCase);
        if (!parse.IsSuccess || parse.Document is null)
        {
            errors.AddRange(parse.Errors);
            return Fail(5, errors, explain);
        }

        var matcher = new StepMatcher(request.Registry);
        var planSteps = new List<PlanStep>();
        var traces = new List<CompilationStepTrace>();
        var usedKnowledge = false;
        var knowledgeErrors = new List<string>();

        // Stable scenario order already applied in IR.
        foreach (var scenario in parse.Document.Scenarios)
        {
            foreach (var step in scenario.Steps)
            {
                var matchResult = matcher.Match(step);
                if (!matchResult.IsSuccess || matchResult.Match is null)
                {
                    errors.AddRange(matchResult.Errors);
                    explain?.Add(new CompilationExplainEntry
                    {
                        Ordinal = step.Ordinal,
                        Keyword = step.Keyword,
                        Text = step.Text,
                        Outcome = "failed",
                        Details = matchResult.Errors
                    });
                    continue;
                }

                var match = matchResult.Match;
                var binding = BuildBindings(match, request.ParameterBindings);

                // Resolve declared approved object refs (static + from params).
                foreach (var objectRefTemplate in match.Mapping.ApprovedObjectRefs)
                {
                    var objectRef = ApplyTemplate(objectRefTemplate, binding);
                    if (!TryResolveObjectRef(objectRef, request.ObjectRepository, out var objectError))
                    {
                        errors.Add(
                            $"object-resolution: step '{step.Keyword} {step.Text}' mapping '{match.Mapping.MappingId}': {objectError}");
                    }
                }

                // Parameter values typed as objectRef must also resolve.
                foreach (var parameter in match.Mapping.Parameters.Where(p => p.Type == "objectRef"))
                {
                    if (!binding.TryGetValue(parameter.Name, out var objectRef))
                        continue;
                    if (!TryResolveObjectRef(objectRef, request.ObjectRepository, out var objectError))
                    {
                        errors.Add(
                            $"object-resolution: step '{step.Keyword} {step.Text}' parameter '{parameter.Name}': {objectError}");
                    }
                }

                // $objectRef inside templates
                foreach (var template in match.Mapping.PlanSteps)
                {
                    CollectObjectRefsFromTemplate(template.Arguments, binding, out var embeddedRefs);
                    foreach (var objectRef in embeddedRefs)
                    {
                        if (!TryResolveObjectRef(objectRef, request.ObjectRepository, out var objectError))
                        {
                            errors.Add(
                                $"object-resolution: step '{step.Keyword} {step.Text}' template objectRef '{objectRef}': {objectError}");
                        }
                    }
                }

                if (match.Mapping.ApprovedKnowledgeRefs.Count > 0)
                {
                    usedKnowledge = true;
                    if (request.ApprovedKnowledge is null)
                    {
                        knowledgeErrors.Add(
                            $"knowledge-missing: mapping '{match.Mapping.MappingId}' requires approved knowledge but none was provided.");
                    }
                    else
                    {
                        foreach (var knowledgeRef in match.Mapping.ApprovedKnowledgeRefs)
                        {
                            var role = string.IsNullOrWhiteSpace(knowledgeRef.Role)
                                ? "navigation"
                                : knowledgeRef.Role.Trim();
                            if (role.Equals("navigation", StringComparison.OrdinalIgnoreCase))
                            {
                                knowledgeErrors.Add(
                                    $"unsupported-feature: knowledge role 'navigation' is not implemented " +
                                    $"(mapping '{match.Mapping.MappingId}'); refusing listwindows/navigation stubs. " +
                                    "Use expectedState/postcondition until workflow/transition expansion ships.");
                                continue;
                            }

                            var id = ApplyTemplate(knowledgeRef.Id, binding);
                            if (!TryResolveKnowledge(request.ApprovedKnowledge, knowledgeRef.Kind, id, out var knowledgeError))
                            {
                                knowledgeErrors.Add(
                                    $"knowledge-resolution: mapping '{match.Mapping.MappingId}' {knowledgeRef.Kind} '{id}': {knowledgeError}");
                            }
                        }
                    }
                }

                var emittedIds = new List<string>();
                foreach (var template in match.Mapping.PlanSteps)
                {
                    var planStep = MaterializeStep(template, binding, planSteps.Count, out var materializeErrors);
                    if (materializeErrors.Count > 0)
                    {
                        errors.AddRange(materializeErrors.Select(e =>
                            $"step '{step.Keyword} {step.Text}' mapping '{match.Mapping.MappingId}': {e}"));
                        continue;
                    }

                    planSteps.Add(planStep);
                    emittedIds.Add(planStep.Id);
                }

                traces.Add(new CompilationStepTrace
                {
                    Ordinal = step.Ordinal,
                    Keyword = step.Keyword,
                    Text = step.Text,
                    MappingId = match.Mapping.MappingId,
                    Parameters = new Dictionary<string, string>(binding, StringComparer.Ordinal),
                    PlanStepIds = emittedIds
                });

                explain?.Add(new CompilationExplainEntry
                {
                    Ordinal = step.Ordinal,
                    Keyword = step.Keyword,
                    Text = step.Text,
                    MappingId = match.Mapping.MappingId,
                    Outcome = "mapped",
                    Details = [$"emitted {emittedIds.Count} plan step(s): {string.Join(", ", emittedIds)}"]
                });
            }
        }

        errors.AddRange(knowledgeErrors);

        if (errors.Count > 0)
            return Fail(5, errors, explain);

        if (planSteps.Count == 0)
            return Fail(5, ["compilation produced zero plan steps."], explain);

        var tags = parse.Document.Tags
            .Concat(request.TestCase.Tags ?? [])
            .Distinct(StringComparer.Ordinal)
            .OrderBy(t => t, StringComparer.Ordinal)
            .ToList();

        string objectRepoPath;
        try
        {
            objectRepoPath = CompilationFingerprint.CanonicalizeWorkspaceRelativePath(
                request.ObjectRepositoryRelativePath);
        }
        catch (ArgumentException ex)
        {
            return Fail(5, [ex.Message], explain);
        }

        var plan = new PlanManifest
        {
            Schema = "../schemas/plan.schema.json",
            SchemaVersion = 1,
            CatalogSchemaVersion = 2,
            PlanId = request.TestCase.JiraKey,
            Name = request.TestCase.Summary,
            Description = $"Compiled from offline BDD export {request.TestCase.JiraKey} (revision {request.TestCase.ResolveRevisionToken()}).",
            Tags = tags,
            ObjectRepository = string.IsNullOrWhiteSpace(objectRepoPath) ? null : objectRepoPath,
            Steps = planSteps,
            Metadata = new Dictionary<string, JsonElement>
            {
                ["compilerVersion"] = JsonSerializer.SerializeToElement(BddCompilerVersions.CompilerVersion),
                ["parserVersion"] = JsonSerializer.SerializeToElement(BddCompilerVersions.ParserVersion),
                ["jiraKey"] = JsonSerializer.SerializeToElement(request.TestCase.JiraKey),
                ["jiraRevision"] = JsonSerializer.SerializeToElement(request.TestCase.ResolveRevisionToken()),
                ["featureName"] = JsonSerializer.SerializeToElement(parse.Document.FeatureName)
            }
        };

        var validation = _planValidator.Validate(plan, $"compiled/{request.TestCase.JiraKey}/plan.json");
        if (!validation.IsValid)
        {
            errors.AddRange(validation.Errors);
            return Fail(5, errors, explain);
        }

        // Object expansion (approved repository only) — fail closed on missing refs.
        if (request.ObjectRepository is not null && !string.IsNullOrWhiteSpace(request.ObjectRepositoryRelativePath))
        {
            var expander = new PlanObjectReferenceExpander(new ObjectReferenceResolver());
            // Expand on a clone so the cached plan retains $objectRef markers (Phase 2 expansion at run time).
            var clone = ClonePlan(plan);
            var expansion = expander.Expand(clone, request.ObjectRepository, $"compiled/{request.TestCase.JiraKey}/plan.json");
            if (expansion.Errors.Count > 0)
            {
                errors.AddRange(expansion.Errors);
                return Fail(5, errors, explain);
            }

            warnings.AddRange(expansion.Warnings);
        }

        if (request.Catalog is not null)
        {
            var preflight = new PlanCatalogPreflight(request.Catalog);
            var catalogErrors = preflight.Validate(plan, $"compiled/{request.TestCase.JiraKey}/plan.json");
            if (catalogErrors.Count > 0)
            {
                errors.AddRange(catalogErrors);
                return Fail(5, errors, explain);
            }
        }

        var knowledgeSha = usedKnowledge
            ? request.ApprovedKnowledgeSha256
            : null;

        var bindingsSha = CompilationFingerprint.ComputeParameterBindingsSha256(request.ParameterBindings);
        var fingerprint = CompilationFingerprint.Compute(new CompilationFingerprintInput
        {
            InputSha256 = request.InputSha256,
            JiraRevision = request.TestCase.ResolveRevisionToken(),
            CompilerVersion = BddCompilerVersions.CompilerVersion,
            ParserVersion = BddCompilerVersions.ParserVersion,
            StepRegistrySha256 = request.StepRegistrySha256,
            ObjectRepositorySha256 = request.ObjectRepositorySha256,
            ObjectRepositoryPath = objectRepoPath,
            ApprovedKnowledgeSha256 = knowledgeSha,
            CatalogFingerprint = request.CatalogFingerprint,
            ParameterBindingsSha256 = bindingsSha
        });

        var planBytes = CompilationJson.CanonicalSerializeUtf8(plan);
        // Prefer trailing newline for text tools while keeping bytes deterministic.
        if (planBytes.Length == 0 || planBytes[^1] != (byte)'\n')
        {
            var withNl = new byte[planBytes.Length + 1];
            Buffer.BlockCopy(planBytes, 0, withNl, 0, planBytes.Length);
            withNl[^1] = (byte)'\n';
            planBytes = withNl;
        }

        var planSha = CompilationFingerprint.ComputeBytesSha256(planBytes);
        var metadata = new CompilationMetadata
        {
            JiraKey = request.TestCase.JiraKey,
            Fingerprint = fingerprint,
            InputSha256 = request.InputSha256,
            JiraRevision = request.TestCase.ResolveRevisionToken(),
            StepRegistrySha256 = request.StepRegistrySha256,
            ObjectRepositorySha256 = request.ObjectRepositorySha256,
            ObjectRepositoryPath = objectRepoPath,
            ApprovedKnowledgeSha256 = knowledgeSha,
            CatalogFingerprint = request.CatalogFingerprint,
            ParameterBindingsSha256 = bindingsSha,
            PlanSha256 = planSha,
            CompiledAtUtc = request.CompiledAtUtc ?? DateTimeOffset.UtcNow,
            Steps = traces
        };

        return new CompilationResult
        {
            Success = true,
            ExitCode = 0,
            JiraKey = request.TestCase.JiraKey,
            Fingerprint = fingerprint,
            Plan = plan,
            PlanSha256 = planSha,
            Metadata = metadata,
            Errors = errors,
            Warnings = warnings,
            Explain = explain
        };
    }

    private static CompilationResult Fail(int exitCode, List<string> errors, List<CompilationExplainEntry>? explain) =>
        new()
        {
            Success = false,
            ExitCode = exitCode,
            Errors = errors,
            Explain = explain
        };

    private static Dictionary<string, string> BuildBindings(
        StepMatch match,
        IReadOnlyDictionary<string, string>? extras)
    {
        var binding = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var pair in match.Parameters)
            binding[pair.Key] = pair.Value;
        if (extras is not null)
        {
            foreach (var pair in extras.OrderBy(p => p.Key, StringComparer.Ordinal))
                binding[pair.Key] = pair.Value;
        }

        return binding;
    }

    private static string ApplyTemplate(string template, IReadOnlyDictionary<string, string> binding)
    {
        var result = template;
        foreach (var pair in binding.OrderByDescending(p => p.Key.Length).ThenBy(p => p.Key, StringComparer.Ordinal))
            result = result.Replace("{" + pair.Key + "}", pair.Value, StringComparison.Ordinal);
        return result;
    }

    private static bool TryResolveObjectRef(
        string objectRef,
        ObjectRepositorySnapshot? snapshot,
        out string error)
    {
        error = string.Empty;
        if (snapshot is null)
        {
            error = $"object repository is required to resolve '{objectRef}'.";
            return false;
        }

        var resolution = new ObjectReferenceResolver().Resolve(snapshot, objectRef);
        if (!resolution.IsResolved)
        {
            error = resolution.Errors.Count > 0
                ? string.Join("; ", resolution.Errors)
                : $"object '{objectRef}' could not be resolved.";
            return false;
        }

        return true;
    }

    private static bool TryResolveKnowledge(
        ApplicationKnowledgeDocument knowledge,
        string kind,
        string id,
        out string error)
    {
        error = string.Empty;
        if (!string.Equals(knowledge.Status, "approved", StringComparison.Ordinal))
        {
            error = "knowledge document is not approved.";
            return false;
        }

        switch (kind.ToLowerInvariant())
        {
            case "state":
                var states = knowledge.Graph.States.Where(s => s.StateId == id).ToList();
                if (states.Count == 0)
                {
                    error = "state not found in approved knowledge.";
                    return false;
                }

                if (states.Count > 1)
                {
                    error = "ambiguous state id in approved knowledge.";
                    return false;
                }

                return true;
            case "transition":
                var transitions = knowledge.Graph.Transitions.Where(t => t.TransitionId == id).ToList();
                if (transitions.Count == 0)
                {
                    error = "transition not found in approved knowledge.";
                    return false;
                }

                if (transitions.Count > 1)
                {
                    error = "ambiguous transition id in approved knowledge.";
                    return false;
                }

                // Conflicting transition groups fail closed for navigation.
                if (!string.IsNullOrWhiteSpace(transitions[0].ConflictGroupId))
                {
                    var siblings = knowledge.Graph.Transitions
                        .Count(t => t.ConflictGroupId == transitions[0].ConflictGroupId);
                    if (siblings > 1)
                    {
                        error =
                            $"transition '{id}' belongs to conflict group '{transitions[0].ConflictGroupId}' with {siblings} members.";
                        return false;
                    }
                }

                return true;
            case "workflow":
                var workflows = knowledge.Workflows.Where(w => w.WorkflowId == id).ToList();
                if (workflows.Count != 1)
                {
                    error = workflows.Count == 0 ? "workflow not found." : "ambiguous workflow id.";
                    return false;
                }

                return true;
            case "element":
                var elements = knowledge.Graph.Elements.Where(e => e.ElementId == id).ToList();
                if (elements.Count != 1)
                {
                    error = elements.Count == 0 ? "element not found." : "ambiguous element id.";
                    return false;
                }

                return true;
            case "window":
                var windows = knowledge.Graph.Windows.Where(w => w.WindowId == id).ToList();
                if (windows.Count != 1)
                {
                    error = windows.Count == 0 ? "window not found." : "ambiguous window id.";
                    return false;
                }

                return true;
            default:
                error = $"unsupported knowledge kind '{kind}'.";
                return false;
        }
    }

    private static void CollectObjectRefsFromTemplate(
        Dictionary<string, JsonElement>? arguments,
        IReadOnlyDictionary<string, string> binding,
        out List<string> refs)
    {
        refs = [];
        if (arguments is null)
            return;

        foreach (var value in arguments.Values)
            CollectObjectRefs(value, binding, refs);
    }

    private static void CollectObjectRefs(
        JsonElement element,
        IReadOnlyDictionary<string, string> binding,
        List<string> refs)
    {
        switch (element.ValueKind)
        {
            case JsonValueKind.Object:
                if (element.TryGetProperty("$objectRef", out var objectRef)
                    && objectRef.ValueKind == JsonValueKind.String)
                {
                    refs.Add(ApplyTemplate(objectRef.GetString() ?? string.Empty, binding));
                }

                foreach (var property in element.EnumerateObject())
                    CollectObjectRefs(property.Value, binding, refs);
                break;
            case JsonValueKind.Array:
                foreach (var item in element.EnumerateArray())
                    CollectObjectRefs(item, binding, refs);
                break;
            case JsonValueKind.String:
                // no-op
                break;
        }
    }

    private static PlanStep MaterializeStep(
        PlanStepTemplate template,
        IReadOnlyDictionary<string, string> binding,
        int index,
        out List<string> errors)
    {
        errors = [];
        var id = ApplyTemplate(template.IdTemplate, binding);
        id = SanitizeStepId(id, index);

        var argsJson = template.Arguments is null
            ? "{}"
            : ApplyTemplateToJson(JsonSerializer.Serialize(template.Arguments), binding);
        Dictionary<string, JsonElement>? arguments;
        try
        {
            arguments = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(argsJson, CompilationJson.DocumentRead)
                        ?? new Dictionary<string, JsonElement>(StringComparer.Ordinal);
        }
        catch (JsonException ex)
        {
            errors.Add($"failed to materialize arguments ({ex.Message}).");
            arguments = new Dictionary<string, JsonElement>(StringComparer.Ordinal);
        }

        List<PlanAssertion>? assertions = null;
        if (template.Assertions is { Count: > 0 })
        {
            var assertionJson = ApplyTemplateToJson(JsonSerializer.Serialize(template.Assertions), binding);
            try
            {
                assertions = JsonSerializer.Deserialize<List<PlanAssertion>>(assertionJson, CompilationJson.DocumentRead);
            }
            catch (JsonException ex)
            {
                errors.Add($"failed to materialize assertions ({ex.Message}).");
            }
        }

        return new PlanStep
        {
            Id = id,
            Operation = ApplyTemplate(template.Operation, binding),
            Arguments = arguments,
            Assertions = assertions,
            Sensitive = template.Sensitive ?? false,
            CaptureResponse = template.CaptureResponse
        };
    }

    private static string ApplyTemplateToJson(string json, IReadOnlyDictionary<string, string> binding)
    {
        var result = json;
        foreach (var pair in binding.OrderByDescending(p => p.Key.Length).ThenBy(p => p.Key, StringComparer.Ordinal))
        {
            // Escape for JSON string context.
            var escaped = JsonSerializer.Serialize(pair.Value);
            // Serialize adds quotes — strip them for in-string replacement of {param}.
            escaped = escaped.Length >= 2 ? escaped[1..^1] : escaped;
            result = result.Replace("{" + pair.Key + "}", escaped, StringComparison.Ordinal);
        }

        return result;
    }

    private static string SanitizeStepId(string id, int index)
    {
        var cleaned = new string(id.Select(ch =>
            char.IsAsciiLetterOrDigit(ch) || ch is '.' or '_' or '-' ? ch : '-').ToArray());
        if (string.IsNullOrWhiteSpace(cleaned))
            cleaned = $"step-{index + 1}";
        if (cleaned.Length > 128)
            cleaned = cleaned[..128];
        if (!char.IsAsciiLetterOrDigit(cleaned[0]))
            cleaned = "s" + cleaned;
        return cleaned;
    }

    private static PlanManifest ClonePlan(PlanManifest plan)
    {
        var json = JsonSerializer.Serialize(plan, CompilationJson.Compact);
        return JsonSerializer.Deserialize<PlanManifest>(json, CompilationJson.DocumentRead)
               ?? throw new InvalidOperationException("Failed to clone plan.");
    }
}

public sealed class DeterministicCompileRequest
{
    public required JiraTestCaseDocument TestCase { get; init; }

    public required string InputSha256 { get; init; }

    public required StepRegistryDocument Registry { get; init; }

    public required string StepRegistrySha256 { get; init; }

    public ObjectRepositorySnapshot? ObjectRepository { get; init; }

    public string? ObjectRepositoryRelativePath { get; init; }

    public required string ObjectRepositorySha256 { get; init; }

    public ApplicationKnowledgeDocument? ApprovedKnowledge { get; init; }

    public string? ApprovedKnowledgeSha256 { get; init; }

    public OperationsCatalogDto? Catalog { get; init; }

    public string? CatalogFingerprint { get; init; }

    public IReadOnlyDictionary<string, string>? ParameterBindings { get; init; }

    public bool Explain { get; init; }

    public DateTimeOffset? CompiledAtUtc { get; init; }
}
