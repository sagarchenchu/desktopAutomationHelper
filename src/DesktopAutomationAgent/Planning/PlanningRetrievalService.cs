using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using DesktopAutomationAgent.Compilation;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver.Models;
using DesktopAutomationAgent.Knowledge;
using DesktopAutomationAgent.ObjectRepository;
using DesktopAutomationAgent.Workspace;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Planning;

/// <summary>
/// Deterministic lexical + filter retrieval over approved knowledge, page objects,
/// step registry, and operation catalog. Embeddings are optional and never SoT.
/// Never sends entire UIA trees, secrets, editable values, or unrelated corpora.
/// </summary>
public sealed class PlanningRetrievalService
{
    private readonly IWorkspaceManager _workspace;
    private readonly PlanningOptions _options;
    private readonly ObjectRepositoryReader _objectRepositoryReader;
    private readonly KnowledgeSearcher _searcher = new();
    private readonly StepRegistryValidator _registryValidator = new();

    public PlanningRetrievalService(
        IWorkspaceManager workspace,
        ObjectRepositoryReader objectRepositoryReader,
        IOptions<AgentOptions> options)
    {
        _workspace = workspace;
        _objectRepositoryReader = objectRepositoryReader;
        _options = options.Value.Planning;
    }

    public RetrievalManifest Retrieve(PlanningRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        _workspace.EnsureInitialized();

        var canonicalGoal = PlanningPromptGuard.CanonicalizeGoal(request.Goal);
        var manifest = new RetrievalManifest
        {
            CanonicalGoal = canonicalGoal,
            GoalHash = Sha256Hex(canonicalGoal),
            ApplicationId = request.ApplicationId,
            EmbeddingsUsed = false,
            Exclusions =
            [
                "raw-uia-trees",
                "secrets",
                "editable-values",
                "customer-data",
                "unrelated-repos",
                "provider-credentials"
            ]
        };

        var limit = Math.Clamp(_options.MaxRetrievalHits, 1, 200);
        var hits = new List<RetrievalHitRecord>();

        // Knowledge (approved only)
        if (!string.IsNullOrWhiteSpace(request.KnowledgePath))
        {
            var knowledgeDir = _workspace.ResolveSafePath(request.KnowledgePath);
            KnowledgePathSafety.EnsureUnderApproved(knowledgeDir, _workspace.RootPath);
            PlanningPathSafety.EnsureReadablePathInside(knowledgeDir, _workspace.RootPath);
            var knowledgeFile = Path.Combine(knowledgeDir, "application-knowledge.json");
            if (!File.Exists(knowledgeFile))
                throw new InvalidOperationException($"Approved knowledge file missing at '{knowledgeFile}'.");

            PlanningPathSafety.EnsureReadablePathInside(knowledgeFile, _workspace.RootPath);
            var knowledgeJson = File.ReadAllText(knowledgeFile);
            var knowledge = JsonSerializer.Deserialize<ApplicationKnowledgeDocument>(knowledgeJson, KnowledgeJson.Compact)
                ?? throw new InvalidOperationException("Failed to deserialize approved knowledge.");

            if (!string.Equals(knowledge.Status, "approved", StringComparison.Ordinal))
                throw new InvalidOperationException("Planning retrieval requires status=approved knowledge.");

            if (!string.IsNullOrWhiteSpace(request.ExpectedBuildFingerprint)
                && !string.Equals(knowledge.BuildFingerprint, request.ExpectedBuildFingerprint, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    $"stale-build: knowledge buildFingerprint '{knowledge.BuildFingerprint}' does not match expected '{request.ExpectedBuildFingerprint}'.");
            }

            if (!string.IsNullOrWhiteSpace(request.ApplicationId)
                && !string.Equals(knowledge.ApplicationId, request.ApplicationId, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    $"applicationId mismatch: knowledge has '{knowledge.ApplicationId}', request has '{request.ApplicationId}'.");
            }

            manifest.BuildFingerprint = knowledge.BuildFingerprint;
            manifest.ApplicationId ??= knowledge.ApplicationId;
            manifest.KnowledgeContentHash = Sha256Hex(knowledgeJson);

            foreach (var hit in _searcher.Search(knowledge, canonicalGoal, limit))
            {
                hits.Add(new RetrievalHitRecord
                {
                    Corpus = "knowledge",
                    Kind = hit.Kind,
                    Id = hit.Id,
                    Score = hit.Score,
                    ContentHash = Sha256Hex($"{hit.Kind}:{hit.Id}:{hit.Text}"),
                    Snippet = Truncate(hit.Text, 200)
                });
            }

            // Always include a small stable slice of windows for multi-window goals.
            foreach (var window in knowledge.Graph.Windows.OrderBy(w => w.WindowId, StringComparer.Ordinal).Take(4))
            {
                if (hits.Any(h => h.Corpus == "knowledge" && h.Id == window.WindowId))
                    continue;
                hits.Add(new RetrievalHitRecord
                {
                    Corpus = "knowledge",
                    Kind = "window",
                    Id = window.WindowId,
                    Score = 1,
                    ContentHash = Sha256Hex($"window:{window.WindowId}:{window.Title}"),
                    Snippet = Truncate($"{window.WindowId} {window.Title}", 200)
                });
            }
        }

        // Object repository (approved pages only — reader already scopes to pages/)
        var repoPath = string.IsNullOrWhiteSpace(request.ObjectRepositoryPath)
            ? "object-repository/repository.json"
            : request.ObjectRepositoryPath;
        var repoFull = _workspace.ResolveSafePath(repoPath);
        if (File.Exists(repoFull))
        {
            var repoBytes = File.ReadAllBytes(repoFull);
            manifest.ObjectRepositoryHash = Sha256Hex(repoBytes);
            var repoResult = _objectRepositoryReader.Read(repoPath);
            if (!repoResult.IsValid || repoResult.Snapshot is null)
                throw new InvalidOperationException("Object repository invalid: " + string.Join("; ", repoResult.Errors));
            var snapshot = repoResult.Snapshot;
            var terms = Tokenize(canonicalGoal);
            foreach (var (pageId, page) in snapshot.Pages.OrderBy(p => p.Key, StringComparer.Ordinal))
            {
                if (page.Elements is null)
                    continue;
                foreach (var (elementId, element) in page.Elements.OrderBy(e => e.Key, StringComparer.Ordinal))
                {
                    var objectRef = $"{pageId}.{elementId}";
                    var locatorBits = element.Locator is null
                        ? string.Empty
                        : $"{element.Locator.AutomationId} {element.Locator.Name} {element.Locator.ControlType}";
                    var text = $"{objectRef} {page.Name} {element.Description} {locatorBits}".ToLowerInvariant();
                    var score = Score(text, terms);
                    if (score <= 0)
                        continue;
                    hits.Add(new RetrievalHitRecord
                    {
                        Corpus = "object-repository",
                        Kind = "objectRef",
                        Id = objectRef,
                        Score = score,
                        ContentHash = Sha256Hex(objectRef),
                        Snippet = Truncate(text, 200)
                    });
                }
            }
        }

        // Step registry
        var registryPath = string.IsNullOrWhiteSpace(request.StepRegistryPath)
            ? "step-registry/registry.json"
            : request.StepRegistryPath;
        var registryFull = _workspace.ResolveSafePath(registryPath);
        if (File.Exists(registryFull))
        {
            PlanningPathSafety.EnsureReadablePathInside(registryFull, _workspace.RootPath);
            var registryJson = File.ReadAllText(registryFull);
            manifest.StepRegistryHash = Sha256Hex(registryJson);
            var validation = _registryValidator.ValidateFile(registryFull, registryPath);
            if (!validation.IsValid || validation.Document is null)
                throw new InvalidOperationException("Step registry invalid: " + string.Join("; ", validation.Errors));

            var registry = validation.Document;
            var terms = Tokenize(canonicalGoal);
            foreach (var mapping in registry.Mappings.OrderBy(m => m.MappingId, StringComparer.Ordinal))
            {
                var text = $"{mapping.MappingId} {mapping.Pattern.Value} {mapping.Risk}".ToLowerInvariant();
                var score = Score(text, terms);
                if (score <= 0 && !string.Equals(mapping.MappingId, "list-visible-windows", StringComparison.Ordinal))
                    continue;
                if (score <= 0)
                    score = 1;
                hits.Add(new RetrievalHitRecord
                {
                    Corpus = "step-registry",
                    Kind = "mapping",
                    Id = mapping.MappingId,
                    Score = score,
                    ContentHash = Sha256Hex($"{mapping.MappingId}:{mapping.Version}"),
                    Snippet = Truncate(text, 200)
                });
            }
        }

        // Operation catalog (offline file or minimal built-in safe ops)
        OperationsCatalogDto catalog;
        if (!string.IsNullOrWhiteSpace(request.CatalogPath))
        {
            var catalogFull = _workspace.ResolveSafePath(request.CatalogPath);
            PlanningPathSafety.EnsureReadablePathInside(catalogFull, _workspace.RootPath);
            var catalogJson = File.ReadAllText(catalogFull);
            catalog = JsonSerializer.Deserialize<OperationsCatalogDto>(catalogJson, PlanningJson.Compact)
                ?? throw new InvalidOperationException("Failed to deserialize operations catalog.");
            manifest.CatalogFingerprint = CompilationFingerprint.ComputeCatalogFingerprint(catalog);
        }
        else
        {
            catalog = MinimalSafeCatalog();
            manifest.CatalogFingerprint = CompilationFingerprint.ComputeCatalogFingerprint(catalog);
        }

        var opTerms = Tokenize(canonicalGoal);
        foreach (var op in catalog.Operations.OrderBy(o => o.Name, StringComparer.Ordinal))
        {
            var text = $"{op.Name} {op.Category} {op.OperationType}".ToLowerInvariant();
            var score = Score(text, opTerms);
            if (score <= 0 && !string.Equals(op.Name, "listwindows", StringComparison.OrdinalIgnoreCase))
                continue;
            if (score <= 0)
                score = 1;
            hits.Add(new RetrievalHitRecord
            {
                Corpus = "operation-catalog",
                Kind = "operation",
                Id = op.Name,
                Score = score,
                ContentHash = Sha256Hex(op.Name),
                Snippet = Truncate(text, 200)
            });
        }

        manifest.Hits = hits
            .OrderByDescending(h => h.Score)
            .ThenBy(h => h.Corpus, StringComparer.Ordinal)
            .ThenBy(h => h.Id, StringComparer.Ordinal)
            .Take(limit)
            .ToList();

        // Optional embeddings explicitly never used as SoT (flag stays false unless enabled and implemented).
        if (_options.AllowOptionalEmbeddings)
            manifest.Exclusions.Add("embeddings-not-source-of-truth");

        return manifest;
    }

    public static OperationsCatalogDto MinimalSafeCatalog() =>
        new()
        {
            SchemaVersion = 2,
            DriverVersion = "test",
            Operations =
            [
                new OperationDescriptorDto
                {
                    Name = "listwindows",
                    Category = "query",
                    OperationType = "read",
                    RequiresSession = false,
                    RequiredInputs = []
                },
                new OperationDescriptorDto
                {
                    Name = "clickuia",
                    Category = "action",
                    OperationType = "invoke",
                    RequiresSession = false,
                    RequiredInputs = ["locator"]
                },
                new OperationDescriptorDto
                {
                    Name = "typeuia",
                    Category = "action",
                    OperationType = "input",
                    RequiresSession = false,
                    RequiredInputs = ["locator", "value"]
                },
                new OperationDescriptorDto
                {
                    Name = "finduia",
                    Category = "query",
                    OperationType = "read",
                    RequiresSession = false,
                    RequiredInputs = ["locator"]
                },
                new OperationDescriptorDto
                {
                    Name = "dumpuia",
                    Category = "query",
                    OperationType = "read",
                    RequiresSession = false,
                    RequiredInputs = []
                }
            ]
        };

    private static string[] Tokenize(string query) =>
        query.Split([' ', '\t', '\r', '\n', ',', ';', '.', ':', '/', '\\'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(t => t.ToLowerInvariant())
            .Where(t => t.Length > 1)
            .Distinct(StringComparer.Ordinal)
            .OrderBy(t => t, StringComparer.Ordinal)
            .ToArray();

    private static int Score(string haystack, string[] terms)
    {
        var score = 0;
        foreach (var term in terms)
        {
            if (haystack.Contains(term, StringComparison.Ordinal))
                score += 1 + (haystack.StartsWith(term, StringComparison.Ordinal) ? 1 : 0);
        }

        return score;
    }

    private static string Truncate(string text, int max) =>
        text.Length <= max ? text : text[..max];

    private static string Sha256Hex(string text) =>
        Sha256Hex(Encoding.UTF8.GetBytes(text));

    private static string Sha256Hex(byte[] bytes) =>
        Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
}
