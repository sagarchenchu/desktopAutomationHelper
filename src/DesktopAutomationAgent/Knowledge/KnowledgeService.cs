using System.Text.Json;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.ObjectRepository;
using DesktopAutomationAgent.Workspace;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Knowledge;

/// <summary>
/// Offline knowledge orchestration. Never contacts the driver or network.
/// </summary>
public sealed class KnowledgeService
{
    private readonly IWorkspaceManager _workspace;
    private readonly KnowledgeOptions _options;
    private readonly ExplorationRunLoader _loader;
    private readonly KnowledgeGraphBuilder _builder = new();
    private readonly KnowledgeArtifactWriter _writer = new();
    private readonly KnowledgeValidator _validator;
    private readonly KnowledgeComparer _comparer = new();
    private readonly KnowledgeSearcher _searcher = new();
    private readonly KnowledgeApprover _approver;
    private readonly KnowledgePageObjectSuggester _pageObjectSuggester;

    public KnowledgeService(
        IWorkspaceManager workspace,
        ObjectArtifactWriter objectArtifactWriter,
        IOptions<AgentOptions> options)
    {
        _workspace = workspace;
        _options = options.Value.Knowledge;
        _loader = new ExplorationRunLoader(workspace, _options.MaxFileBytes);
        var schemaDir = Path.Combine(workspace.RootPath, "schemas");
        _validator = new KnowledgeValidator(Directory.Exists(schemaDir) ? schemaDir : null);
        _approver = new KnowledgeApprover(workspace, _writer, _validator, _options.MaxFileBytes);
        _pageObjectSuggester = new KnowledgePageObjectSuggester(workspace, objectArtifactWriter);
    }

    public KnowledgeBuildResult Build(
        string applicationId,
        IReadOnlyList<string> explorationRuns,
        bool suggestPageObjects,
        string? outputDirectory = null,
        DateTimeOffset? createdAtUtc = null)
    {
        try
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(applicationId);
            ArgumentNullException.ThrowIfNull(explorationRuns);

            if (explorationRuns.Count == 0)
                return FailBuild(2, "build-knowledge requires at least one --exploration-run.");

            if (explorationRuns.Count > _options.MaxExplorationRuns)
            {
                return FailBuild(
                    2,
                    $"At most {_options.MaxExplorationRuns} --exploration-run values are allowed.");
            }

            _workspace.EnsureInitialized();

            var loaded = explorationRuns
                .Select(r => _loader.Load(r))
                .OrderBy(r => r.ExplorationId, StringComparer.Ordinal)
                .ToList();

            // Order-independence check foundation: builder re-sorts internally.
            var knowledge = _builder.Build(applicationId, loaded, createdAtUtc);

            var validation = _validator.Validate(knowledge);
            if (!validation.IsValid)
                return FailBuild(5, "Built knowledge failed validation: " + string.Join("; ", validation.Errors));

            var appSeg = KnowledgePathSafety.SanitizePathSegment(knowledge.ApplicationId, "applicationId");
            var fpSeg = KnowledgePathSafety.SanitizePathSegment(knowledge.BuildFingerprint, "buildFingerprint");
            var runId = knowledge.KnowledgeId;

            var output = string.IsNullOrWhiteSpace(outputDirectory)
                ? Path.Combine(_workspace.RootPath, "knowledge", "candidates", appSeg, fpSeg, runId)
                : Path.GetFullPath(outputDirectory);

            // Candidate outputs must stay under knowledge/candidates (no approved/pages bypass).
            KnowledgePathSafety.EnsureUnderCandidates(output, _workspace.RootPath);
            Directory.CreateDirectory(output);
            _writer.WriteCandidate(output, knowledge, overwrite: false);

            var suggested = new List<string>();
            if (suggestPageObjects)
                suggested.AddRange(_pageObjectSuggester.Suggest(knowledge));

            return new KnowledgeBuildResult
            {
                Success = true,
                ExitCode = 0,
                OutputDirectory = output,
                Knowledge = knowledge,
                SuggestedPageObjectPaths = suggested,
                ContentHash = KnowledgeArtifactWriter.ComputeContentHash(knowledge)
            };
        }
        catch (KnowledgePathException ex)
        {
            return FailBuild(5, ex.Message);
        }
        catch (Exception ex) when (ex is InvalidOperationException or FileNotFoundException or IOException or JsonException or UnauthorizedAccessException or ArgumentException)
        {
            var code = ex is ArgumentException ? 2 : 5;
            return FailBuild(code, ex.Message);
        }
    }

    public KnowledgeValidationResult ValidatePath(string knowledgePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(knowledgePath);
        _workspace.EnsureInitialized();
        var full = Path.GetFullPath(knowledgePath);
        if (File.Exists(full) && full.EndsWith("application-knowledge.json", StringComparison.OrdinalIgnoreCase))
            full = Path.GetDirectoryName(full)!;

        KnowledgePathSafety.EnsureReadablePathInside(full, _workspace.RootPath);
        return _validator.ValidateDirectory(full);
    }

    public KnowledgeCompareResult Compare(string leftPath, string rightPath)
    {
        var left = LoadDocument(leftPath);
        var right = LoadDocument(rightPath);
        return _comparer.Compare(left, right);
    }

    public IReadOnlyList<KnowledgeSearchHit> Search(string knowledgePath, string query)
    {
        var document = LoadDocument(knowledgePath);
        return _searcher.Search(document, query);
    }

    public KnowledgeApproveResult Approve(string sourceDirectory, string reviewer, string reviewedInputPath) =>
        _approver.Approve(sourceDirectory, reviewer, reviewedInputPath);

    public ApplicationKnowledgeDocument LoadDocument(string knowledgePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(knowledgePath);
        _workspace.EnsureInitialized();

        var full = Path.GetFullPath(knowledgePath);
        if (Directory.Exists(full))
            full = Path.Combine(full, "application-knowledge.json");

        if (!File.Exists(full))
            throw new FileNotFoundException($"Knowledge document not found: {full}");

        var root = Path.GetDirectoryName(full)!;
        KnowledgePathSafety.EnsureReadablePathInside(root, _workspace.RootPath);
        KnowledgePathSafety.EnsureNotSymlinkEscape(full, _workspace.RootPath);

        var info = new FileInfo(full);
        if (info.Length > _options.MaxFileBytes)
            throw new InvalidOperationException($"Knowledge file exceeds maximum size of {_options.MaxFileBytes} bytes.");

        return JsonSerializer.Deserialize<ApplicationKnowledgeDocument>(File.ReadAllText(full), KnowledgeJson.Compact)
               ?? throw new InvalidOperationException($"Failed to deserialize '{full}'.");
    }

    private static KnowledgeBuildResult FailBuild(int exitCode, string error) =>
        new()
        {
            Success = false,
            ExitCode = exitCode,
            Error = error
        };
}
