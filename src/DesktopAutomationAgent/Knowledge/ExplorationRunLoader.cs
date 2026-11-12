using System.Text.Json;
using DesktopAutomationAgent.Exploration;
using DesktopAutomationAgent.Workspace;

namespace DesktopAutomationAgent.Knowledge;

public sealed class ExplorationRunLoader
{
    private readonly IWorkspaceManager _workspace;
    private readonly int _maxFileBytes;

    public ExplorationRunLoader(IWorkspaceManager workspace, int maxFileBytes = KnowledgeConstants.MaxFileBytesDefault)
    {
        _workspace = workspace;
        _maxFileBytes = maxFileBytes;
    }

    public LoadedExplorationRun Load(string explorationRun)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(explorationRun);
        _workspace.EnsureInitialized();

        var runDirectory = ResolveRunDirectory(explorationRun);
        KnowledgePathSafety.EnsureReadablePathInside(runDirectory, _workspace.RootPath);

        var manifestPath = Path.Combine(runDirectory, "manifest.json");
        var summaryPath = Path.Combine(runDirectory, "summary.json");
        var statesPath = Path.Combine(runDirectory, "states.jsonl");
        var transitionsPath = Path.Combine(runDirectory, "transitions.jsonl");

        EnsureFile(manifestPath);
        EnsureFile(summaryPath);
        EnsureFile(statesPath);
        EnsureFile(transitionsPath);

        var manifest = ReadJsonFile<ExplorationManifest>(manifestPath);
        var summary = ReadJsonFile<ExplorationSummary>(summaryPath);
        var states = ReadJsonlFile<ExplorationStateRecord>(statesPath);
        var transitions = ReadJsonlFile<ExplorationTransitionRecord>(transitionsPath);

        if (string.IsNullOrWhiteSpace(manifest.ExplorationId))
            throw new InvalidOperationException($"Exploration manifest at '{manifestPath}' is missing explorationId.");

        ExplorationPolicyDocument? policy = null;
        if (!string.IsNullOrWhiteSpace(manifest.PolicyPath) && File.Exists(manifest.PolicyPath))
        {
            // Policy may live outside the workspace (acceptance fixtures). Read for identity only;
            // do not require containment when the path is absolute and the file exists.
            try
            {
                EnsureFileSize(manifest.PolicyPath);
                policy = JsonSerializer.Deserialize<ExplorationPolicyDocument>(
                    File.ReadAllText(manifest.PolicyPath),
                    KnowledgeJson.Compact);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException)
            {
                policy = null;
            }
        }

        var policySha = string.IsNullOrWhiteSpace(manifest.PolicySha256)
            ? KnowledgeHash.Sha256Hex(string.Empty)
            : manifest.PolicySha256.Trim().ToLowerInvariant();

        var artifactHash = KnowledgeHash.Sha256Hex(
            KnowledgeHash.Sha256File(manifestPath)
            + KnowledgeHash.Sha256File(statesPath)
            + KnowledgeHash.Sha256File(transitionsPath)
            + KnowledgeHash.Sha256File(summaryPath));

        return new LoadedExplorationRun
        {
            ExplorationId = manifest.ExplorationId,
            RunDirectory = runDirectory,
            ArtifactHash = artifactHash,
            PolicySha256 = policySha,
            Manifest = manifest,
            Summary = summary,
            States = states,
            Transitions = transitions,
            Policy = policy
        };
    }

    private string ResolveRunDirectory(string explorationRun)
    {
        var trimmed = explorationRun.Trim();
        if (Directory.Exists(trimmed))
            return Path.GetFullPath(trimmed);

        var underRuns = Path.Combine(_workspace.RootPath, "exploration", "runs", trimmed);
        if (Directory.Exists(underRuns))
            return Path.GetFullPath(underRuns);

        throw new FileNotFoundException(
            $"Exploration run '{explorationRun}' was not found as a directory or under exploration/runs/.");
    }

    private void EnsureFile(string path)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException($"Required exploration artifact is missing: {path}");
        EnsureFileSize(path);
        KnowledgePathSafety.EnsureNotSymlinkEscape(path, _workspace.RootPath);
    }

    private void EnsureFileSize(string path)
    {
        var info = new FileInfo(path);
        if (info.Length > _maxFileBytes)
        {
            throw new InvalidOperationException(
                $"Exploration artifact '{path}' exceeds maximum size of {_maxFileBytes} bytes.");
        }
    }

    private T ReadJsonFile<T>(string path)
    {
        EnsureFileSize(path);
        var text = File.ReadAllText(path);
        return JsonSerializer.Deserialize<T>(text, KnowledgeJson.Compact)
               ?? throw new InvalidOperationException($"Failed to deserialize '{path}'.");
    }

    private List<T> ReadJsonlFile<T>(string path)
    {
        EnsureFileSize(path);
        var lines = File.ReadAllLines(path);
        if (lines.Length > KnowledgeConstants.MaxTransitions)
        {
            throw new InvalidOperationException(
                $"Exploration artifact '{path}' exceeds maximum line count of {KnowledgeConstants.MaxTransitions}.");
        }

        var items = new List<T>(lines.Length);
        foreach (var line in lines)
        {
            if (string.IsNullOrWhiteSpace(line))
                continue;

            var item = JsonSerializer.Deserialize<T>(line, KnowledgeJson.Compact)
                       ?? throw new InvalidOperationException($"Failed to deserialize a line in '{path}'.");
            items.Add(item);
        }

        return items;
    }
}
