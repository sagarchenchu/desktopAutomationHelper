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
        var identityPath = Path.Combine(runDirectory, "build-identity.json");

        EnsureFile(manifestPath);
        EnsureFile(summaryPath);
        EnsureFile(statesPath);
        EnsureFile(transitionsPath);

        if (!File.Exists(identityPath))
        {
            throw new InvalidOperationException(
                $"Exploration run '{runDirectory}' is missing build-identity.json. "
                + "Knowledge merge fails closed without a captured build identity.");
        }

        EnsureFile(identityPath);

        var manifest = ReadJsonFile<ExplorationManifest>(manifestPath);
        var summary = ReadJsonFile<ExplorationSummary>(summaryPath);
        var states = ReadJsonlFile<ExplorationStateRecord>(statesPath);
        var transitions = ReadJsonlFile<ExplorationTransitionRecord>(transitionsPath);
        var buildIdentity = ReadJsonFile<ExplorationBuildIdentity>(identityPath);

        if (string.IsNullOrWhiteSpace(manifest.ExplorationId))
            throw new InvalidOperationException($"Exploration manifest at '{manifestPath}' is missing explorationId.");

        ValidateBuildIdentity(buildIdentity, manifest, identityPath);

        // Never trust a mutable on-disk policy for fingerprinting. Optionally verify that a
        // still-present policy file matches the sealed hash; mismatches fail closed.
        if (!string.IsNullOrWhiteSpace(manifest.PolicyPath) && File.Exists(manifest.PolicyPath))
        {
            EnsureFileSize(manifest.PolicyPath);
            var liveHash = KnowledgeHash.Sha256File(manifest.PolicyPath);
            if (!string.Equals(liveHash, buildIdentity.PolicySha256, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    $"Exploration policy at '{manifest.PolicyPath}' no longer matches sealed policySha256 "
                    + $"'{buildIdentity.PolicySha256}' (live '{liveHash}'). Refusing mutable policy re-read.");
            }
        }

        var artifactHash = KnowledgeHash.Sha256Hex(
            KnowledgeHash.Sha256File(manifestPath)
            + KnowledgeHash.Sha256File(statesPath)
            + KnowledgeHash.Sha256File(transitionsPath)
            + KnowledgeHash.Sha256File(summaryPath)
            + KnowledgeHash.Sha256File(identityPath));

        return new LoadedExplorationRun
        {
            ExplorationId = manifest.ExplorationId,
            RunDirectory = runDirectory,
            ArtifactHash = artifactHash,
            PolicySha256 = buildIdentity.PolicySha256,
            Manifest = manifest,
            Summary = summary,
            States = states,
            Transitions = transitions,
            BuildIdentity = buildIdentity
        };
    }

    private static void ValidateBuildIdentity(
        ExplorationBuildIdentity identity,
        ExplorationManifest manifest,
        string identityPath)
    {
        if (identity.BuildIdentitySchemaVersion != ExplorationBuildIdentity.SchemaVersion)
        {
            throw new InvalidOperationException(
                $"Unsupported buildIdentitySchemaVersion '{identity.BuildIdentitySchemaVersion}' in '{identityPath}'.");
        }

        if (string.IsNullOrWhiteSpace(identity.ProcessName))
            throw new InvalidOperationException($"build-identity.json missing required processName ({identityPath}).");

        if (string.IsNullOrWhiteSpace(identity.DriverVersion))
            throw new InvalidOperationException($"build-identity.json missing required driverVersion ({identityPath}).");

        if (string.IsNullOrWhiteSpace(identity.ExplorerVersion))
            throw new InvalidOperationException($"build-identity.json missing required explorerVersion ({identityPath}).");

        if (string.IsNullOrWhiteSpace(identity.PolicySha256) || identity.PolicySha256.Length != 64)
            throw new InvalidOperationException($"build-identity.json missing valid policySha256 ({identityPath}).");

        if (!string.IsNullOrWhiteSpace(manifest.PolicySha256)
            && !string.Equals(manifest.PolicySha256.Trim(), identity.PolicySha256, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                $"manifest.policySha256 '{manifest.PolicySha256}' does not match build-identity policySha256 "
                + $"'{identity.PolicySha256}'.");
        }

        if (identity.SnapshotSchemaVersion <= 0)
            throw new InvalidOperationException($"build-identity.json has invalid snapshotSchemaVersion ({identityPath}).");

        if (!identity.HasImmutableBuildId())
        {
            throw new InvalidOperationException(
                $"build-identity.json requires at least one immutable build id "
                + $"(executableSha256, productVersion, or fileVersion) ({identityPath}).");
        }

        if (!string.IsNullOrWhiteSpace(identity.ExecutableSha256)
            && !IsLowerHex64(identity.ExecutableSha256))
        {
            throw new InvalidOperationException(
                $"build-identity.json executableSha256 must be a 64-char lowercase hex digest ({identityPath}).");
        }
    }

    private static bool IsLowerHex64(string value) =>
        value.Length == 64
        && value.All(c => c is (>= '0' and <= '9') or (>= 'a' and <= 'f'));

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
