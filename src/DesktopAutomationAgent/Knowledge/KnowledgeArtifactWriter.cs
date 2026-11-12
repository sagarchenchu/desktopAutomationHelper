using System.Text;
using System.Text.Json;
using DesktopAutomationAgent.Driver;

namespace DesktopAutomationAgent.Knowledge;

public sealed class KnowledgeArtifactWriter
{
    public void WriteCandidate(
        string outputDirectory,
        ApplicationKnowledgeDocument knowledge,
        bool overwrite = false)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputDirectory);
        ArgumentNullException.ThrowIfNull(knowledge);

        Directory.CreateDirectory(outputDirectory);
        WriteDocument(outputDirectory, knowledge, overwrite);

        var contentHash = ComputeContentHash(knowledge);
        WriteAtomicText(
            Path.Combine(outputDirectory, "content-hash.txt"),
            contentHash + "\n",
            overwrite);

        var template = new KnowledgeApproval
        {
            ApprovalSchemaVersion = KnowledgeConstants.ApprovalSchemaVersion,
            Reviewer = "REPLACE_WITH_REVIEWER",
            ReviewedAtUtc = DateTimeOffset.UnixEpoch,
            SourceKnowledgeHash = contentHash,
            SourceKnowledgeId = knowledge.KnowledgeId,
            Decision = "approve",
            Notes = "Replace reviewer and notes after human review.",
            PromotesPageObjects = false
        };
        WriteAtomicJson(Path.Combine(outputDirectory, "approval-template.json"), template, overwrite);
    }

    public void WriteApproved(
        string outputDirectory,
        ApplicationKnowledgeDocument knowledge,
        bool overwrite = false)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputDirectory);
        ArgumentNullException.ThrowIfNull(knowledge);

        if (!string.Equals(knowledge.Status, "approved", StringComparison.Ordinal))
            throw new InvalidOperationException("WriteApproved requires status=approved.");
        if (knowledge.Approval is null)
            throw new InvalidOperationException("WriteApproved requires an approval record.");

        // Strip machine-specific absolute paths before persisting approved knowledge.
        SanitizeAbsolutePaths(knowledge);

        Directory.CreateDirectory(outputDirectory);
        WriteDocument(outputDirectory, knowledge, overwrite);
        WriteAtomicJson(Path.Combine(outputDirectory, "knowledge-approval.json"), knowledge.Approval, overwrite);
        WriteAtomicText(
            Path.Combine(outputDirectory, "content-hash.txt"),
            ComputeContentHash(knowledge) + "\n",
            overwrite);
    }

    public static string ComputeContentHash(ApplicationKnowledgeDocument knowledge)
    {
        ArgumentNullException.ThrowIfNull(knowledge);
        // Canonical projection: excludes status, approval, and volatile timestamps so
        // candidate and approved documents share the same sourceKnowledgeHash.
        var projection = BuildCanonicalProjection(knowledge);
        return KnowledgeHash.Sha256Hex(KnowledgeJson.CanonicalSerialize(projection));
    }

    public static object BuildCanonicalProjection(ApplicationKnowledgeDocument knowledge) =>
        new
        {
            knowledge.KnowledgeSchemaVersion,
            knowledge.ApplicationId,
            knowledge.BuildFingerprint,
            fingerprint = ApplicationFingerprintBuilder.ToHashable(knowledge.Fingerprint),
            sourceExplorationRuns = knowledge.SourceExplorationRuns
                .OrderBy(s => s.ExplorationId, StringComparer.Ordinal)
                .Select(s => new
                {
                    s.ExplorationId,
                    s.ArtifactHash,
                    s.PolicySha256,
                    manifestPath = SanitizeRelativePath(s.ManifestPath)
                }),
            graph = ProjectGraph(knowledge.Graph),
            workflows = knowledge.Workflows
                .OrderBy(w => w.WorkflowId, StringComparer.Ordinal)
                .Select(w => new
                {
                    w.WorkflowId,
                    w.EntryStateId,
                    w.TerminalStateId,
                    steps = w.Steps.Select(s => new { s.Ordinal, s.TransitionId, s.Action, s.SourceStateId, s.TargetStateId, s.Outcome }),
                    w.ObservationCount
                }),
            coverage = knowledge.Coverage,
            idCollisions = knowledge.IdCollisions
        };

    public static void SanitizeAbsolutePaths(ApplicationKnowledgeDocument knowledge)
    {
        ArgumentNullException.ThrowIfNull(knowledge);
        foreach (var source in knowledge.SourceExplorationRuns)
            source.ManifestPath = SanitizeRelativePath(source.ManifestPath);

        if (knowledge.Approval is not null)
            knowledge.Approval.ReviewedInputPath = SanitizeRelativePath(knowledge.Approval.ReviewedInputPath);
    }

    private static string? SanitizeRelativePath(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return null;

        var normalized = path.Replace('\\', '/');
        if (Path.IsPathRooted(path)
            || normalized.Contains(":/", StringComparison.Ordinal)
            || normalized.StartsWith("//", StringComparison.Ordinal))
        {
            var marker = "/exploration/runs/";
            var idx = normalized.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
            if (idx >= 0)
                return "exploration/runs/" + normalized[(idx + marker.Length)..];

            return Path.GetFileName(normalized);
        }

        return normalized;
    }

    private static void WriteDocument(string outputDirectory, ApplicationKnowledgeDocument knowledge, bool overwrite)
    {
        WriteAtomicJson(Path.Combine(outputDirectory, "application-knowledge.json"), knowledge, overwrite);
        WriteAtomicJson(Path.Combine(outputDirectory, "state-graph.json"), knowledge.Graph, overwrite);
        WriteAtomicJson(Path.Combine(outputDirectory, "workflows.json"), knowledge.Workflows, overwrite);
        WriteAtomicJson(Path.Combine(outputDirectory, "coverage.json"), knowledge.Coverage, overwrite);
    }

    private static object ProjectGraph(ApplicationStateGraph graph) =>
        new
        {
            graph.GraphSchemaVersion,
            windows = graph.Windows.Select(w => new
            {
                w.WindowId,
                w.Title,
                w.ClassName,
                w.IsModal,
                w.Role,
                w.ObservationCount
            }),
            states = graph.States.Select(s => new
            {
                s.StateId,
                s.SemanticFingerprint,
                s.WindowId,
                s.Label,
                s.IsEntry,
                s.IsTerminal,
                s.ObservedValues,
                s.ElementIds,
                s.ObservationCount,
                provenance = s.Provenance
            }),
            elements = graph.Elements.Select(e => new
            {
                e.ElementId,
                e.Identity,
                e.Capabilities,
                locatorEvidence = e.LocatorEvidence,
                e.ApprovedObjectRef,
                e.ObservationCount,
                provenance = e.Provenance
            }),
            transitions = graph.Transitions.Select(t => new
            {
                t.TransitionId,
                t.Action,
                t.ActionKind,
                t.ActionCategory,
                t.ActionParams,
                t.ElementId,
                t.SourceStateId,
                t.TargetStateId,
                t.Outcome,
                t.ConflictGroupId,
                t.ObservationCount,
                evidence = t.Evidence.Select(ev => new { ev.ExplorationId, ev.TransitionId, ev.Detail })
            }),
            graph.EntryStateIds,
            graph.TerminalStateIds,
            graph.Dialogs,
            graph.ChildWindows,
            graph.UnresolvedElements,
            graph.AmbiguousElements
        };

    private static void WriteAtomicJson<T>(string path, T value, bool overwrite)
    {
        if (File.Exists(path) && !overwrite)
            throw new IOException($"Knowledge artifact already exists and must not be overwritten: {path}");

        var json = SecretRedactor.Redact(JsonSerializer.Serialize(value, KnowledgeJson.Indented));
        WriteAtomicText(path, json.EndsWith('\n') ? json : json + "\n", overwrite: true);
    }

    private static void WriteAtomicText(string path, string content, bool overwrite = false)
    {
        if (File.Exists(path) && !overwrite)
            throw new IOException($"Knowledge artifact already exists and must not be overwritten: {path}");

        var directory = Path.GetDirectoryName(path)
                        ?? throw new IOException($"Unable to resolve directory for '{path}'.");
        Directory.CreateDirectory(directory);

        var tempPath = Path.Combine(directory, $".{Path.GetFileName(path)}.{Guid.NewGuid():N}.tmp");
        try
        {
            using (var stream = new FileStream(tempPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            using (var writer = new StreamWriter(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false)))
            {
                writer.Write(content);
                writer.Flush();
                stream.Flush(true);
            }

            if (File.Exists(path))
                File.Delete(path);
            File.Move(tempPath, path);
        }
        catch
        {
            try
            {
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
            catch
            {
                // best-effort cleanup
            }

            throw;
        }
    }
}
