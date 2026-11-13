using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Plans;
using DesktopAutomationAgent.Workspace;

namespace DesktopAutomationAgent.Planning;

/// <summary>
/// Writes redacted, bounded, non-overwriting planning artifacts under
/// automation/planning/&lt;planningId&gt;/.
/// </summary>
public sealed class PlanningArtifactWriter
{
    private readonly IWorkspaceManager _workspace;

    public PlanningArtifactWriter(IWorkspaceManager workspace)
    {
        _workspace = workspace;
    }

    public string CreatePlanningDirectory(string planningId)
    {
        var safeId = SanitizePlanningId(planningId);
        // Resolve through workspace so path traversal is rejected.
        var dir = _workspace.ResolveSafePath(Path.Combine("planning", safeId));
        var planningRoot = EnsurePlanningRootInsideWorkspace();
        EnsureWritableUnderPlanning(dir, planningRoot);
        if (Directory.Exists(dir) && Directory.EnumerateFileSystemEntries(dir).Any())
            throw new IOException($"Planning artifact directory '{dir}' already exists and must not be overwritten.");

        Directory.CreateDirectory(dir);
        // Re-validate after create — planning/ or the id directory may be a junction.
        planningRoot = EnsurePlanningRootInsideWorkspace();
        EnsureWritableUnderPlanning(dir, planningRoot);
        return dir;
    }

    public void WriteAll(
        string directory,
        PlanningRequestSummary requestSummary,
        RetrievalManifest retrieval,
        PlanProposalDocument? proposal,
        PlanningValidationReport validation,
        PlanManifest? candidatePlan)
    {
        WriteJsonAtomic(Path.Combine(directory, "request-summary.json"), requestSummary);
        WriteJsonAtomic(Path.Combine(directory, "retrieval-manifest.json"), retrieval);
        if (proposal is not null)
            WriteJsonAtomic(Path.Combine(directory, "proposal.json"), RedactProposal(proposal));
        WriteJsonAtomic(Path.Combine(directory, "validation.json"), validation);
        if (candidatePlan is not null)
            WriteJsonAtomic(Path.Combine(directory, "candidate.plan.json"), candidatePlan);
    }

    public static string SanitizePlanningId(string planningId)
    {
        if (string.IsNullOrWhiteSpace(planningId))
            throw new ArgumentException("planningId is required.", nameof(planningId));

        var trimmed = planningId.Trim();
        if (trimmed.Length > 128
            || trimmed.Contains('/')
            || trimmed.Contains('\\')
            || trimmed.Contains(':')
            || trimmed.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
        {
            throw new ArgumentException($"planningId '{planningId}' is not a safe path segment.", nameof(planningId));
        }

        return trimmed;
    }

    public static string NewPlanningId() =>
        "pln_" + DateTimeOffset.UtcNow.ToString("yyyyMMddTHHmmss") + "_" +
        Convert.ToHexString(RandomNumberGenerator.GetBytes(4)).ToLowerInvariant();

    public static string HashObject<T>(T value) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(PlanningJson.CanonicalSerialize(value))))
            .ToLowerInvariant();

    private static PlanProposalDocument RedactProposal(PlanProposalDocument proposal)
    {
        // Deep clone via JSON then redact string fields.
        var json = SecretRedactor.Redact(JsonSerializer.Serialize(proposal, PlanningJson.Compact));
        return JsonSerializer.Deserialize<PlanProposalDocument>(json, PlanningJson.Compact) ?? proposal;
    }

    private void WriteJsonAtomic(string targetPath, object payload)
    {
        if (File.Exists(targetPath))
            throw new IOException($"Planning artifact '{targetPath}' already exists and must not be overwritten.");

        var planningRoot = EnsurePlanningRootInsideWorkspace();
        EnsureWritableUnderPlanning(targetPath, planningRoot);

        var dir = Path.GetDirectoryName(targetPath)!;
        Directory.CreateDirectory(dir);
        planningRoot = EnsurePlanningRootInsideWorkspace();
        EnsureWritableUnderPlanning(dir, planningRoot);
        var tempPath = Path.Combine(dir, $".{Path.GetFileName(targetPath)}.{Guid.NewGuid():N}.tmp");
        try
        {
            var json = SecretRedactor.Redact(JsonSerializer.Serialize(payload, PlanningJson.Indented));
            // Bound artifact size
            if (Encoding.UTF8.GetByteCount(json) > 1_048_576)
                throw new InvalidOperationException($"Planning artifact '{targetPath}' exceeds 1 MiB bound.");

            using (var stream = new FileStream(tempPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            using (var writer = new StreamWriter(stream, Encoding.UTF8))
            {
                writer.Write(json);
                writer.Flush();
                stream.Flush(true);
            }

            EnsureWritableUnderPlanning(tempPath, planningRoot);
            File.Move(tempPath, targetPath);
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
                // best-effort
            }

            throw;
        }
    }

    /// <summary>
    /// Ensures <c>planning/</c> exists and is not a junction/symlink escape outside the workspace.
    /// Workspace root is the outer containment boundary.
    /// </summary>
    private string EnsurePlanningRootInsideWorkspace()
    {
        var workspaceRoot = Path.GetFullPath(_workspace.RootPath)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var planningRoot = PlanningPathSafety.PlanningRoot(_workspace.RootPath);

        // Before create: reject an existing planning/ reparse that escapes the workspace.
        PlanningPathSafety.EnsureWritablePathInside(planningRoot, workspaceRoot);
        Directory.CreateDirectory(planningRoot);
        // After create: re-check in case create followed or replaced a link.
        PlanningPathSafety.EnsureWritablePathInside(planningRoot, workspaceRoot);

        return planningRoot;
    }

    private void EnsureWritableUnderPlanning(string fullPath, string planningRoot)
    {
        var workspaceRoot = Path.GetFullPath(_workspace.RootPath)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        // Inner: must stay under planning/. Outer: planning/ itself must stay under workspace.
        PlanningPathSafety.EnsureWritablePathInside(fullPath, planningRoot);
        PlanningPathSafety.EnsureWritablePathInside(fullPath, workspaceRoot);
    }
}
