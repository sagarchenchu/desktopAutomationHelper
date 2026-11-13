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
        var planningRoot = PlanningPathSafety.PlanningRoot(_workspace.RootPath);
        Directory.CreateDirectory(planningRoot);
        PlanningPathSafety.EnsureWritablePathInside(dir, planningRoot);
        if (Directory.Exists(dir) && Directory.EnumerateFileSystemEntries(dir).Any())
            throw new IOException($"Planning artifact directory '{dir}' already exists and must not be overwritten.");

        Directory.CreateDirectory(dir);
        PlanningPathSafety.EnsureWritablePathInside(dir, planningRoot);
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

        var planningRoot = PlanningPathSafety.PlanningRoot(_workspace.RootPath);
        PlanningPathSafety.EnsureWritablePathInside(targetPath, planningRoot);

        var dir = Path.GetDirectoryName(targetPath)!;
        Directory.CreateDirectory(dir);
        PlanningPathSafety.EnsureWritablePathInside(dir, planningRoot);
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

            PlanningPathSafety.EnsureWritablePathInside(tempPath, planningRoot);
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
}
