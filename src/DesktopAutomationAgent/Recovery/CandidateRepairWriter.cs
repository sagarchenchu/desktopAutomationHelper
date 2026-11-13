using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Workspace;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Recovery;

/// <summary>
/// Writes candidate repair proposals only under object-repository/candidates/.
/// Link-aware containment before and after create/write. Never auto-promotes.
/// </summary>
public sealed class CandidateRepairWriter
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    private readonly IWorkspaceManager _workspace;
    private readonly RecoveryOptions _options;

    public CandidateRepairWriter(IWorkspaceManager workspace, IOptions<AgentOptions> options)
    {
        _workspace = workspace;
        _options = options.Value.Recovery;
    }

    /// <summary>Test helper — not registered with DI (avoids ambiguous constructors).</summary>
    public static CandidateRepairWriter ForTests(IWorkspaceManager workspace, RecoveryOptions? recovery = null) =>
        new(workspace, Options.Create(new AgentOptions { Recovery = recovery ?? new RecoveryOptions() }));

    public string Write(CandidateRepairProposal proposal)
    {
        ArgumentNullException.ThrowIfNull(proposal);
        if (proposal.AutoPromoted)
            throw new InvalidOperationException("Candidate repairs must never be auto-promoted.");

        _workspace.EnsureInitialized();
        var candidatesRoot = RecoveryPathSafety.CandidatesRoot(_workspace.RootPath);
        Directory.CreateDirectory(candidatesRoot);
        RecoveryPathSafety.EnsureWritableUnderCandidates(candidatesRoot, _workspace.RootPath);

        var safeId = Sanitize(proposal.ProposalId);
        var folder = Path.Combine(candidatesRoot, "recovery", safeId);
        var fullFolder = Path.GetFullPath(folder);

        // Validate containment BEFORE create (rejects junctions that escape to pages/).
        RecoveryPathSafety.EnsureWritableUnderCandidates(fullFolder, _workspace.RootPath);
        Directory.CreateDirectory(fullFolder);
        // Re-validate AFTER create — folder may have been replaced by a junction.
        RecoveryPathSafety.EnsureWritableUnderCandidates(fullFolder, _workspace.RootPath);

        var payload = new
        {
            proposal.ProposalId,
            proposal.StepId,
            evidence = SecretRedactor.Redact(proposal.Evidence),
            proposal.Confidence,
            oldLocator = proposal.OldLocator,
            newLocator = proposal.NewLocator,
            proposal.ApplicationFingerprint,
            verificationCommand = SecretRedactor.Redact(proposal.VerificationCommand),
            autoPromoted = false,
            requiresHumanApproval = true,
            proposal.ObjectReference,
            writtenAtUtc = DateTimeOffset.UtcNow
        };

        var json = JsonSerializer.Serialize(payload, JsonOptions);
        json = SecretRedactor.Redact(json);
        var bytes = Encoding.UTF8.GetBytes(json);
        if (bytes.Length > _options.MaxCandidateArtifactBytes)
        {
            throw new IOException(
                $"Candidate repair artifact size {bytes.Length} exceeds Recovery:MaxCandidateArtifactBytes {_options.MaxCandidateArtifactBytes}.");
        }

        var target = Path.Combine(fullFolder, "candidate-repair.json");
        RecoveryPathSafety.EnsureWritableUnderCandidates(target, _workspace.RootPath);
        if (File.Exists(target))
            throw new IOException($"Candidate repair already exists at '{target}'.");

        var temp = target + $".{Guid.NewGuid():N}.tmp";
        try
        {
            File.WriteAllBytes(temp, bytes);
            RecoveryPathSafety.EnsureWritableUnderCandidates(temp, _workspace.RootPath);
            File.Move(temp, target);
            RecoveryPathSafety.EnsureWritableUnderCandidates(target, _workspace.RootPath);
        }
        catch
        {
            try
            {
                if (File.Exists(temp))
                    File.Delete(temp);
            }
            catch
            {
                // best-effort
            }

            throw;
        }

        return target;
    }

    private static string Sanitize(string value)
    {
        var trimmed = (value ?? "rrp").Trim();
        if (trimmed.Length == 0)
            trimmed = "rrp";
        foreach (var c in Path.GetInvalidFileNameChars())
            trimmed = trimmed.Replace(c, '_');
        return trimmed.Length > 64 ? trimmed[..64] : trimmed;
    }
}
