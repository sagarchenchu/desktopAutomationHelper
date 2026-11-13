using System.Text.Json;
using System.Text.Json.Serialization;
using DesktopAutomationAgent.ObjectRepository;
using DesktopAutomationAgent.Workspace;

namespace DesktopAutomationAgent.Recovery;

/// <summary>
/// Writes candidate repair proposals only under object-repository/candidates/.
/// Never edits active pages, approved knowledge, or the step registry.
/// Never auto-promotes.
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

    public CandidateRepairWriter(IWorkspaceManager workspace)
    {
        _workspace = workspace;
    }

    public string Write(CandidateRepairProposal proposal)
    {
        ArgumentNullException.ThrowIfNull(proposal);
        if (proposal.AutoPromoted)
            throw new InvalidOperationException("Candidate repairs must never be auto-promoted.");

        _workspace.EnsureInitialized();
        var candidatesRoot = Path.GetFullPath(
            Path.Combine(_workspace.RootPath, "object-repository", "candidates"));
        Directory.CreateDirectory(candidatesRoot);

        var folder = Path.Combine(candidatesRoot, "recovery", proposal.ProposalId);
        var fullFolder = Path.GetFullPath(folder);
        EnsureUnderCandidates(fullFolder, candidatesRoot);
        Directory.CreateDirectory(fullFolder);

        // Defense: never touch approved pages or knowledge.
        var pagesRoot = Path.GetFullPath(Path.Combine(_workspace.RootPath, "object-repository", "pages"));
        if (fullFolder.StartsWith(
                pagesRoot.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar,
                StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("Candidate repairs must not write under object-repository/pages.");
        }

        var payload = new
        {
            proposal.ProposalId,
            proposal.StepId,
            proposal.Evidence,
            proposal.Confidence,
            oldLocator = proposal.OldLocator,
            newLocator = proposal.NewLocator,
            proposal.ApplicationFingerprint,
            proposal.VerificationCommand,
            autoPromoted = false,
            requiresHumanApproval = true,
            proposal.ObjectReference,
            writtenAtUtc = DateTimeOffset.UtcNow
        };

        var target = Path.Combine(fullFolder, "candidate-repair.json");
        if (File.Exists(target))
            throw new IOException($"Candidate repair already exists at '{target}'.");

        var temp = target + $".{Guid.NewGuid():N}.tmp";
        try
        {
            File.WriteAllText(temp, JsonSerializer.Serialize(payload, JsonOptions));
            File.Move(temp, target);
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

    private static void EnsureUnderCandidates(string path, string candidatesRoot)
    {
        var full = Path.GetFullPath(path);
        var prefix = candidatesRoot.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                     + Path.DirectorySeparatorChar;
        if (!full.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
            && !string.Equals(full, candidatesRoot, StringComparison.OrdinalIgnoreCase))
        {
            throw new IOException($"Candidate repairs must stay under '{candidatesRoot}'.");
        }
    }
}
