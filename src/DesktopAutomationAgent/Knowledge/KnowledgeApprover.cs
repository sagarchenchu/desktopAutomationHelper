using System.Text.Json;
using DesktopAutomationAgent.Workspace;

namespace DesktopAutomationAgent.Knowledge;

public sealed class KnowledgeApprover
{
    private readonly IWorkspaceManager _workspace;
    private readonly KnowledgeArtifactWriter _writer;
    private readonly KnowledgeValidator _validator;
    private readonly int _maxFileBytes;

    public KnowledgeApprover(
        IWorkspaceManager workspace,
        KnowledgeArtifactWriter writer,
        KnowledgeValidator validator,
        int maxFileBytes = KnowledgeConstants.MaxFileBytesDefault)
    {
        _workspace = workspace;
        _writer = writer;
        _validator = validator;
        _maxFileBytes = maxFileBytes;
    }

    public KnowledgeApproveResult Approve(
        string sourceDirectory,
        string reviewer,
        string reviewedInputPath)
    {
        try
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(sourceDirectory);
            ArgumentException.ThrowIfNullOrWhiteSpace(reviewer);
            ArgumentException.ThrowIfNullOrWhiteSpace(reviewedInputPath);

            _workspace.EnsureInitialized();

            var sourceDir = Path.GetFullPath(sourceDirectory);
            // Approval sources must be under knowledge/candidates — never approved/ or pages/.
            KnowledgePathSafety.EnsureCandidateSource(sourceDir, _workspace.RootPath);

            var knowledgePath = Path.Combine(sourceDir, "application-knowledge.json");
            if (!File.Exists(knowledgePath))
            {
                return Fail(5, $"Source knowledge document not found: {knowledgePath}");
            }

            EnsureFileSize(knowledgePath);
            KnowledgePathSafety.EnsureNotSymlinkEscape(knowledgePath, KnowledgePathSafety.CandidatesRoot(_workspace.RootPath));

            var reviewedFull = Path.GetFullPath(reviewedInputPath);
            if (!File.Exists(reviewedFull))
                return Fail(2, $"Reviewed input file not found: {reviewedInputPath}");
            EnsureFileSize(reviewedFull);

            var reviewedJson = File.ReadAllText(reviewedFull);
            var reviewedInput = JsonSerializer.Deserialize<KnowledgeApproval>(reviewedJson, KnowledgeJson.Compact)
                                ?? throw new InvalidOperationException("Failed to deserialize reviewed input.");

            if (!string.Equals(reviewedInput.Decision, "approve", StringComparison.Ordinal))
                return Fail(2, "reviewed input decision must be 'approve'.");

            if (reviewedInput.PromotesPageObjects)
            {
                return Fail(
                    2,
                    "reviewed input sets promotesPageObjects=true, which is forbidden. Knowledge approval never promotes Page Objects.");
            }

            if (string.IsNullOrWhiteSpace(reviewedInput.SourceKnowledgeHash))
                return Fail(2, "reviewed input must include sourceKnowledgeHash.");

            var knowledge = JsonSerializer.Deserialize<ApplicationKnowledgeDocument>(
                                File.ReadAllText(knowledgePath),
                                KnowledgeJson.Compact)
                            ?? throw new InvalidOperationException("Failed to deserialize source knowledge.");

            var actualHash = KnowledgeArtifactWriter.ComputeContentHash(knowledge);
            if (!string.Equals(actualHash, reviewedInput.SourceKnowledgeHash.Trim().ToLowerInvariant(), StringComparison.Ordinal))
            {
                return Fail(
                    5,
                    $"sourceKnowledgeHash mismatch: reviewed input has '{reviewedInput.SourceKnowledgeHash}', actual content hash is '{actualHash}'.");
            }

            var effectiveReviewer = string.IsNullOrWhiteSpace(reviewedInput.Reviewer)
                ? reviewer.Trim()
                : reviewedInput.Reviewer.Trim();
            if (!string.Equals(effectiveReviewer, reviewer.Trim(), StringComparison.Ordinal))
            {
                return Fail(
                    2,
                    $"reviewer mismatch: CLI --reviewer '{reviewer}' does not match reviewed input reviewer '{reviewedInput.Reviewer}'.");
            }

            var validation = _validator.Validate(knowledge);
            if (!validation.IsValid)
                return Fail(5, "Source knowledge failed validation: " + string.Join("; ", validation.Errors));

            var approval = new KnowledgeApproval
            {
                ApprovalSchemaVersion = KnowledgeConstants.ApprovalSchemaVersion,
                Reviewer = effectiveReviewer,
                ReviewedAtUtc = reviewedInput.ReviewedAtUtc == default
                    ? DateTimeOffset.UnixEpoch
                    : reviewedInput.ReviewedAtUtc,
                SourceKnowledgeHash = actualHash,
                SourceKnowledgeId = knowledge.KnowledgeId,
                Decision = "approve",
                Notes = reviewedInput.Notes,
                PromotesPageObjects = false,
                // Store basename only — never machine-specific absolute paths in approved knowledge.
                ReviewedInputPath = Path.GetFileName(reviewedFull),
                ReviewedInputHash = KnowledgeHash.Sha256Hex(reviewedJson)
            };

            knowledge.Status = "approved";
            knowledge.Approval = approval;
            knowledge.UpdatedAtUtc = approval.ReviewedAtUtc;
            KnowledgeArtifactWriter.SanitizeAbsolutePaths(knowledge);

            var approvedRoot = Path.Combine(
                _workspace.RootPath,
                "knowledge",
                "approved",
                KnowledgePathSafety.SanitizePathSegment(knowledge.ApplicationId, "applicationId"),
                KnowledgePathSafety.SanitizePathSegment(knowledge.BuildFingerprint, "buildFingerprint"));

            KnowledgePathSafety.EnsureWritablePathInside(approvedRoot, _workspace.RootPath);
            Directory.CreateDirectory(approvedRoot);

            // Never write into object-repository/pages or mutate repository.json here.
            _writer.WriteApproved(approvedRoot, knowledge, overwrite: false);

            return new KnowledgeApproveResult
            {
                Success = true,
                ExitCode = 0,
                OutputDirectory = approvedRoot,
                Knowledge = knowledge
            };
        }
        catch (KnowledgePathException ex)
        {
            return Fail(5, ex.Message);
        }
        catch (IOException ex)
        {
            return Fail(5, ex.Message);
        }
        catch (Exception ex) when (ex is InvalidOperationException or JsonException or UnauthorizedAccessException)
        {
            return Fail(5, ex.Message);
        }
    }

    private void EnsureFileSize(string path)
    {
        var info = new FileInfo(path);
        if (info.Length > _maxFileBytes)
            throw new InvalidOperationException($"File '{path}' exceeds maximum size of {_maxFileBytes} bytes.");
    }

    private static KnowledgeApproveResult Fail(int exitCode, string error) =>
        new()
        {
            Success = false,
            ExitCode = exitCode,
            Error = error
        };
}
