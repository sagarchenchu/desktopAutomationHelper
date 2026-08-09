using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Workspace;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Plans;

public sealed class PlanManifestReader
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true
    };

    private readonly AgentOptions _options;
    private readonly IWorkspaceManager _workspace;

    public PlanManifestReader(IOptions<AgentOptions> options, IWorkspaceManager workspace)
    {
        _options = options.Value;
        _workspace = workspace;
    }

    public PlanValidationResult Read(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return Failure(
                path ?? string.Empty,
                sha256: null,
                ["Plan file path is required."]);
        }

        string fullPath;
        string relativePath;
        try
        {
            fullPath = _workspace.ResolveSafePath(path);
            relativePath = ToDisplayRelativePath(fullPath);
        }
        catch (WorkspaceException ex)
        {
            return Failure(path, sha256: null, [ex.Message]);
        }

        if (!File.Exists(fullPath))
        {
            return Failure(
                relativePath,
                sha256: null,
                [$"{relativePath}: plan file not found."]);
        }

        byte[] bytes;
        try
        {
            var fileInfo = new FileInfo(fullPath);
            if (fileInfo.Length > _options.Runner.MaxPlanBytes)
            {
                return Failure(
                    relativePath,
                    sha256: null,
                    [
                        $"{relativePath}: plan file exceeds maximum size of {_options.Runner.MaxPlanBytes} bytes."
                    ]);
            }

            bytes = File.ReadAllBytes(fullPath);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return Failure(
                relativePath,
                sha256: null,
                [$"{relativePath}: failed to read plan file ({ex.Message})."]);
        }

        var sha256 = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

        var duplicateErrors = JsonDuplicatePropertyDetector.DetectDuplicates(bytes)
            .Select(message => $"{relativePath}: {message}")
            .ToList();
        if (duplicateErrors.Count > 0)
        {
            return Failure(relativePath, sha256, duplicateErrors);
        }

        PlanManifest? manifest;
        try
        {
            manifest = JsonSerializer.Deserialize<PlanManifest>(bytes, JsonOptions);
        }
        catch (JsonException ex)
        {
            return Failure(
                relativePath,
                sha256,
                [$"{relativePath}: invalid JSON ({ex.Message})."]);
        }

        if (manifest is null)
        {
            return Failure(
                relativePath,
                sha256,
                [$"{relativePath}: plan manifest was empty."]);
        }

        var validator = new PlanValidator();
        var validation = validator.Validate(manifest, relativePath);
        return new PlanValidationResult
        {
            PlanPath = validation.PlanPath,
            PlanId = validation.PlanId,
            Name = validation.Name,
            StepCount = validation.StepCount,
            OnFailureStepCount = validation.OnFailureStepCount,
            CleanupStepCount = validation.CleanupStepCount,
            TotalStepCount = validation.TotalStepCount,
            Errors = validation.Errors,
            Warnings = validation.Warnings,
            Sha256 = sha256,
            Plan = validation.IsValid ? manifest : null
        };
    }

    private string ToDisplayRelativePath(string fullPath)
    {
        var root = Path.GetFullPath(_workspace.RootPath)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var normalized = Path.GetFullPath(fullPath)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

        if (WorkspaceManager.PathsEqual(normalized, root))
            return ".";

        var relative = Path.GetRelativePath(root, normalized);
        return relative.Replace(Path.DirectorySeparatorChar, '/');
    }

    private static PlanValidationResult Failure(string planPath, string? sha256, IReadOnlyList<string> errors) =>
        new()
        {
            PlanPath = planPath,
            Errors = errors,
            Sha256 = sha256
        };
}
