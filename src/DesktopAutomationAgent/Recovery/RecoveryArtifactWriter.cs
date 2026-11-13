using System.Text.Json;
using System.Text.Json.Serialization;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Workspace;

namespace DesktopAutomationAgent.Recovery;

/// <summary>
/// Writes recovery artifacts under runs/&lt;runId&gt;/recovery/ without mutating the plan.
/// </summary>
public sealed class RecoveryArtifactWriter
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) }
    };

    private readonly IWorkspaceManager _workspace;

    public RecoveryArtifactWriter(IWorkspaceManager workspace)
    {
        _workspace = workspace;
    }

    public string EnsureRecoveryDirectory(string runDirectory)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(runDirectory);
        var dir = Path.Combine(runDirectory, "recovery");
        Directory.CreateDirectory(dir);
        EnsureUnderRuns(dir);
        return dir;
    }

    public string WriteSnapshot(string recoveryDirectory, string label, object snapshot)
    {
        var safe = Sanitize(label);
        var path = Path.Combine(recoveryDirectory, $"{safe}.json");
        WriteJsonAtomic(path, snapshot);
        return path;
    }

    public void WriteSummary(string recoveryDirectory, RecoveryRunSummary summary) =>
        WriteJsonAtomic(Path.Combine(recoveryDirectory, "recovery-summary.json"), summary);

    public void WriteAttempt(string recoveryDirectory, RecoveryAttemptRecord attempt) =>
        WriteJsonAtomic(
            Path.Combine(
                recoveryDirectory,
                $"attempt-{attempt.AttemptNumber:00}-{Sanitize(attempt.StepId)}-tier-{(int)attempt.Tier}-{Sanitize(attempt.Outcome)}.json"),
            attempt);

    private void WriteJsonAtomic(string targetPath, object value)
    {
        EnsureUnderRuns(Path.GetDirectoryName(targetPath)!);
        if (File.Exists(targetPath))
            throw new IOException($"Recovery artifact already exists at '{targetPath}' and must not be overwritten.");

        var json = JsonSerializer.Serialize(value, JsonOptions);
        json = SecretRedactor.Redact(json);
        var tempPath = targetPath + $".{Guid.NewGuid():N}.tmp";
        try
        {
            File.WriteAllText(tempPath, json);
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

    private void EnsureUnderRuns(string path)
    {
        var full = Path.GetFullPath(path);
        var runsRoot = Path.GetFullPath(Path.Combine(_workspace.RootPath, "runs"));
        var prefix = runsRoot.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                     + Path.DirectorySeparatorChar;
        if (!full.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
            && !string.Equals(full, runsRoot, StringComparison.OrdinalIgnoreCase))
        {
            throw new IOException($"Recovery artifacts must stay under '{runsRoot}'.");
        }
    }

    private static string Sanitize(string value)
    {
        var trimmed = (value ?? "item").Trim();
        if (trimmed.Length == 0)
            trimmed = "item";
        foreach (var c in Path.GetInvalidFileNameChars())
            trimmed = trimmed.Replace(c, '_');
        return trimmed.Length > 64 ? trimmed[..64] : trimmed;
    }
}
