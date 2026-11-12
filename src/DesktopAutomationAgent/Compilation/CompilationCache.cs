using System.Text;
using System.Text.Json;

namespace DesktopAutomationAgent.Compilation;

public sealed class CompilationCache
{
    private readonly string _compiledRoot;

    public CompilationCache(string workspaceRoot)
    {
        _compiledRoot = Path.Combine(workspaceRoot, "compiled");
    }

    public string GetArtifactDirectory(string jiraKey, string fingerprint) =>
        Path.Combine(_compiledRoot, jiraKey, fingerprint);

    public bool TryReadExisting(
        string jiraKey,
        string fingerprint,
        out byte[]? planBytes,
        out CompilationMetadata? metadata,
        out List<string> errors)
    {
        planBytes = null;
        metadata = null;
        errors = [];

        var dir = GetArtifactDirectory(jiraKey, fingerprint);
        var planPath = Path.Combine(dir, "plan.json");
        var metaPath = Path.Combine(dir, "compilation.json");
        if (!File.Exists(planPath) || !File.Exists(metaPath))
            return false;

        try
        {
            planBytes = File.ReadAllBytes(planPath);
            var metaJson = File.ReadAllText(metaPath);
            metadata = JsonSerializer.Deserialize<CompilationMetadata>(metaJson, CompilationJson.DocumentRead);
            if (metadata is null)
            {
                errors.Add($"compiled artifact metadata missing for {jiraKey}/{fingerprint}.");
                return false;
            }

            if (!string.Equals(metadata.Fingerprint, fingerprint, StringComparison.OrdinalIgnoreCase))
            {
                errors.Add(
                    $"compiled artifact fingerprint mismatch for {jiraKey}/{fingerprint}: metadata has '{metadata.Fingerprint}'.");
                return false;
            }

            var planSha = CompilationFingerprint.ComputeBytesSha256(planBytes);
            if (!string.Equals(metadata.PlanSha256, planSha, StringComparison.OrdinalIgnoreCase))
            {
                errors.Add(
                    $"compiled plan.json sha mismatch for {jiraKey}/{fingerprint}.");
                return false;
            }

            return true;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException)
        {
            errors.Add($"failed to read compiled artifact {jiraKey}/{fingerprint}: {ex.Message}");
            return false;
        }
    }

    public void WriteAtomic(
        string jiraKey,
        string fingerprint,
        byte[] planBytes,
        CompilationMetadata metadata,
        out string planPath,
        out string compilationPath)
    {
        var dir = GetArtifactDirectory(jiraKey, fingerprint);
        Directory.CreateDirectory(dir);
        planPath = Path.Combine(dir, "plan.json");
        compilationPath = Path.Combine(dir, "compilation.json");

        if (File.Exists(planPath) || File.Exists(compilationPath))
        {
            // Never overwrite a different artifact. Identical bytes are a cache hit.
            if (File.Exists(planPath))
            {
                var existing = File.ReadAllBytes(planPath);
                if (!existing.AsSpan().SequenceEqual(planBytes))
                {
                    throw new IOException(
                        $"refusing to overwrite different compiled plan at '{planPath}'.");
                }
            }

            if (File.Exists(compilationPath))
            {
                // Metadata may differ only by CompiledAtUtc — do not overwrite on reuse path.
                return;
            }
        }

        var metaJson = JsonSerializer.Serialize(metadata, CompilationJson.Indented);
        if (!metaJson.EndsWith('\n'))
            metaJson += "\n";

        WriteAtomicBytes(planPath, planBytes);
        WriteAtomicText(compilationPath, metaJson);
    }

    private static void WriteAtomicBytes(string path, byte[] content)
    {
        if (File.Exists(path))
            throw new IOException($"Compiled artifact already exists and must not be overwritten: {path}");

        var directory = Path.GetDirectoryName(path)
                        ?? throw new IOException($"Unable to resolve directory for '{path}'.");
        Directory.CreateDirectory(directory);
        var tempPath = Path.Combine(directory, $".{Path.GetFileName(path)}.{Guid.NewGuid():N}.tmp");
        try
        {
            using (var stream = new FileStream(tempPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                stream.Write(content);
                stream.Flush(true);
            }

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
                // best-effort
            }

            throw;
        }
    }

    private static void WriteAtomicText(string path, string content)
    {
        WriteAtomicBytes(path, Encoding.UTF8.GetBytes(content));
    }
}
