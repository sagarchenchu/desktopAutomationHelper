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

        if (!CompilationFingerprint.IsHexFingerprint(fingerprint))
        {
            errors.Add($"compiled fingerprint '{fingerprint}' must be a 64-char lowercase hex SHA-256.");
            return false;
        }

        var dir = GetArtifactDirectory(jiraKey, fingerprint);
        var planPath = Path.Combine(dir, "plan.json");
        var metaPath = Path.Combine(dir, "compilation.json");

        var planExists = File.Exists(planPath);
        var metaExists = File.Exists(metaPath);
        if (!planExists && !metaExists)
            return false;

        if (planExists != metaExists)
        {
            errors.Add(
                $"compiled artifact {jiraKey}/{fingerprint} is incomplete (partial/corrupt: plan={planExists}, compilation={metaExists}).");
            return false;
        }

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

            // Fail closed on null/missing required fields before any ToLowerInvariant / recompute.
            if (!TryValidateMetadataIntegrity(metadata, jiraKey, fingerprint, errors))
            {
                metadata = null;
                planBytes = null;
                return false;
            }

            if (!string.Equals(metadata.Fingerprint, fingerprint, StringComparison.OrdinalIgnoreCase))
            {
                errors.Add(
                    $"compiled artifact fingerprint mismatch for {jiraKey}/{fingerprint}: metadata has '{metadata.Fingerprint}'.");
                metadata = null;
                planBytes = null;
                return false;
            }

            var planSha = CompilationFingerprint.ComputeBytesSha256(planBytes);
            if (!string.Equals(metadata.PlanSha256, planSha, StringComparison.OrdinalIgnoreCase))
            {
                errors.Add($"compiled plan.json sha mismatch for {jiraKey}/{fingerprint}.");
                metadata = null;
                planBytes = null;
                return false;
            }

            return true;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException)
        {
            errors.Add($"failed to read compiled artifact {jiraKey}/{fingerprint}: {ex.Message}");
            planBytes = null;
            metadata = null;
            return false;
        }
    }

    internal static bool TryValidateMetadataIntegrity(
        CompilationMetadata metadata,
        string jiraKey,
        string fingerprint,
        List<string> errors)
    {
        var before = errors.Count;
        RequireMetadataField(metadata.JiraKey, "jiraKey", jiraKey, fingerprint, errors);
        RequireMetadataField(metadata.Fingerprint, "fingerprint", jiraKey, fingerprint, errors);
        RequireMetadataField(metadata.CompilerVersion, "compilerVersion", jiraKey, fingerprint, errors);
        RequireMetadataField(metadata.ParserVersion, "parserVersion", jiraKey, fingerprint, errors);
        RequireMetadataField(metadata.InputSha256, "inputSha256", jiraKey, fingerprint, errors);
        RequireMetadataField(metadata.JiraRevision, "jiraRevision", jiraKey, fingerprint, errors);
        RequireMetadataField(metadata.StepRegistrySha256, "stepRegistrySha256", jiraKey, fingerprint, errors);
        RequireMetadataField(metadata.ObjectRepositorySha256, "objectRepositorySha256", jiraKey, fingerprint, errors);
        // objectRepositoryPath may be empty only if explicitly present; null is corrupt.
        if (metadata.ObjectRepositoryPath is null)
        {
            errors.Add(
                $"compiled artifact {jiraKey}/{fingerprint} compilation.json has null objectRepositoryPath.");
        }

        RequireMetadataField(metadata.ParameterBindingsSha256, "parameterBindingsSha256", jiraKey, fingerprint, errors);
        RequireMetadataField(metadata.PlanSha256, "planSha256", jiraKey, fingerprint, errors);
        return errors.Count == before;
    }

    private static void RequireMetadataField(
        string? value,
        string fieldName,
        string jiraKey,
        string fingerprint,
        List<string> errors)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            errors.Add(
                $"compiled artifact {jiraKey}/{fingerprint} compilation.json missing or null required field '{fieldName}'.");
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
        if (!CompilationFingerprint.IsHexFingerprint(fingerprint))
            throw new ArgumentException($"Fingerprint '{fingerprint}' must be a 64-char hex SHA-256.");

        if (!JiraKeyLooksSafe(jiraKey))
            throw new ArgumentException($"Jira key '{jiraKey}' is not a safe path segment.");

        var finalDir = GetArtifactDirectory(jiraKey, fingerprint);
        planPath = Path.Combine(finalDir, "plan.json");
        compilationPath = Path.Combine(finalDir, "compilation.json");

        if (Directory.Exists(finalDir) || File.Exists(planPath) || File.Exists(compilationPath))
        {
            if (TryReadExisting(jiraKey, fingerprint, out var existingPlan, out _, out var readErrors)
                && existingPlan is not null
                && existingPlan.AsSpan().SequenceEqual(planBytes))
            {
                // Identical cache hit — do not rewrite.
                return;
            }

            var detail = readErrors.Count > 0 ? string.Join("; ", readErrors) : "content differs";
            throw new IOException(
                $"refusing to overwrite existing compiled artifact at '{finalDir}' ({detail}).");
        }

        var parent = Path.Combine(_compiledRoot, jiraKey);
        Directory.CreateDirectory(parent);

        var tempDir = Path.Combine(parent, $".tmp-{Guid.NewGuid():N}");
        try
        {
            Directory.CreateDirectory(tempDir);
            var tempPlan = Path.Combine(tempDir, "plan.json");
            var tempMeta = Path.Combine(tempDir, "compilation.json");

            var metaJson = JsonSerializer.Serialize(metadata, CompilationJson.Indented);
            if (!metaJson.EndsWith('\n'))
                metaJson += "\n";

            WriteFileCreateNew(tempPlan, planBytes);
            WriteFileCreateNew(tempMeta, Encoding.UTF8.GetBytes(metaJson));

            // Atomic publish: rename sibling temp directory into place.
            Directory.Move(tempDir, finalDir);
        }
        catch
        {
            try
            {
                if (Directory.Exists(tempDir))
                    Directory.Delete(tempDir, recursive: true);
            }
            catch
            {
                // best-effort cleanup of partial temp sibling
            }

            throw;
        }
    }

    private static void WriteFileCreateNew(string path, byte[] content)
    {
        using var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        stream.Write(content);
        stream.Flush(true);
    }

    private static bool JiraKeyLooksSafe(string jiraKey) =>
        !string.IsNullOrWhiteSpace(jiraKey)
        && jiraKey.IndexOfAny(['/', '\\', ':', '.', ' ']) < 0
        && jiraKey.IndexOf("..", StringComparison.Ordinal) < 0;
}
