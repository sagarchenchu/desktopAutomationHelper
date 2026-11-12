using System.Text;
using System.Text.Json;
using DesktopAutomationAgent.Compilation;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;

namespace DesktopAutomationAgent.Jira;

/// <summary>
/// Atomic redacted cache under <c>automation/jira-cache/&lt;jiraKey&gt;/&lt;revision-or-hash&gt;/</c>.
/// Stores only whitelist-projected Phase 4D input + source metadata — never unrestricted issue dumps.
/// </summary>
internal sealed class JiraRawCache
{
    private readonly string _root;

    public JiraRawCache(string workspaceRoot)
    {
        _root = Path.Combine(workspaceRoot, "jira-cache");
    }

    public string RootPath => _root;

    public string GetArtifactDirectory(string jiraKey, string revisionOrHash) =>
        Path.Combine(_root, jiraKey, SanitizeSegment(revisionOrHash));

    public bool TryRead(
        string jiraKey,
        string? revisionOrHash,
        out JiraTestCaseDocument? document,
        out JiraSourceMetadata? metadata,
        out string? directory,
        out List<string> errors)
    {
        document = null;
        metadata = null;
        directory = null;
        errors = [];

        if (!JiraKeyContract.IsCanonical(jiraKey))
        {
            errors.Add($"invalid jiraKey '{jiraKey}'.");
            return false;
        }

        string dir;
        if (!string.IsNullOrWhiteSpace(revisionOrHash))
        {
            dir = GetArtifactDirectory(jiraKey, revisionOrHash!);
            if (!Directory.Exists(dir))
            {
                errors.Add($"jira-cache entry not found for {jiraKey}/{revisionOrHash}.");
                return false;
            }
        }
        else
        {
            var keyRoot = Path.Combine(_root, jiraKey);
            if (!Directory.Exists(keyRoot))
            {
                errors.Add($"jira-cache entry not found for {jiraKey}.");
                return false;
            }

            var candidates = Directory.GetDirectories(keyRoot)
                .Where(d => !Path.GetFileName(d).StartsWith(".", StringComparison.Ordinal))
                .OrderByDescending(d => Directory.GetLastWriteTimeUtc(d))
                .ToArray();
            if (candidates.Length == 0)
            {
                errors.Add($"jira-cache entry not found for {jiraKey}.");
                return false;
            }

            dir = candidates[0];
        }

        var testCasePath = Path.Combine(dir, "test-case.json");
        var metaPath = Path.Combine(dir, "source-metadata.json");
        if (!File.Exists(testCasePath) || !File.Exists(metaPath))
        {
            errors.Add($"jira-cache entry incomplete at '{dir}'.");
            return false;
        }

        try
        {
            var testBytes = File.ReadAllBytes(testCasePath);
            metadata = JsonSerializer.Deserialize<JiraSourceMetadata>(
                File.ReadAllText(metaPath),
                CompilationJson.DocumentRead);
            if (metadata is null)
            {
                errors.Add($"jira-cache entry corrupt at '{dir}'.");
                return false;
            }

            if (!string.Equals(metadata.JiraKey, jiraKey, StringComparison.Ordinal))
            {
                errors.Add($"jira-cache metadata key mismatch at '{dir}'.");
                return false;
            }

            var display = ToDisplay(dir, "test-case.json");
            if (!JiraTestCaseProjector.TryValidateCachedBytes(
                    testBytes,
                    jiraKey,
                    metadata.ContentSha256,
                    display,
                    out document,
                    out var validateErrors))
            {
                errors.AddRange(validateErrors);
                return false;
            }

            directory = dir;
            return true;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException)
        {
            errors.Add($"failed to read jira-cache at '{dir}': {SecretRedactor.Redact(ex.Message)}");
            return false;
        }
    }

    public void WriteAtomic(
        string jiraKey,
        JiraTestCaseDocument document,
        JiraSourceMetadata metadata,
        out string directory,
        out JiraSourceMetadata writtenMetadata,
        out JiraTestCaseDocument writtenDocument)
    {
        if (!JiraKeyContract.IsCanonical(jiraKey))
            throw new ArgumentException($"invalid jiraKey '{jiraKey}'.");

        var display = $"jira-cache/{jiraKey}/test-case.json";
        if (!JiraTestCaseProjector.TrySanitizeAndHash(
                document,
                display,
                out var sanitized,
                out var testCaseBytes,
                out var contentSha,
                out var sanitizeErrors))
        {
            throw new InvalidOperationException(
                "refusing to cache invalid/unsanitized Jira test case: " +
                string.Join("; ", sanitizeErrors.Select(SecretRedactor.Redact)));
        }

        writtenDocument = sanitized;
        writtenMetadata = metadata with
        {
            JiraKey = jiraKey,
            ContentSha256 = contentSha,
            RevisionToken = string.IsNullOrWhiteSpace(metadata.RevisionToken)
                ? sanitized.ResolveRevisionToken()
                : metadata.RevisionToken,
            Redacted = true
        };
        metadata = writtenMetadata;

        var revision = SanitizeSegment(
            string.IsNullOrWhiteSpace(metadata.RevisionToken)
                ? contentSha
                : metadata.RevisionToken);

        var parent = Path.Combine(_root, jiraKey);
        Directory.CreateDirectory(parent);
        var finalDir = Path.Combine(parent, revision);
        directory = finalDir;

        if (Directory.Exists(finalDir))
        {
            if (TryRead(jiraKey, revision, out var existingDoc, out var existingMeta, out _, out _)
                && existingDoc is not null
                && existingMeta is not null
                && string.Equals(existingMeta.ContentSha256, contentSha, StringComparison.OrdinalIgnoreCase))
            {
                writtenDocument = existingDoc;
                writtenMetadata = existingMeta;
                return;
            }

            throw new IOException(
                $"refusing to overwrite jira-cache at '{finalDir}' with differing content.");
        }

        var tempDir = Path.Combine(parent, $".tmp-{Guid.NewGuid():N}");
        try
        {
            Directory.CreateDirectory(tempDir);
            var metaJson = SecretRedactor.Redact(JsonSerializer.Serialize(metadata, CompilationJson.Indented));
            if (!metaJson.EndsWith('\n'))
                metaJson += "\n";

            WriteCreateNew(Path.Combine(tempDir, "test-case.json"), testCaseBytes);
            WriteCreateNew(Path.Combine(tempDir, "source-metadata.json"), Encoding.UTF8.GetBytes(metaJson));
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
                // best-effort
            }

            throw;
        }
    }

    public ConditionalCacheHeaders? GetConditionalHeaders(string jiraKey)
    {
        if (!TryRead(jiraKey, null, out _, out var metadata, out _, out _))
            return null;
        if (metadata is null)
            return null;

        var headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (!string.IsNullOrWhiteSpace(metadata.ETag))
            headers["If-None-Match"] = metadata.ETag!;
        if (!string.IsNullOrWhiteSpace(metadata.LastModified))
            headers["If-Modified-Since"] = metadata.LastModified!;
        return headers.Count == 0 ? null : new ConditionalCacheHeaders(headers, metadata);
    }

    private static void WriteCreateNew(string path, byte[] content)
    {
        using var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        stream.Write(content);
        stream.Flush(true);
    }

    private static string SanitizeSegment(string value)
    {
        var trimmed = value.Trim();
        if (trimmed.Length == 0 || trimmed.Length > 128)
            throw new ArgumentException("revision-or-hash segment is empty or too long.");
        if (trimmed.IndexOfAny(['/', '\\', ':', ' ', '<', '>', '|', '?', '*']) >= 0
            || trimmed.Contains("..", StringComparison.Ordinal))
        {
            return JiraVendorMapper.Sha256Hex(Encoding.UTF8.GetBytes(trimmed))[..32];
        }

        return trimmed;
    }

    private static string ToDisplay(string dir, string file) =>
        Path.Combine(dir, file).Replace('\\', '/');
}

internal sealed record ConditionalCacheHeaders(
    IDictionary<string, string> Headers,
    JiraSourceMetadata Metadata);
