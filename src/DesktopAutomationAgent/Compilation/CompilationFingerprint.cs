using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using DesktopAutomationAgent.Driver.Models;

namespace DesktopAutomationAgent.Compilation;

public sealed record CompilationFingerprintInput
{
    public required string InputSha256 { get; init; }

    public required string JiraRevision { get; init; }

    public required string CompilerVersion { get; init; }

    public required string ParserVersion { get; init; }

    public required string StepRegistrySha256 { get; init; }

    public required string ObjectRepositorySha256 { get; init; }

    /// <summary>Workspace-relative object-repository path, already canonicalized.</summary>
    public required string ObjectRepositoryPath { get; init; }

    public string? ApprovedKnowledgeSha256 { get; init; }

    public string? CatalogFingerprint { get; init; }

    /// <summary>SHA-256 of canonical sorted parameter bindings (empty map => empty string hash input).</summary>
    public string? ParameterBindingsSha256 { get; init; }
}

public static class CompilationFingerprint
{
    private static readonly Regex Hex64 = new(
        @"^[a-f0-9]{64}$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled,
        BddCompilerVersions.RegexMatchTimeout);

    public static bool IsHexFingerprint(string? fingerprint) =>
        !string.IsNullOrWhiteSpace(fingerprint) && Hex64.IsMatch(fingerprint);

    public static string Compute(CompilationFingerprintInput input)
    {
        ArgumentNullException.ThrowIfNull(input);

        // Fail closed before ToLowerInvariant — null/missing required components must not NRE.
        var inputSha = RequireNonEmpty(input.InputSha256, nameof(input.InputSha256));
        var jiraRevision = RequireNonEmpty(input.JiraRevision, nameof(input.JiraRevision));
        var compilerVersion = RequireNonEmpty(input.CompilerVersion, nameof(input.CompilerVersion));
        var parserVersion = RequireNonEmpty(input.ParserVersion, nameof(input.ParserVersion));
        var stepRegistrySha = RequireNonEmpty(input.StepRegistrySha256, nameof(input.StepRegistrySha256));
        var objectRepoSha = RequireNonEmpty(input.ObjectRepositorySha256, nameof(input.ObjectRepositorySha256));
        _ = input.ObjectRepositoryPath
            ?? throw new ArgumentException("ObjectRepositoryPath is required (may be empty).", nameof(input));

        var payload = new SortedDictionary<string, string?>(StringComparer.Ordinal)
        {
            ["catalogFingerprint"] = NullIfEmpty(input.CatalogFingerprint)?.ToLowerInvariant(),
            ["compilerVersion"] = compilerVersion,
            ["inputSha256"] = inputSha.ToLowerInvariant(),
            ["jiraRevision"] = jiraRevision,
            ["knowledgeSha256"] = NullIfEmpty(input.ApprovedKnowledgeSha256)?.ToLowerInvariant(),
            ["objectRepositoryPath"] = CanonicalizeWorkspaceRelativePath(input.ObjectRepositoryPath),
            ["objectRepositorySha256"] = objectRepoSha.ToLowerInvariant(),
            ["parameterBindingsSha256"] = NullIfEmpty(input.ParameterBindingsSha256)?.ToLowerInvariant(),
            ["parserVersion"] = parserVersion,
            ["stepRegistrySha256"] = stepRegistrySha.ToLowerInvariant()
        };

        var canonical = CompilationJson.CanonicalSerialize(payload);
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(canonical))).ToLowerInvariant();
    }

    public static string CanonicalizeWorkspaceRelativePath(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return string.Empty;

        var normalized = path.Trim().Replace('\\', '/');
        while (normalized.StartsWith("./", StringComparison.Ordinal))
            normalized = normalized[2..];
        normalized = normalized.Trim('/');

        // Reject traversal / absolute forms — fingerprint must be a safe relative path.
        if (normalized.StartsWith('/')
            || normalized.Contains(':', StringComparison.Ordinal)
            || normalized.Split('/', StringSplitOptions.RemoveEmptyEntries)
                .Any(segment => segment is "." or ".."))
        {
            throw new ArgumentException($"Object repository path '{path}' is not a safe workspace-relative path.");
        }

        return normalized;
    }

    public static string ComputeParameterBindingsSha256(IReadOnlyDictionary<string, string>? bindings)
    {
        var ordered = new SortedDictionary<string, string>(StringComparer.Ordinal);
        if (bindings is not null)
        {
            foreach (var pair in bindings)
                ordered[pair.Key] = pair.Value ?? string.Empty;
        }

        var canonical = CompilationJson.CanonicalSerialize(ordered);
        return ComputeBytesSha256(Encoding.UTF8.GetBytes(canonical));
    }

    public static string ComputeFileSha256(string path)
    {
        var bytes = File.ReadAllBytes(path);
        return Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    }

    public static string ComputeBytesSha256(byte[] bytes) =>
        Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    public static string? ComputeCatalogFingerprint(OperationsCatalogDto? catalog)
    {
        if (catalog is null)
            return null;

        // Prefer first descriptor per case-insensitive name; fingerprint includes semantic fields
        // that affect PlanCatalogPreflight (aliases, deprecation, required inputs, session).
        var operations = catalog.Operations
            .Where(o => !string.IsNullOrWhiteSpace(o.Name))
            .GroupBy(o => o.Name.Trim(), StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First())
            .OrderBy(o => o.Name, StringComparer.Ordinal)
            .Select(o => new
            {
                name = o.Name.Trim(),
                aliases = o.Aliases
                    .Where(a => !string.IsNullOrWhiteSpace(a))
                    .Select(a => a.Trim())
                    .Distinct(StringComparer.Ordinal)
                    .OrderBy(a => a, StringComparer.Ordinal)
                    .ToArray(),
                deprecatedAliases = o.DeprecatedAliases
                    .Where(a => !string.IsNullOrWhiteSpace(a))
                    .Select(a => a.Trim())
                    .Distinct(StringComparer.Ordinal)
                    .OrderBy(a => a, StringComparer.Ordinal)
                    .ToArray(),
                deprecated = o.Deprecated,
                requiredInputs = o.RequiredInputs
                    .Where(i => !string.IsNullOrWhiteSpace(i))
                    .Select(i => i.Trim())
                    .Distinct(StringComparer.Ordinal)
                    .OrderBy(i => i, StringComparer.Ordinal)
                    .ToArray(),
                requiredInputAlternatives = o.RequiredInputAlternatives
                    .Select(group => group
                        .Where(i => !string.IsNullOrWhiteSpace(i))
                        .Select(i => i.Trim())
                        .Distinct(StringComparer.Ordinal)
                        .OrderBy(i => i, StringComparer.Ordinal)
                        .ToArray())
                    .Where(group => group.Length > 0)
                    .OrderBy(group => string.Join('\u001f', group), StringComparer.Ordinal)
                    .ToArray(),
                requiresSession = o.RequiresSession
            })
            .ToArray();

        var payload = new
        {
            schemaVersion = catalog.SchemaVersion,
            driverVersion = catalog.DriverVersion ?? string.Empty,
            operations
        };

        var canonical = CompilationJson.CanonicalSerialize(payload);
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(canonical))).ToLowerInvariant();
    }

    private static string RequireNonEmpty(string? value, string name)
    {
        if (string.IsNullOrWhiteSpace(value))
            throw new ArgumentException($"{name} is required and must be non-empty.", name);
        return value;
    }

    private static string? NullIfEmpty(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value;
}
