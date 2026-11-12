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

        var payload = new SortedDictionary<string, string?>(StringComparer.Ordinal)
        {
            ["catalogFingerprint"] = NullIfEmpty(input.CatalogFingerprint),
            ["compilerVersion"] = input.CompilerVersion,
            ["inputSha256"] = input.InputSha256.ToLowerInvariant(),
            ["jiraRevision"] = input.JiraRevision,
            ["knowledgeSha256"] = NullIfEmpty(input.ApprovedKnowledgeSha256)?.ToLowerInvariant(),
            ["objectRepositoryPath"] = CanonicalizeWorkspaceRelativePath(input.ObjectRepositoryPath),
            ["objectRepositorySha256"] = input.ObjectRepositorySha256.ToLowerInvariant(),
            ["parameterBindingsSha256"] = NullIfEmpty(input.ParameterBindingsSha256)?.ToLowerInvariant(),
            ["parserVersion"] = input.ParserVersion,
            ["stepRegistrySha256"] = input.StepRegistrySha256.ToLowerInvariant()
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

        var operationNames = catalog.Operations
            .Select(o => o.Name)
            .Where(n => !string.IsNullOrWhiteSpace(n))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(n => n, StringComparer.Ordinal)
            .ToArray();

        var payload = new
        {
            schemaVersion = catalog.SchemaVersion,
            driverVersion = catalog.DriverVersion ?? string.Empty,
            operations = operationNames
        };

        var canonical = CompilationJson.CanonicalSerialize(payload);
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(canonical))).ToLowerInvariant();
    }

    private static string? NullIfEmpty(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value;
}
