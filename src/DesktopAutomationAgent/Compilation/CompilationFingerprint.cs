using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using DesktopAutomationAgent.Driver.Models;

namespace DesktopAutomationAgent.Compilation;

public sealed class CompilationFingerprintInput
{
    public required string InputSha256 { get; init; }

    public required string JiraRevision { get; init; }

    public required string CompilerVersion { get; init; }

    public required string ParserVersion { get; init; }

    public required string StepRegistrySha256 { get; init; }

    public required string ObjectRepositorySha256 { get; init; }

    public string? ApprovedKnowledgeSha256 { get; init; }

    public string? CatalogFingerprint { get; init; }
}

public static class CompilationFingerprint
{
    public static string Compute(CompilationFingerprintInput input)
    {
        ArgumentNullException.ThrowIfNull(input);

        var payload = new SortedDictionary<string, string?>(StringComparer.Ordinal)
        {
            ["catalogFingerprint"] = input.CatalogFingerprint,
            ["compilerVersion"] = input.CompilerVersion,
            ["inputSha256"] = input.InputSha256.ToLowerInvariant(),
            ["jiraRevision"] = input.JiraRevision,
            ["knowledgeSha256"] = input.ApprovedKnowledgeSha256?.ToLowerInvariant(),
            ["objectRepositorySha256"] = input.ObjectRepositorySha256.ToLowerInvariant(),
            ["parserVersion"] = input.ParserVersion,
            ["stepRegistrySha256"] = input.StepRegistrySha256.ToLowerInvariant()
        };

        var canonical = CompilationJson.CanonicalSerialize(payload);
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(canonical))).ToLowerInvariant();
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

    public static string ComputeDirectorySha256(string directory, IEnumerable<string> relativeFiles)
    {
        var entries = new SortedDictionary<string, string>(StringComparer.Ordinal);
        foreach (var relative in relativeFiles.OrderBy(f => f.Replace('\\', '/'), StringComparer.Ordinal))
        {
            var full = Path.Combine(directory, relative.Replace('/', Path.DirectorySeparatorChar));
            if (!File.Exists(full))
                continue;
            entries[relative.Replace('\\', '/')] = ComputeFileSha256(full);
        }

        var canonical = CompilationJson.CanonicalSerialize(entries);
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(canonical))).ToLowerInvariant();
    }
}
