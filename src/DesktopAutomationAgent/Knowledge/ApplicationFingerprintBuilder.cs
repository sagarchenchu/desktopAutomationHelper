using DesktopAutomationAgent.Exploration;
using DesktopAutomationAgent.Observation;

namespace DesktopAutomationAgent.Knowledge;

public sealed class ApplicationFingerprintBuilder
{
    /// <summary>
    /// Build fingerprint exclusively from sealed exploration build identity (fail closed).
    /// </summary>
    public ApplicationFingerprint BuildFromIdentity(string applicationId, ExplorationBuildIdentity identity)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(applicationId);
        ArgumentNullException.ThrowIfNull(identity);

        if (string.IsNullOrWhiteSpace(identity.DriverVersion))
            throw new InvalidOperationException("build identity is missing required driverVersion.");
        if (string.IsNullOrWhiteSpace(identity.PolicySha256) || identity.PolicySha256.Length != 64)
            throw new InvalidOperationException("build identity is missing required policySha256.");
        if (string.IsNullOrWhiteSpace(identity.ExplorerVersion))
            throw new InvalidOperationException("build identity is missing required explorerVersion.");
        if (string.IsNullOrWhiteSpace(identity.ProcessName))
            throw new InvalidOperationException("build identity is missing required processName.");

        if (!identity.HasImmutableBuildId())
        {
            throw new InvalidOperationException(
                "build identity requires at least one immutable build id "
                + "(executableSha256, productVersion, or fileVersion).");
        }

        var fingerprint = new ApplicationFingerprint
        {
            ApplicationId = applicationId.Trim().ToLowerInvariant(),
            ExecutablePath = NormalizePath(identity.ExecutablePath),
            ExecutableSha256 = NormalizeSha(identity.ExecutableSha256),
            ProductVersion = string.IsNullOrWhiteSpace(identity.ProductVersion) ? null : identity.ProductVersion.Trim(),
            FileVersion = string.IsNullOrWhiteSpace(identity.FileVersion) ? null : identity.FileVersion.Trim(),
            DriverVersion = identity.DriverVersion.Trim(),
            SnapshotSchemaVersion = identity.SnapshotSchemaVersion > 0
                ? identity.SnapshotSchemaVersion
                : SemanticSnapshot.SchemaVersion,
            ExplorerVersion = identity.ExplorerVersion.Trim(),
            ExplorationPolicyHash = identity.PolicySha256.Trim().ToLowerInvariant()
        };

        foreach (var pair in identity.ConfigIdentity.OrderBy(p => p.Key, StringComparer.Ordinal))
        {
            if (IsSecretLike(pair.Key) || IsSecretLike(pair.Value))
                continue;
            fingerprint.ConfigIdentity[pair.Key.Trim()] = pair.Value.Trim();
        }

        // Always include processName in config identity for stable fingerprinting.
        fingerprint.ConfigIdentity["processName"] = identity.ProcessName.Trim();

        return fingerprint;
    }

    public string ComputeBuildFingerprint(ApplicationFingerprint fingerprint)
    {
        ArgumentNullException.ThrowIfNull(fingerprint);
        var canonical = KnowledgeJson.CanonicalSerialize(ToHashable(fingerprint));
        return KnowledgeHash.StableIdPrefix(KnowledgeHash.Sha256Hex(canonical), 32);
    }

    public bool AreCompatible(ApplicationFingerprint left, ApplicationFingerprint right, out List<string> mismatches)
    {
        ArgumentNullException.ThrowIfNull(left);
        ArgumentNullException.ThrowIfNull(right);
        mismatches = [];

        Compare(nameof(left.ApplicationId), left.ApplicationId, right.ApplicationId, mismatches);
        Compare(nameof(left.ExecutablePath), left.ExecutablePath, right.ExecutablePath, mismatches);
        Compare(nameof(left.ProductVersion), left.ProductVersion, right.ProductVersion, mismatches);
        Compare(nameof(left.FileVersion), left.FileVersion, right.FileVersion, mismatches);
        Compare(nameof(left.ExecutableSha256), left.ExecutableSha256, right.ExecutableSha256, mismatches);
        Compare(nameof(left.DriverVersion), left.DriverVersion, right.DriverVersion, mismatches);
        Compare(
            nameof(left.SnapshotSchemaVersion),
            left.SnapshotSchemaVersion.ToString(),
            right.SnapshotSchemaVersion.ToString(),
            mismatches);
        Compare(nameof(left.ExplorerVersion), left.ExplorerVersion, right.ExplorerVersion, mismatches);
        Compare(
            nameof(left.ExplorationPolicyHash),
            left.ExplorationPolicyHash,
            right.ExplorationPolicyHash,
            mismatches);

        var leftConfig = KnowledgeJson.CanonicalSerialize(left.ConfigIdentity.OrderBy(p => p.Key, StringComparer.Ordinal).ToDictionary(p => p.Key, p => p.Value));
        var rightConfig = KnowledgeJson.CanonicalSerialize(right.ConfigIdentity.OrderBy(p => p.Key, StringComparer.Ordinal).ToDictionary(p => p.Key, p => p.Value));
        if (!string.Equals(leftConfig, rightConfig, StringComparison.Ordinal))
            mismatches.Add("configIdentity");

        return mismatches.Count == 0;
    }

    /// <summary>
    /// Hashable fingerprint projection — excludes nothing volatile because fingerprint
    /// fields are themselves identity; timestamps never appear here.
    /// </summary>
    public static object ToHashable(ApplicationFingerprint fingerprint) =>
        new
        {
            applicationId = fingerprint.ApplicationId,
            executablePath = fingerprint.ExecutablePath,
            productVersion = fingerprint.ProductVersion,
            fileVersion = fingerprint.FileVersion,
            executableSha256 = fingerprint.ExecutableSha256,
            configIdentity = fingerprint.ConfigIdentity
                .OrderBy(p => p.Key, StringComparer.Ordinal)
                .ToDictionary(p => p.Key, p => p.Value, StringComparer.Ordinal),
            driverVersion = fingerprint.DriverVersion,
            snapshotSchemaVersion = fingerprint.SnapshotSchemaVersion,
            explorerVersion = fingerprint.ExplorerVersion,
            explorationPolicyHash = fingerprint.ExplorationPolicyHash
        };

    private static void Compare(string field, string? left, string? right, List<string> mismatches)
    {
        var a = string.IsNullOrWhiteSpace(left) ? null : left.Trim();
        var b = string.IsNullOrWhiteSpace(right) ? null : right.Trim();
        if (!string.Equals(a, b, StringComparison.Ordinal))
            mismatches.Add(field);
    }

    private static string? NormalizePath(string? path) =>
        string.IsNullOrWhiteSpace(path) ? null : path.Trim();

    private static string? NormalizeSha(string? sha) =>
        string.IsNullOrWhiteSpace(sha) ? null : sha.Trim().ToLowerInvariant();

    private static bool IsSecretLike(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return false;

        var v = value.Trim();
        return v.Contains("password", StringComparison.OrdinalIgnoreCase)
               || v.Contains("secret", StringComparison.OrdinalIgnoreCase)
               || v.Contains("token", StringComparison.OrdinalIgnoreCase)
               || v.Contains("bearer", StringComparison.OrdinalIgnoreCase)
               || v.Contains("apikey", StringComparison.OrdinalIgnoreCase)
               || v.Contains("api_key", StringComparison.OrdinalIgnoreCase);
    }
}
