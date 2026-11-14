namespace DesktopAutomationDriver.Services.Debug;

/// <summary>
/// Verifies a candidate locator resolves uniquely against a provided match count.
/// Does not promote candidates. Never treats probe-unavailable as unique.
/// </summary>
public sealed class CandidateVerificationService
{
    public CandidateVerificationResult Verify(LocatorCandidate candidate, LocatorMatchProbeResult probe)
    {
        ArgumentNullException.ThrowIfNull(candidate);
        ArgumentNullException.ThrowIfNull(probe);

        if (candidate.AutoPromoted)
        {
            return new CandidateVerificationResult
            {
                Unique = false,
                Ambiguous = false,
                NotFound = true,
                ProbeUnavailable = false,
                MatchCount = 0,
                Message = "Auto-promoted candidates are rejected; Debug Mode never promotes."
            };
        }

        if (!probe.ProbeSucceeded)
        {
            return new CandidateVerificationResult
            {
                Unique = false,
                Ambiguous = false,
                NotFound = false,
                ProbeUnavailable = true,
                MatchCount = 0,
                Message = probe.Error ?? "Live locator match probe unavailable; refusing uniqueness from stale capture."
            };
        }

        return Verify(candidate, probe.MatchCount);
    }

    public CandidateVerificationResult Verify(LocatorCandidate candidate, int matchCount)
    {
        ArgumentNullException.ThrowIfNull(candidate);

        if (candidate.AutoPromoted)
        {
            return new CandidateVerificationResult
            {
                Unique = false,
                Ambiguous = false,
                NotFound = true,
                MatchCount = 0,
                Message = "Auto-promoted candidates are rejected; Debug Mode never promotes."
            };
        }

        if (matchCount <= 0)
        {
            return new CandidateVerificationResult
            {
                Unique = false,
                Ambiguous = false,
                NotFound = true,
                MatchCount = 0,
                Message = "Candidate locator matched no elements."
            };
        }

        if (matchCount > 1)
        {
            return new CandidateVerificationResult
            {
                Unique = false,
                Ambiguous = true,
                NotFound = false,
                MatchCount = matchCount,
                Message = $"Candidate locator is ambiguous ({matchCount} matches)."
            };
        }

        return new CandidateVerificationResult
        {
            Unique = true,
            Ambiguous = false,
            NotFound = false,
            MatchCount = 1,
            Message = "Candidate locator resolves uniquely."
        };
    }

    /// <summary>
    /// Counts matches in an in-memory node set using exact Phase-3-style field equality.
    /// </summary>
    public int CountMatches(
        IReadOnlyDictionary<string, object?> locator,
        IEnumerable<DebugUiaNode> nodes)
    {
        ArgumentNullException.ThrowIfNull(locator);
        ArgumentNullException.ThrowIfNull(nodes);

        return nodes.Count(node => Matches(locator, node));
    }

    public static bool Matches(IReadOnlyDictionary<string, object?> locator, DebugUiaNode node)
    {
        if (locator.TryGetValue("automationId", out var a) && a is string automationId)
        {
            if (!string.Equals(Trim(node.AutomationId), Trim(automationId), StringComparison.OrdinalIgnoreCase))
                return false;
        }

        if (locator.TryGetValue("name", out var n) && n is string name)
        {
            if (!string.Equals(Trim(node.Name), Trim(name), StringComparison.OrdinalIgnoreCase))
                return false;
        }

        if (locator.TryGetValue("className", out var c) && c is string className)
        {
            if (!string.Equals(Trim(node.ClassName), Trim(className), StringComparison.OrdinalIgnoreCase))
                return false;
        }

        if (locator.TryGetValue("controlType", out var t) && t is string controlType)
        {
            if (!string.Equals(Trim(node.ControlType), Trim(controlType), StringComparison.OrdinalIgnoreCase))
                return false;
        }

        return locator.Count > 0;
    }

    private static string Trim(string? value) => (value ?? string.Empty).Trim();
}
