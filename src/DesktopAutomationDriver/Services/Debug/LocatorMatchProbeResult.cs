namespace DesktopAutomationDriver.Services.Debug;

/// <summary>
/// Distinguishes a genuine live zero-match from a probe that could not run.
/// Never treat probe failure as unique via a stale in-memory capture.
/// </summary>
public sealed class LocatorMatchProbeResult
{
    public bool ProbeSucceeded { get; init; }
    public int MatchCount { get; init; }
    public string? Error { get; init; }

    public static LocatorMatchProbeResult Unavailable(string? error = null) =>
        new() { ProbeSucceeded = false, MatchCount = 0, Error = error ?? "Locator match probe unavailable." };

    public static LocatorMatchProbeResult Ok(int matchCount) =>
        new() { ProbeSucceeded = true, MatchCount = Math.Max(0, matchCount) };
}
