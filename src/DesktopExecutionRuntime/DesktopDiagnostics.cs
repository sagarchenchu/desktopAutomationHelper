namespace DesktopExecutionRuntime;

public sealed record DesktopDiagnostics(
    string SchemaVersion,
    bool PreflightPassed,
    IReadOnlyList<string> BlockingReasons,
    IReadOnlyList<string> UnverifiedCapabilities,
    DesktopSnapshot Snapshot)
{
    /// <summary>
    /// Conservative attach-mode preflight. This never proves a test can execute.
    /// No active input, screen capture, target process, or session provisioning is performed.
    /// </summary>
    public static DesktopDiagnostics Evaluate(DesktopSnapshot snapshot)
    {
        var reasons = new List<string>();
        if (!snapshot.IsWindows)
        {
            reasons.Add("UNSUPPORTED_PLATFORM");
        }
        else
        {
            if (snapshot.SessionId is null or <= 0)
                reasons.Add("INTERACTIVE_SESSION_UNAVAILABLE");
            if (snapshot.ConnectionState != "Active")
                reasons.Add("SESSION_NOT_ACTIVE");
            if (!string.Equals(snapshot.WindowStation, "WinSta0", StringComparison.OrdinalIgnoreCase)
                || snapshot.WindowStationVisible != true)
                reasons.Add("INTERACTIVE_WINDOW_STATION_UNAVAILABLE");
            if (string.IsNullOrWhiteSpace(snapshot.ThreadDesktop)
                || string.IsNullOrWhiteSpace(snapshot.InputDesktop)
                || !string.Equals(snapshot.ThreadDesktop, snapshot.InputDesktop, StringComparison.OrdinalIgnoreCase)
                || snapshot.ThreadDesktopReceivesInput != true)
                reasons.Add("INPUT_DESKTOP_UNAVAILABLE");
            if (snapshot.MonitorCount is null or <= 0
                || snapshot.VirtualScreenWidth is null or <= 0
                || snapshot.VirtualScreenHeight is null or <= 0)
                reasons.Add("DISPLAY_UNAVAILABLE");
            if (snapshot.ProbeFailures.Count > 0)
                reasons.Add("PROBE_FAILED");
        }

        return new DesktopDiagnostics("1.0", reasons.Count == 0, reasons,
            new[] { "INPUT_INJECTION", "SCREEN_CAPTURE", "TARGET_SESSION_AND_INTEGRITY", "VIEWER_INDEPENDENT_EXECUTION" },
            snapshot);
    }
}
