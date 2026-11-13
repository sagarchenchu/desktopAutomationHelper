namespace DesktopAutomationDriver.Services.Debug;

/// <summary>
/// Owns Debug Mode start/stop, inspect-under-cursor orchestration, highlight, verification,
/// and candidate artifact writing. Overlay/UiService must only call this — no feature logic there.
/// </summary>
public interface IDebugSessionCoordinator
{
    bool IsActive { get; }

    DebugSessionSnapshot Snapshot { get; }

    DebugSessionSnapshot Start(string? outputDirectory = null);

    DebugSessionSnapshot Stop();

    DebugSessionSnapshot Toggle(string? outputDirectory = null);

    /// <summary>Inspect element under cursor without clicking it.</summary>
    DebugSessionSnapshot InspectUnderCursor(Point screenPoint, bool writeArtifact = false);

    DebugSessionSnapshot VerifyLastCandidate();

    HoverProbeResult ProbeHover(Point screenPoint);
}
