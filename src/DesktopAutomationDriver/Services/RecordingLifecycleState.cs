namespace DesktopAutomationDriver.Services;

/// <summary>
/// Session lifecycle distinct from export single-flight.
/// <see cref="Stopping"/> covers the gap after inactive is set and before/while export runs.
/// </summary>
internal enum RecordingLifecycleState
{
    Idle = 0,
    Active = 1,
    Stopping = 2
}
