namespace DesktopExecutionRuntime;

/// <summary>Observations from the calling process only. Null means unknown, never success.</summary>
public sealed record DesktopSnapshot
{
    public DateTimeOffset CapturedAtUtc { get; init; } = DateTimeOffset.UtcNow;
    public bool IsWindows { get; init; }
    public string OperatingSystem { get; init; } = "";
    public int ProcessId { get; init; }
    public int? SessionId { get; init; }
    public string? ConnectionState { get; init; }
    public string? WindowStation { get; init; }
    public bool? WindowStationVisible { get; init; }
    public string? ThreadDesktop { get; init; }
    public string? InputDesktop { get; init; }
    public bool? ThreadDesktopReceivesInput { get; init; }
    public int? MonitorCount { get; init; }
    // GetSystemMetrics values in the caller's DPI context, not guaranteed physical pixels.
    public int? VirtualScreenWidth { get; init; }
    public int? VirtualScreenHeight { get; init; }
    public IReadOnlyList<ProbeFailure> ProbeFailures { get; init; } = Array.Empty<ProbeFailure>();
}

public sealed record ProbeFailure(string Api, int? Win32Error, string Message);

public interface IDesktopEnvironmentProbe
{
    DesktopSnapshot Capture();
}
