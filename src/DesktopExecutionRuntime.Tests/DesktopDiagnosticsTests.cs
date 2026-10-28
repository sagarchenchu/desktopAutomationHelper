using System.Text.Json;
using DesktopRuntime;

namespace DesktopExecutionRuntime.Tests;

public sealed class DesktopDiagnosticsTests
{
    private static DesktopSnapshot Eligible() => new()
    {
        IsWindows = true, ProcessId = 123, SessionId = 2, ConnectionState = "Active",
        WindowStation = "WinSta0", WindowStationVisible = true,
        ThreadDesktop = "Default", InputDesktop = "Default", ThreadDesktopReceivesInput = true,
        MonitorCount = 1, VirtualScreenWidth = 1920, VirtualScreenHeight = 1080
    };

    [Fact]
    public void PassingPreflightDoesNotClaimInputCaptureOrUnattendedExecution()
    {
        var report = DesktopDiagnostics.Evaluate(Eligible());
        Assert.True(report.PreflightPassed);
        Assert.Empty(report.BlockingReasons);
        Assert.Contains("INPUT_INJECTION", report.UnverifiedCapabilities);
        Assert.Contains("SCREEN_CAPTURE", report.UnverifiedCapabilities);
        Assert.Contains("VIEWER_INDEPENDENT_EXECUTION", report.UnverifiedCapabilities);
    }

    [Theory]
    [InlineData("service", "INTERACTIVE_SESSION_UNAVAILABLE")]
    [InlineData("unknown-session", "INTERACTIVE_SESSION_UNAVAILABLE")]
    [InlineData("disconnected", "SESSION_NOT_ACTIVE")]
    [InlineData("unknown-state", "SESSION_NOT_ACTIVE")]
    [InlineData("unknown-station", "INTERACTIVE_WINDOW_STATION_UNAVAILABLE")]
    [InlineData("invisible", "INTERACTIVE_WINDOW_STATION_UNAVAILABLE")]
    [InlineData("unknown-visible", "INTERACTIVE_WINDOW_STATION_UNAVAILABLE")]
    [InlineData("locked-desktop", "INPUT_DESKTOP_UNAVAILABLE")]
    [InlineData("unknown-input", "INPUT_DESKTOP_UNAVAILABLE")]
    [InlineData("not-receiving-input", "INPUT_DESKTOP_UNAVAILABLE")]
    [InlineData("unknown-receiving-input", "INPUT_DESKTOP_UNAVAILABLE")]
    [InlineData("no-monitor", "DISPLAY_UNAVAILABLE")]
    [InlineData("unknown-monitor", "DISPLAY_UNAVAILABLE")]
    [InlineData("empty-screen", "DISPLAY_UNAVAILABLE")]
    [InlineData("unknown-screen", "DISPLAY_UNAVAILABLE")]
    [InlineData("probe-error", "PROBE_FAILED")]
    public void FailedOrUnknownObservationBlocksPreflight(string scenario, string expectedReason)
    {
        var good = Eligible();
        var snapshot = scenario switch
        {
            "service" => good with { SessionId = 0 },
            "unknown-session" => good with { SessionId = null },
            "disconnected" => good with { ConnectionState = "Disconnected" },
            "unknown-state" => good with { ConnectionState = null },
            "unknown-station" => good with { WindowStation = null },
            "invisible" => good with { WindowStationVisible = false },
            "unknown-visible" => good with { WindowStationVisible = null },
            "locked-desktop" => good with { InputDesktop = "Winlogon" },
            "unknown-input" => good with { InputDesktop = null },
            "not-receiving-input" => good with { ThreadDesktopReceivesInput = false },
            "unknown-receiving-input" => good with { ThreadDesktopReceivesInput = null },
            "no-monitor" => good with { MonitorCount = 0 },
            "unknown-monitor" => good with { MonitorCount = null },
            "empty-screen" => good with { VirtualScreenWidth = 0 },
            "unknown-screen" => good with { VirtualScreenHeight = null },
            "probe-error" => good with { ProbeFailures = new[] { new ProbeFailure("OpenInputDesktop", 5, "Denied") } },
            _ => throw new ArgumentOutOfRangeException(nameof(scenario))
        };
        var report = DesktopDiagnostics.Evaluate(snapshot);
        Assert.False(report.PreflightPassed);
        Assert.Contains(expectedReason, report.BlockingReasons);
    }

    [Fact]
    public void UnsupportedPlatformDoesNotMasqueradeAsMissingWindowsDesktop()
    {
        var report = DesktopDiagnostics.Evaluate(new DesktopSnapshot());
        Assert.False(report.PreflightPassed);
        Assert.Equal(new[] { "UNSUPPORTED_PLATFORM" }, report.BlockingReasons);
    }

    [Theory]
    [InlineData(true, 0)]
    [InlineData(false, 2)]
    public void JsonCommandReturnsStructuredEvidenceAndMeaningfulExitCode(bool eligible, int expectedExit)
    {
        var probe = new FakeProbe(eligible ? Eligible() : Eligible() with { SessionId = 0 });
        using var output = new StringWriter();
        using var error = new StringWriter();
        Assert.Equal(expectedExit, DoctorCommand.Run(new[] { "doctor", "--json" }, probe, output, error));
        using var json = JsonDocument.Parse(output.ToString());
        Assert.Equal("1.0", json.RootElement.GetProperty("schemaVersion").GetString());
        Assert.Equal(eligible, json.RootElement.GetProperty("preflightPassed").GetBoolean());
        Assert.Equal(123, json.RootElement.GetProperty("snapshot").GetProperty("processId").GetInt32());
        Assert.Equal(1, probe.Calls);
        Assert.Equal("", error.ToString());
    }

    [Theory]
    [InlineData("run")]
    [InlineData("--json")]
    [InlineData("doctor", "--force")]
    [InlineData("doctor", "--json", "--json")]
    public void UnknownCommandsCannotAccidentallyProbeOrLaunch(params string[] args)
    {
        var probe = new FakeProbe(Eligible());
        using var output = new StringWriter();
        using var error = new StringWriter();
        Assert.Equal(64, DoctorCommand.Run(args, probe, output, error));
        Assert.Equal(0, probe.Calls);
        Assert.Equal("", output.ToString());
        Assert.Contains("Usage:", error.ToString());
    }

    [Fact]
    public void NativeProbeReportsCurrentProcessWithoutRequiringAnInteractiveCiAgent()
    {
        var snapshot = new WindowsDesktopEnvironmentProbe().Capture();
        Assert.Equal(Environment.ProcessId, snapshot.ProcessId);
        Assert.Equal(OperatingSystem.IsWindows(), snapshot.IsWindows);
        if (OperatingSystem.IsWindows())
        {
            using var process = System.Diagnostics.Process.GetCurrentProcess();
            Assert.Equal(process.SessionId, snapshot.SessionId);
        }
        else Assert.Contains("UNSUPPORTED_PLATFORM", DesktopDiagnostics.Evaluate(snapshot).BlockingReasons);
        // Deliberately do not assert PreflightPassed: CI may correctly run in session 0.
    }

    private sealed class FakeProbe(DesktopSnapshot snapshot) : IDesktopEnvironmentProbe
    {
        public int Calls { get; private set; }
        public DesktopSnapshot Capture() { Calls++; return snapshot; }
    }
}
