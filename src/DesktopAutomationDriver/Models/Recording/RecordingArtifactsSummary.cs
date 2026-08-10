namespace DesktopAutomationDriver.Models.Recording;

/// <summary>Optional sidecar artifact summary returned with a recording export.</summary>
public sealed class RecordingArtifactsSummary
{
    public string? Directory { get; set; }

    public string? BddActionMapFile { get; set; }

    public List<string> PageObjectFiles { get; set; } = [];

    public List<string> Warnings { get; set; } = [];
}
