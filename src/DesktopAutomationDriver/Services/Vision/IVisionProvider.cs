namespace DesktopAutomationDriver.Services.Vision;

public interface IVisionProvider
{
    string Name { get; }

    IReadOnlyList<VisionCandidateTarget> Analyze(VisionScreenshot screenshot, string? query = null);
}
