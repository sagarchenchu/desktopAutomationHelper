namespace DesktopAutomationDriver.Services.Debug;

/// <summary>Testable surface for the yellow non-interactive highlight rectangle.</summary>
public interface IHighlightSurface : IDisposable
{
    bool IsVisible { get; }
    Rectangle? CurrentBounds { get; }
    bool InterceptsClicks { get; }
    void Show(Rectangle screenBounds);
    void Hide();
}
