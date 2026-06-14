using FlaUI.Core.AutomationElements;

namespace DesktopAutomationDriver.Services.Resolution;

public sealed class ElementMatchResult
{
    public AutomationElement Element { get; init; } = null!;
    public string Strategy { get; init; } = string.Empty;
    public int Score { get; init; }
    public int Index { get; init; }
    public string Reason { get; init; } = string.Empty;
    public ElementSnapshot Snapshot { get; init; } = null!;
}
