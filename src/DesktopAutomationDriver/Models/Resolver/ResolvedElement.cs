using FlaUI.Core.AutomationElements;

namespace DesktopAutomationDriver.Models.Resolver;

public sealed class ResolvedElement
{
    public AutomationElement Element { get; init; } = null!;
    public string Strategy { get; init; } = string.Empty;
    public int Score { get; init; }
    public int Index { get; init; }
    public object? Snapshot { get; init; }
}
