using FlaUI.Core.AutomationElements;

namespace DesktopAutomationDriver.Services.ElementResolution;

public sealed class ElementCandidate
{
    public AutomationElement Element { get; init; } = default!;
    public ElementSnapshot Snapshot { get; init; } = default!;

    public int Index { get; init; }
    public int Score { get; set; }

    public List<string> MatchedBy { get; init; } = new();
    public List<string> RejectedBy { get; init; } = new();

    public static ElementCandidate FromResolutionCandidate(
        Resolution.ElementCandidate source,
        int index)
    {
        return new ElementCandidate
        {
            Element = source.Element,
            Snapshot = ElementSnapshot.FromResolutionSnapshot(source.Snapshot),
            Index = index,
            Score = source.Score,
            MatchedBy = new List<string>(source.MatchReasons),
            RejectedBy = new List<string>(source.RejectReasons)
        };
    }
}
