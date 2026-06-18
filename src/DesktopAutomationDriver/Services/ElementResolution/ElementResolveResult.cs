using FlaUI.Core.AutomationElements;

namespace DesktopAutomationDriver.Services.ElementResolution;

public sealed class ElementResolveResult
{
    public bool Success { get; init; }
    public AutomationElement? Element { get; init; }
    public ElementSnapshot? Snapshot { get; init; }

    public string Strategy { get; init; } = "";
    public string? Error { get; init; }

    public int CandidateCount { get; init; }
    public List<ElementSnapshot> Candidates { get; init; } = new();

    public object? Criteria { get; init; }
    public object? SearchRoot { get; init; }

    public long ElapsedMs { get; init; }
    public bool Ambiguous { get; init; }
    public bool FallbackUsed { get; init; }

    public IReadOnlyList<ElementCandidate> RawCandidates { get; init; } =
        Array.Empty<ElementCandidate>();
}
