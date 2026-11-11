namespace DesktopAutomationAgent.Exploration;

/// <summary>
/// Tracks in-scope child/modal windows so exploration can safely close them and
/// switch back to the parent when <c>allowCloseChildWindowsForBacktrack</c> is enabled.
/// </summary>
public sealed class ExplorationWindowStack
{
    private readonly Stack<ExplorationWindowFrame> _frames = new();

    public int Count => _frames.Count;

    public void PushChild(ExplorationWindowFrame frame)
    {
        ArgumentNullException.ThrowIfNull(frame);
        _frames.Push(frame);
    }

    public bool TryPop(out ExplorationWindowFrame? frame) => _frames.TryPop(out frame);

    public ExplorationWindowFrame? Peek() => _frames.Count == 0 ? null : _frames.Peek();
}

public sealed class ExplorationWindowFrame
{
    public string? ChildTitle { get; init; }

    public long? ChildHwnd { get; init; }

    public string? ParentTitle { get; init; }

    public long? ParentHwnd { get; init; }

    public bool IsModal { get; init; }
}
