namespace DesktopAutomationAgent.Exploration;

/// <summary>
/// Tracks in-scope child/modal windows so exploration can safely close them and
/// switch back to the parent when <c>allowCloseChildWindowsForBacktrack</c> is enabled.
/// Frames are keyed by HWND + verified PID (not title).
/// </summary>
public sealed class ExplorationWindowStack
{
    private readonly Stack<ExplorationWindowFrame> _frames = new();

    public int Count => _frames.Count;

    public void PushChild(ExplorationWindowFrame frame)
    {
        ArgumentNullException.ThrowIfNull(frame);
        if (frame.ChildHwnd <= 0)
            throw new ArgumentException("ChildHwnd is required.", nameof(frame));
        if (frame.ChildProcessId <= 0)
            throw new ArgumentException("ChildProcessId is required.", nameof(frame));

        _frames.Push(frame);
    }

    public bool TryPeek(out ExplorationWindowFrame? frame)
    {
        if (_frames.Count == 0)
        {
            frame = null;
            return false;
        }

        frame = _frames.Peek();
        return true;
    }

    public bool TryPop(out ExplorationWindowFrame? frame) => _frames.TryPop(out frame);
}

public sealed class ExplorationWindowFrame
{
    public long ChildHwnd { get; init; }

    public int ChildProcessId { get; init; }

    public string? ChildTitle { get; init; }

    public long? ParentHwnd { get; init; }

    public int? ParentProcessId { get; init; }

    public string? ParentTitle { get; init; }

    public bool IsModal { get; init; }

    /// <summary>Depth at which the child was opened; preserved across backtrack.</summary>
    public int Depth { get; init; }
}
