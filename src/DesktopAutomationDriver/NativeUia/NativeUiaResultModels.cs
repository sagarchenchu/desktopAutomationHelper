using Interop.UIAutomationClient;

namespace DesktopAutomationDriver.NativeUia;

internal sealed class NativeUiaFindResult
{
    public bool Found { get; init; }
    public IUIAutomationElement? Element { get; init; }
    public string Strategy { get; init; } = "";
    public string LastError { get; init; } = "";
    public bool Ambiguous { get; init; }
    public List<NativeUiaElementSnapshot> Candidates { get; init; } = new();
}

internal sealed class NativeUiaCandidate
{
    public IUIAutomationElement Element { get; init; } = null!;
    public NativeUiaElementSnapshot Snapshot { get; init; } = null!;
    public int Score { get; init; }
    public string Reason { get; init; } = "";
}

internal sealed class NativeUiaComboBoxSelectResult
{
    public string Operation { get; init; } = "selectcomboboxuia";
    public bool Success { get; init; }
    public string RequestedValue { get; init; } = "";
    public int? RequestedIndex { get; init; }
    public string ActualValue { get; init; } = "";
    public string Strategy { get; init; } = "";
    public bool Verified { get; init; }
    public string VerificationReason { get; init; } = "";
    public NativeUiaElementSnapshot? ComboBox { get; init; }
    public NativeUiaElementSnapshot? SelectedItem { get; init; }
    public List<string> StrategiesTried { get; init; } = new();
    public List<NativeUiaElementSnapshot> CandidateItems { get; init; } = new();
    public object? Diagnostics { get; init; }
    public long ElapsedMs { get; init; }
}
