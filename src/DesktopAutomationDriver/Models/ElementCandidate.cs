namespace DesktopAutomationDriver.Models;

/// <summary>
/// Snapshot of a candidate element captured during element resolution for diagnostics.
/// </summary>
public class ElementCandidate
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
    public string AutomationId { get; set; } = string.Empty;
    public string ClassName { get; set; } = string.Empty;
    public string ControlType { get; set; } = string.Empty;
    public string FrameworkId { get; set; } = string.Empty;
    public string RuntimeId { get; set; } = string.Empty;
    public int? ProcessId { get; set; }
    public long? Hwnd { get; set; }
    public object? Rectangle { get; set; }
    public string Value { get; set; } = string.Empty;
    public string Text { get; set; } = string.Empty;
    public bool? IsEnabled { get; set; }
    public bool? IsOffscreen { get; set; }
    public int Score { get; set; }
    public string Reason { get; set; } = string.Empty;

    // -------------------------------------------------------------------------
    // Backend separation fields
    // -------------------------------------------------------------------------
    public string Backend { get; set; } = "uia";
    public object? NativeElement { get; set; }
    public FlaUI.Core.AutomationElements.AutomationElement? UiaElement { get; set; }
    public System.IntPtr? HwndPtr { get; set; }
    public int? ControlId { get; set; }
    public bool? Enabled { get; set; }
    public bool? Visible { get; set; }
    public int CtrlIndex { get; set; }
    public int FoundIndex { get; set; }
    public double ScoreDouble { get; set; }
}
