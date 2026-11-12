namespace DesktopAutomationTestApp;

/// <summary>
/// WinForms <see cref="Button"/> clears <see cref="ControlStyles.StandardDoubleClick"/>,
/// so <see cref="Control.MouseDoubleClick"/> is never raised. Re-enable standard click
/// styles so physical/driver double-clicks produce a distinct double-click signal.
/// </summary>
public sealed class DoubleClickAwareButton : Button
{
    public DoubleClickAwareButton()
    {
        SetStyle(ControlStyles.StandardClick | ControlStyles.StandardDoubleClick, true);
    }
}