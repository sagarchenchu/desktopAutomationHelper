namespace DesktopAutomationExplorerFixture;

internal sealed class ChildForm : Form
{
    public ChildForm()
    {
        Text = "DA Explorer Child";
        Name = UiIds.ChildWindow;
        Width = 320;
        Height = 180;
        StartPosition = FormStartPosition.CenterParent;

        Controls.Add(new Label
        {
            Text = "Non-modal child window.",
            Left = 16,
            Top = 24,
            Width = 280,
            Height = 40
        });
    }
}
