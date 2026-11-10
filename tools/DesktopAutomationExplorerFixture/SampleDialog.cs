namespace DesktopAutomationExplorerFixture;

internal sealed class SampleDialog : Form
{
    public SampleDialog()
    {
        Text = "DA Explorer Dialog";
        Name = UiIds.DialogWindow;
        Width = 360;
        Height = 180;
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;

        Controls.Add(new Label
        {
            Text = "Modal child dialog for explorer scope tests.",
            Left = 16,
            Top = 24,
            Width = 300,
            Height = 40
        });

        var close = new Button
        {
            Text = "Close",
            Name = "ExplorerDialogClose",
            Left = 120,
            Top = 90,
            Width = 100,
            DialogResult = DialogResult.OK
        };
        Controls.Add(close);
        AcceptButton = close;
    }
}
