namespace DesktopAutomationTestApp;

public sealed class SampleDialog : Form
{
    public SampleDialog()
    {
        Name = UiIds.SampleDialog;
        Text = "DA Sample Dialog";
        AccessibleName = "DA Sample Dialog";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterParent;
        MinimizeBox = false;
        MaximizeBox = false;
        ShowInTaskbar = false;
        Width = 380;
        Height = 180;

        var message = new Label
        {
            Name = UiIds.DialogMessageLabel,
            AccessibleName = "Dialog Message",
            Text = "Confirm the sample action?",
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter
        };

        var ok = new Button
        {
            Name = UiIds.DialogOkButton,
            AccessibleName = "OK",
            Text = "OK",
            DialogResult = DialogResult.OK,
            AutoSize = true
        };
        var cancel = new Button
        {
            Name = UiIds.DialogCancelButton,
            AccessibleName = "Cancel",
            Text = "Cancel",
            DialogResult = DialogResult.Cancel,
            AutoSize = true
        };

        AcceptButton = ok;
        CancelButton = cancel;

        var buttons = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            FlowDirection = FlowDirection.RightToLeft,
            Padding = new Padding(8),
            Height = 44
        };
        buttons.Controls.Add(cancel);
        buttons.Controls.Add(ok);

        Controls.Add(message);
        Controls.Add(buttons);
    }
}
