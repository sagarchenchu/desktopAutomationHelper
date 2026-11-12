namespace DesktopAutomationTestApp;

public sealed class SecondWindowForm : Form
{
    public SecondWindowForm()
    {
        Name = UiIds.SecondWindow;
        Text = "DA Second Window";
        AccessibleName = "DA Second Window";
        StartPosition = FormStartPosition.Manual;
        Location = new Point(120, 120);
        Width = 460;
        Height = 220;

        var label = new Label
        {
            Name = UiIds.SecondWindowLabel,
            AccessibleName = "Second Window Label",
            Text = "Second top-level window for switchwindow tests.",
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter
        };

        var close = new Button
        {
            Name = UiIds.SecondWindowCloseButton,
            AccessibleName = "Close Second Window",
            Text = "Close",
            Dock = DockStyle.Bottom,
            Height = 36
        };
        close.Click += (_, _) => Close();

        Controls.Add(label);
        Controls.Add(close);
    }
}
