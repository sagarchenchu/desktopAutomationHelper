namespace DesktopAutomationTestApp;

public sealed class ChildForm : Form
{
    public ChildForm()
    {
        Name = UiIds.ChildWindow;
        Text = "DA Child Window";
        AccessibleName = "DA Child Window";
        StartPosition = FormStartPosition.CenterParent;
        Width = 420;
        Height = 240;

        var message = new TextBox
        {
            Name = UiIds.ChildMessageEdit,
            AccessibleName = "Child Message",
            Dock = DockStyle.Top,
            Height = 28,
            Margin = new Padding(8)
        };

        var save = new Button
        {
            Name = UiIds.ChildSaveButton,
            AccessibleName = "Save Child",
            Text = "Save",
            AutoSize = true
        };
        save.Click += (_, _) =>
        {
            message.Text = $"Saved at {DateTime.Now:HH:mm:ss}";
        };

        var close = new Button
        {
            Name = UiIds.ChildCloseButton,
            AccessibleName = "Close Child",
            Text = "Close",
            AutoSize = true
        };
        close.Click += (_, _) => Close();

        var buttons = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 44,
            FlowDirection = FlowDirection.RightToLeft,
            Padding = new Padding(8)
        };
        buttons.Controls.Add(close);
        buttons.Controls.Add(save);

        var help = new Label
        {
            Dock = DockStyle.Fill,
            Text = "Use this window to exercise switchwindow / closewindow and multi-window Assistive capture.",
            Padding = new Padding(12)
        };

        Controls.Add(help);
        Controls.Add(buttons);
        Controls.Add(message);
    }
}
