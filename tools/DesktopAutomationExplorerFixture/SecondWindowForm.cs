namespace DesktopAutomationExplorerFixture;

internal sealed class SecondWindowForm : Form
{
    public SecondWindowForm()
    {
        Text = "DA Explorer Second";
        Name = UiIds.SecondWindow;
        Width = 360;
        Height = 200;
        StartPosition = FormStartPosition.CenterScreen;

        Controls.Add(new Label
        {
            Text = "Second top-level window.",
            Left = 16,
            Top = 24,
            Width = 300,
            Height = 40
        });

        var close = new Button
        {
            Text = "Close Second",
            Name = "ExplorerSecondClose",
            Left = 110,
            Top = 90,
            Width = 120
        };
        close.Click += (_, _) => Close();
        Controls.Add(close);
    }
}
