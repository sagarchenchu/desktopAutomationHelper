namespace DesktopAutomationExplorerFixture;

internal sealed class SecondWindowForm : Form
{
    private readonly Label _status;

    public SecondWindowForm()
    {
        Text = "DA Explorer Second";
        Name = UiIds.SecondWindow;
        Width = 360;
        Height = 220;
        StartPosition = FormStartPosition.CenterScreen;

        Controls.Add(new Label
        {
            Text = "Second top-level window.",
            Left = 16,
            Top = 24,
            Width = 300,
            Height = 40
        });

        _status = new Label
        {
            Text = "Status: Ready",
            Name = UiIds.SecondStatusLabel,
            Left = 16,
            Top = 160,
            Width = 320,
            Height = 24
        };
        Controls.Add(_status);

        var ping = new Button
        {
            Text = "Ping Second",
            Name = UiIds.SecondPingButton,
            Left = 110,
            Top = 70,
            Width = 120
        };
        ping.Click += (_, _) => _status.Text = "Status: Pinged second window";
        Controls.Add(ping);

        var close = new Button
        {
            Text = "Close Second",
            Name = "ExplorerSecondClose",
            Left = 110,
            Top = 110,
            Width = 120
        };
        close.Click += (_, _) => Close();
        Controls.Add(close);
    }
}
