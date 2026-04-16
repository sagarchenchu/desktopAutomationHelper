namespace DesktopAutomationDriver.Services;

/// <summary>
/// A small always-on-top notification window shown in the top-right corner of the
/// primary screen when the recording session stops.  It displays "Recording stopped"
/// and the path of the exported JSON file, then auto-dismisses after a few seconds.
/// </summary>
internal sealed class RecordingStoppedNotification : Form
{
    private const int AutoCloseSecs = 6;

    public RecordingStoppedNotification(string? exportFilePath)
    {
        var screen = Screen.PrimaryScreen ?? Screen.AllScreens[0];

        const int width = 480;
        const int height = 72;
        int x = screen.Bounds.Right - width - 12;
        int y = screen.Bounds.Top + 12;

        FormBorderStyle = FormBorderStyle.None;
        TopMost = true;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.Manual;
        Opacity = 0.92;
        BackColor = Color.FromArgb(30, 30, 30);
        ForeColor = Color.White;
        Bounds = new Rectangle(x, y, width, height);

        var titleLabel = new Label
        {
            Text = "✓  Recording stopped",
            Font = new Font("Segoe UI", 10f, FontStyle.Bold),
            ForeColor = Color.LimeGreen,
            AutoSize = false,
            Bounds = new Rectangle(10, 8, width - 20, 22),
            TextAlign = ContentAlignment.MiddleLeft
        };
        Controls.Add(titleLabel);

        var pathText = string.IsNullOrEmpty(exportFilePath)
            ? "(file path unavailable)"
            : exportFilePath;

        var pathLabel = new Label
        {
            Text = pathText,
            Font = new Font("Segoe UI", 8.5f, FontStyle.Regular),
            ForeColor = Color.Silver,
            AutoSize = false,
            Bounds = new Rectangle(10, 34, width - 20, 30),
            TextAlign = ContentAlignment.MiddleLeft
        };
        Controls.Add(pathLabel);

        // Auto-close timer
        var timer = new System.Windows.Forms.Timer { Interval = AutoCloseSecs * 1000 };
        timer.Tick += (_, _) =>
        {
            timer.Stop();
            timer.Dispose();
            Close();
        };
        timer.Start();
    }

    protected override void OnShown(EventArgs e)
    {
        base.OnShown(e);

        // Make window click-through so it doesn't interfere with the user
        const int GWL_EXSTYLE = -20;
        const int WS_EX_TRANSPARENT = 0x00000020;
        const int WS_EX_LAYERED = 0x00080000;

        int style = GetWindowLong(Handle, GWL_EXSTYLE);
        SetWindowLong(Handle, GWL_EXSTYLE, style | WS_EX_TRANSPARENT | WS_EX_LAYERED);
    }

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
}
