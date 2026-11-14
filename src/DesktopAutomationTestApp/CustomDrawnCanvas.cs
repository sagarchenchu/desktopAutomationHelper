namespace DesktopAutomationTestApp;

/// <summary>
/// Custom-drawn control with intentionally empty UIA AutomationId/Name for Phase 8 vision fixtures.
/// Painted text is not exposed as a standard AutomationId/Name pair. Find via parent
/// <see cref="UiIds.CustomDrawnGroup"/> (not via control Name — Name would break the inaccessibility rule).
/// </summary>
public sealed class CustomDrawnCanvas : Control
{
    public CustomDrawnCanvas()
    {
        // Deliberately leave WinForms Name empty so UIA AutomationId is empty.
        // Do NOT set Name = UiIds.CustomDrawnCanvas — that made the control semantically accessible.
        Name = string.Empty;
        AccessibleName = string.Empty;
        AccessibleDescription = string.Empty;
        // Client → Pane in UIA; empty AutomationId+Name + Pane/Custom = inaccessible under Phase-8 rule.
        AccessibleRole = AccessibleRole.Client;
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint, true);
        BackColor = Color.FromArgb(32, 48, 64);
        Size = new Size(280, 80);
        Cursor = Cursors.Hand;
        Tag = UiIds.CustomDrawnCanvas; // app-only marker; not exposed as UIA AutomationId
    }

    public string DrawnLabel { get; set; } = "Custom Drawn Target";

    public int ClickCount { get; private set; }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        using var brush = new SolidBrush(Color.FromArgb(220, 236, 255));
        using var font = new Font("Segoe UI", 12f, FontStyle.Bold);
        var size = e.Graphics.MeasureString(DrawnLabel, font);
        var point = new PointF((Width - size.Width) / 2f, (Height - size.Height) / 2f);
        e.Graphics.DrawString(DrawnLabel, font, brush, point);
        using var border = new Pen(Color.Goldenrod, 2);
        e.Graphics.DrawRectangle(border, 1, 1, Width - 3, Height - 3);
    }

    protected override void OnClick(EventArgs e)
    {
        ClickCount++;
        base.OnClick(e);
    }
}
