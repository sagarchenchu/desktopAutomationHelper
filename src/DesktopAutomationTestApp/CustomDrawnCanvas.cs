namespace DesktopAutomationTestApp;

/// <summary>
/// Custom-drawn control with intentionally weak UIA identity for Phase 8 vision/debug fixtures.
/// Painted text is not exposed as a standard AutomationId/Name pair.
/// </summary>
public sealed class CustomDrawnCanvas : Control
{
    public CustomDrawnCanvas()
    {
        Name = UiIds.CustomDrawnCanvas;
        // Deliberately leave AccessibleName empty so semantic UIA treats this as inaccessible/custom.
        AccessibleName = string.Empty;
        AccessibleRole = AccessibleRole.Client;
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint, true);
        BackColor = Color.FromArgb(32, 48, 64);
        Size = new Size(280, 80);
        Cursor = Cursors.Hand;
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
