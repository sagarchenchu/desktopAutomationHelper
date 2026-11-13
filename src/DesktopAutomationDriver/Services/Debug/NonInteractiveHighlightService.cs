namespace DesktopAutomationDriver.Services.Debug;

/// <summary>
/// Yellow highlight that never intercepts clicks. Lifecycle is explicit and cleaned up on stop.
/// </summary>
public sealed class NonInteractiveHighlightService : IDisposable
{
    public static readonly Color HighlightColor = Color.Yellow;

    private readonly IHighlightSurface _surface;
    private bool _disposed;

    public NonInteractiveHighlightService(IHighlightSurface surface)
    {
        _surface = surface ?? throw new ArgumentNullException(nameof(surface));
    }

    public bool IsVisible => !_disposed && _surface.IsVisible;

    public Rectangle? CurrentBounds => _disposed ? null : _surface.CurrentBounds;

    public bool InterceptsClicks => _surface.InterceptsClicks;

    public void Highlight(Rectangle screenBounds)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        if (screenBounds.Width <= 0 || screenBounds.Height <= 0)
        {
            Clear();
            return;
        }

        _surface.Show(screenBounds);
    }

    public void Highlight(DebugUiaNode? node)
    {
        if (node?.Bounds is not { } bounds)
        {
            Clear();
            return;
        }

        Highlight(bounds);
    }

    public void Clear()
    {
        if (_disposed)
            return;
        _surface.Hide();
    }

    public void Dispose()
    {
        if (_disposed)
            return;
        _disposed = true;
        try
        {
            _surface.Hide();
        }
        finally
        {
            _surface.Dispose();
        }
    }
}

/// <summary>In-memory highlight surface for unit tests (no WinForms).</summary>
public sealed class FakeHighlightSurface : IHighlightSurface
{
    public bool IsVisible { get; private set; }
    public Rectangle? CurrentBounds { get; private set; }
    public bool InterceptsClicks => false;
    public int ShowCount { get; private set; }
    public int HideCount { get; private set; }
    public bool Disposed { get; private set; }

    public void Show(Rectangle screenBounds)
    {
        ObjectDisposedException.ThrowIf(Disposed, this);
        CurrentBounds = screenBounds;
        IsVisible = true;
        ShowCount++;
    }

    public void Hide()
    {
        if (Disposed)
            return;
        IsVisible = false;
        CurrentBounds = null;
        HideCount++;
    }

    public void Dispose()
    {
        Hide();
        Disposed = true;
    }
}

/// <summary>
/// WinForms click-through yellow border overlay. Created only on interactive desktops.
/// </summary>
public sealed class WinFormsHighlightSurface : IHighlightSurface
{
    private HighlightForm? _form;

    public bool IsVisible => _form is { Visible: true };
    public Rectangle? CurrentBounds => _form?.Bounds;
    public bool InterceptsClicks => false;

    public void Show(Rectangle screenBounds)
    {
        EnsureForm();
        _form!.Bounds = screenBounds;
        if (!_form.Visible)
            _form.Show();
        else
            _form.Invalidate();
    }

    public void Hide()
    {
        if (_form is null)
            return;
        _form.Hide();
    }

    public void Dispose()
    {
        if (_form is null)
            return;
        _form.Close();
        _form.Dispose();
        _form = null;
    }

    private void EnsureForm()
    {
        if (_form is not null)
            return;
        _form = new HighlightForm();
    }

    private sealed class HighlightForm : Form
    {
        private const int WsExTransparent = 0x00000020;
        private const int WsExToolWindow = 0x00000080;
        private const int WsExNoActivate = 0x08000000;
        private const int WsExLayered = 0x00080000;

        public HighlightForm()
        {
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            TopMost = true;
            StartPosition = FormStartPosition.Manual;
            BackColor = Color.Magenta;
            TransparencyKey = Color.Magenta;
            Opacity = 1.0;
        }

        protected override bool ShowWithoutActivation => true;

        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= WsExTransparent | WsExToolWindow | WsExNoActivate | WsExLayered;
                return cp;
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            using var pen = new Pen(NonInteractiveHighlightService.HighlightColor, 3);
            e.Graphics.DrawRectangle(pen, 1, 1, Width - 3, Height - 3);
        }
    }
}
