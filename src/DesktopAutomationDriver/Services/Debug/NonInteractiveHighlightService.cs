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
/// All UI work runs on a dedicated STA thread so ASP.NET/FlaUI request threads do not hang.
/// </summary>
public sealed class WinFormsHighlightSurface : IHighlightSurface
{
    private readonly object _gate = new();
    private Thread? _uiThread;
    private volatile HighlightForm? _form;
    private volatile bool _disposed;
    private readonly ManualResetEventSlim _ready = new(false);

    public bool IsVisible
    {
        get
        {
            var form = _form;
            if (form is null)
                return false;
            try
            {
                return Invoke(form, () => form.Visible);
            }
            catch
            {
                return false;
            }
        }
    }

    public Rectangle? CurrentBounds
    {
        get
        {
            var form = _form;
            if (form is null)
                return null;
            try
            {
                return Invoke(form, () => form.Bounds);
            }
            catch
            {
                return null;
            }
        }
    }

    public bool InterceptsClicks => false;

    public void Show(Rectangle screenBounds)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        EnsureUiThread();
        var form = _form ?? throw new InvalidOperationException("Highlight UI thread failed to start.");
        Invoke(form, () =>
        {
            form.Bounds = screenBounds;
            if (!form.Visible)
                form.Show();
            else
                form.Invalidate();
        });
    }

    public void Hide()
    {
        if (_disposed)
            return;
        var form = _form;
        if (form is null)
            return;
        try
        {
            Invoke(form, () => form.Hide());
        }
        catch
        {
            // Ignore teardown races.
        }
    }

    public void Dispose()
    {
        if (_disposed)
            return;
        _disposed = true;
        var form = _form;
        if (form is not null)
        {
            try
            {
                Invoke(form, () =>
                {
                    form.Close();
                    form.Dispose();
                });
            }
            catch
            {
                // Ignore teardown races.
            }
        }

        _form = null;
        _ready.Dispose();
    }

    private void EnsureUiThread()
    {
        lock (_gate)
        {
            if (_uiThread is { IsAlive: true } && _form is not null)
                return;

            _ready.Reset();
            _uiThread = new Thread(() =>
            {
                try
                {
                    var form = new HighlightForm();
                    _form = form;
                    _ready.Set();
                    Application.Run();
                }
                catch
                {
                    _ready.Set();
                }
            })
            {
                IsBackground = true,
                Name = "Phase8-HighlightSTA"
            };
            _uiThread.SetApartmentState(ApartmentState.STA);
            _uiThread.Start();
        }

        if (!_ready.Wait(TimeSpan.FromSeconds(5)) || _form is null)
            throw new InvalidOperationException("Timed out starting Phase 8 highlight STA thread.");
    }

    private static T Invoke<T>(HighlightForm form, Func<T> action)
    {
        if (!form.IsHandleCreated || !form.InvokeRequired)
            return action();
        return (T)form.Invoke(action)!;
    }

    private static void Invoke(HighlightForm form, Action action)
    {
        if (!form.IsHandleCreated || !form.InvokeRequired)
        {
            action();
            return;
        }

        form.Invoke(action);
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
            // Create HWND on the STA thread before Application.Run so callers can Invoke.
            CreateHandle();
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

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);
            try { Application.ExitThread(); } catch { /* ignore */ }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            using var pen = new Pen(NonInteractiveHighlightService.HighlightColor, 3);
            e.Graphics.DrawRectangle(pen, 1, 1, Width - 3, Height - 3);
        }
    }
}
