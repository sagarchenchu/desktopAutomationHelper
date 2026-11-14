namespace DesktopAutomationDriver.Services.Vision;

/// <summary>
/// Driver-owned <em>operator confirmation UX</em> for vision approval.
/// <para>
/// Phase 8 intentionally uses an interactive desktop dialog. That dialog is
/// <strong>UIA-automatable</strong> (acceptance may click Yes via UI Automation) and is
/// <strong>not</strong> OS-enforced human-presence / out-of-band identity verification.
/// Same-user ACL file self-read is also not an approval channel. A durable
/// non-UIA / separate-identity broker remains future work.
/// </para>
/// </summary>
public interface IOperatorPresenceService
{
    /// <summary>
    /// Shows or evaluates operator confirmation. Returns false on deny, cancel, timeout,
    /// cancellation, or single-flight conflict.
    /// </summary>
    bool TryConfirmVisionApproval(
        OperatorPresenceRequest request,
        out string denyReason,
        CancellationToken cancellationToken = default);
}

public sealed record OperatorPresenceRequest(
    string CandidateId,
    string Action,
    string Label);

/// <summary>
/// Interactive Yes/No operator confirmation UX (WinForms dialog).
/// Timeout + cancellation + single-flight; never blocks forever on <c>Thread.Join</c>.
/// Not a human-presence enforcement mechanism — dialog buttons are UIA-clickable.
/// </summary>
public sealed class InteractiveOperatorPresenceService : IOperatorPresenceService
{
    public const int DefaultTimeoutMilliseconds = 60_000;
    public const int MinTimeoutMilliseconds = 500;
    public const int MaxTimeoutMilliseconds = 600_000;
    public const int CleanupJoinMilliseconds = 2_000;

    private readonly TimeSpan _timeout;
    private int _inFlight;

    public InteractiveOperatorPresenceService(VisionFallbackOptions? options = null)
        : this(ResolveTimeout(options))
    {
    }

    /// <summary>Test / DI seam for an explicit timeout.</summary>
    public InteractiveOperatorPresenceService(TimeSpan timeout)
    {
        if (timeout < TimeSpan.FromMilliseconds(MinTimeoutMilliseconds))
            timeout = TimeSpan.FromMilliseconds(MinTimeoutMilliseconds);
        if (timeout > TimeSpan.FromMilliseconds(MaxTimeoutMilliseconds))
            timeout = TimeSpan.FromMilliseconds(MaxTimeoutMilliseconds);
        _timeout = timeout;
    }

    public TimeSpan Timeout => _timeout;

    public bool TryConfirmVisionApproval(
        OperatorPresenceRequest request,
        out string denyReason,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);
        denyReason = string.Empty;

        if (cancellationToken.IsCancellationRequested)
        {
            denyReason = "Operator confirmation cancelled.";
            return false;
        }

        if (Interlocked.CompareExchange(ref _inFlight, 1, 0) != 0)
        {
            denyReason = "Operator confirmation already in progress (single-flight).";
            return false;
        }

        try
        {
            return ShowConfirmation(request, out denyReason, cancellationToken);
        }
        finally
        {
            Interlocked.Exchange(ref _inFlight, 0);
        }
    }

    private bool ShowConfirmation(
        OperatorPresenceRequest request,
        out string denyReason,
        CancellationToken cancellationToken)
    {
        denyReason = string.Empty;
        var dialogResult = DialogResultBox.No;
        Exception? uiError = null;

        if (System.Windows.Forms.Application.MessageLoop)
        {
            // Already on a UI thread — ShowDialog with a timeout timer (no unbounded wait).
            using var form = CreateConfirmationForm(request);
            using var timeoutTimer = new System.Windows.Forms.Timer
            {
                Interval = (int)Math.Min(int.MaxValue, _timeout.TotalMilliseconds)
            };
            var timedOut = false;
            timeoutTimer.Tick += (_, _) =>
            {
                timedOut = true;
                timeoutTimer.Stop();
                if (!form.IsDisposed)
                {
                    form.DialogResult = System.Windows.Forms.DialogResult.Cancel;
                    form.Close();
                }
            };
            timeoutTimer.Start();

            using var cancelReg = cancellationToken.Register(() =>
            {
                if (!form.IsDisposed)
                {
                    form.BeginInvoke(new Action(() =>
                    {
                        form.DialogResult = System.Windows.Forms.DialogResult.Cancel;
                        form.Close();
                    }));
                }
            });

            var shown = form.ShowDialog();
            timeoutTimer.Stop();

            if (cancellationToken.IsCancellationRequested)
            {
                denyReason = "Operator confirmation cancelled.";
                return false;
            }

            if (timedOut || shown == System.Windows.Forms.DialogResult.Cancel)
            {
                denyReason = timedOut
                    ? "Operator confirmation timed out."
                    : "Operator confirmation cancelled.";
                return false;
            }

            dialogResult = shown == System.Windows.Forms.DialogResult.Yes
                ? DialogResultBox.Yes
                : DialogResultBox.No;
        }
        else
        {
            // Off UI thread — STA helper with bounded Join (never unlimited).
            System.Windows.Forms.Form? form = null;
            var timedOut = false;
            var cancelled = false;
            var ready = new ManualResetEventSlim(false);

            var thread = new Thread(() =>
            {
                try
                {
                    System.Windows.Forms.Application.EnableVisualStyles();
                    form = CreateConfirmationForm(request);
                    ready.Set();

                    using var timeoutTimer = new System.Windows.Forms.Timer
                    {
                        Interval = (int)Math.Min(int.MaxValue, _timeout.TotalMilliseconds)
                    };
                    timeoutTimer.Tick += (_, _) =>
                    {
                        timedOut = true;
                        timeoutTimer.Stop();
                        if (form is { IsDisposed: false })
                        {
                            form.DialogResult = System.Windows.Forms.DialogResult.Cancel;
                            form.Close();
                        }
                    };
                    timeoutTimer.Start();

                    var shown = form.ShowDialog();
                    timeoutTimer.Stop();
                    if (!timedOut && !cancelled)
                    {
                        dialogResult = shown == System.Windows.Forms.DialogResult.Yes
                            ? DialogResultBox.Yes
                            : DialogResultBox.No;
                    }
                }
                catch (Exception ex)
                {
                    uiError = ex;
                    ready.Set();
                }
            })
            {
                IsBackground = true,
                Name = "VisionOperatorConfirmation"
            };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();

            // Wait until form exists or thread failed, then wait for completion with timeout.
            if (!ready.Wait(_timeout, cancellationToken))
            {
                cancelled = cancellationToken.IsCancellationRequested;
                timedOut = !cancelled;
                TryCloseForm(form);
                thread.Join(CleanupJoinMilliseconds);
                denyReason = cancelled
                    ? "Operator confirmation cancelled."
                    : "Operator confirmation timed out.";
                return false;
            }

            if (uiError is not null)
            {
                denyReason = "Operator confirmation UI failed.";
                return false;
            }

            using var cancelReg = cancellationToken.Register(() =>
            {
                cancelled = true;
                TryCloseForm(form);
            });

            if (!thread.Join(_timeout))
            {
                timedOut = true;
                TryCloseForm(form);
                thread.Join(CleanupJoinMilliseconds);
                denyReason = "Operator confirmation timed out.";
                return false;
            }

            if (cancelled || cancellationToken.IsCancellationRequested)
            {
                denyReason = "Operator confirmation cancelled.";
                return false;
            }

            if (timedOut)
            {
                denyReason = "Operator confirmation timed out.";
                return false;
            }
        }

        if (dialogResult != DialogResultBox.Yes)
        {
            denyReason = "Operator confirmation denied or cancelled.";
            return false;
        }

        return true;
    }

    private static void TryCloseForm(System.Windows.Forms.Form? form)
    {
        if (form is null || form.IsDisposed)
            return;
        try
        {
            if (form.InvokeRequired)
            {
                form.BeginInvoke(new Action(() =>
                {
                    if (!form.IsDisposed)
                    {
                        form.DialogResult = System.Windows.Forms.DialogResult.Cancel;
                        form.Close();
                    }
                }));
            }
            else
            {
                form.DialogResult = System.Windows.Forms.DialogResult.Cancel;
                form.Close();
            }
        }
        catch
        {
            // best effort
        }
    }

    internal static System.Windows.Forms.Form CreateConfirmationForm(OperatorPresenceRequest request)
    {
        var form = new System.Windows.Forms.Form
        {
            Text = "Desktop Automation Driver — Operator confirmation",
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog,
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen,
            MinimizeBox = false,
            MaximizeBox = false,
            ShowInTaskbar = true,
            TopMost = true,
            Width = 480,
            Height = 260,
            Padding = new System.Windows.Forms.Padding(12)
        };

        var label = new System.Windows.Forms.Label
        {
            AutoSize = false,
            Dock = System.Windows.Forms.DockStyle.Top,
            Height = 140,
            Text =
                "Confirm vision coordinate action?\r\n\r\n" +
                $"Candidate: {request.CandidateId}\r\n" +
                $"Action: {request.Action}\r\n" +
                $"Label: {request.Label}\r\n\r\n" +
                "Note: this is operator confirmation UX only — the Yes button is " +
                "UIA-automatable and is not OS-enforced human presence."
        };

        var yes = new System.Windows.Forms.Button
        {
            Text = "Yes",
            DialogResult = System.Windows.Forms.DialogResult.Yes,
            Width = 96,
            Height = 28,
            Left = 250,
            Top = 170
        };
        var no = new System.Windows.Forms.Button
        {
            Text = "No",
            DialogResult = System.Windows.Forms.DialogResult.No,
            Width = 96,
            Height = 28,
            Left = 354,
            Top = 170
        };

        form.Controls.Add(label);
        form.Controls.Add(yes);
        form.Controls.Add(no);
        form.AcceptButton = yes;
        form.CancelButton = no;
        return form;
    }

    private static TimeSpan ResolveTimeout(VisionFallbackOptions? options)
    {
        var ms = options?.OperatorConfirmationTimeoutMilliseconds ?? DefaultTimeoutMilliseconds;
        if (ms < MinTimeoutMilliseconds)
            ms = MinTimeoutMilliseconds;
        if (ms > MaxTimeoutMilliseconds)
            ms = MaxTimeoutMilliseconds;
        return TimeSpan.FromMilliseconds(ms);
    }

    private enum DialogResultBox
    {
        No = 0,
        Yes = 1
    }
}

/// <summary>Test double — allow/deny without UI.</summary>
public sealed class ScriptedOperatorPresenceService : IOperatorPresenceService
{
    public bool Allow { get; set; } = true;
    public int ConfirmCalls { get; private set; }
    public TimeSpan? ArtificialDelay { get; set; }

    public bool TryConfirmVisionApproval(
        OperatorPresenceRequest request,
        out string denyReason,
        CancellationToken cancellationToken = default)
    {
        ConfirmCalls++;
        if (ArtificialDelay is { } delay)
        {
            if (cancellationToken.WaitHandle.WaitOne(delay))
            {
                denyReason = "Operator confirmation cancelled.";
                return false;
            }
        }

        if (cancellationToken.IsCancellationRequested)
        {
            denyReason = "Operator confirmation cancelled.";
            return false;
        }

        if (Allow)
        {
            denyReason = string.Empty;
            return true;
        }

        denyReason = "Operator confirmation denied (scripted).";
        return false;
    }
}
