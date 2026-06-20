using System.Diagnostics;
using System.Drawing;
using System.Net;
using DesktopAutomationDriver.Models.Request;
using Interop.UIAutomationClient;
using WinForms = System.Windows.Forms;

namespace DesktopAutomationDriver.Services.NativeUia;

/// <summary>
/// pywinauto-style native UIA basic operations (click, type, sendkeys, clear, focus).
/// Bounded by request timeoutMs; returns JSON instead of throwing to the controller.
/// </summary>
internal sealed class NativeUiaBasicOperationService : INativeUiaBasicOperationService
{
    private const int ListItemControlTypeId = 50007;
    private const int TabItemControlTypeId = 50019;
    private const int TabControlTypeId = 50018;
    private const int MenuItemControlTypeId = 50010;
    private const int MenuControlTypeId = 50011;
    private const int CheckBoxControlTypeId = 50002;
    private const int RadioButtonControlTypeId = 50013;
    private const int DefaultTimeoutMs = 5000;
    private const int MaxTimeoutMs = 15000;
    private const int DefaultWaitTimeoutMs = 10000;
    private const int MaxWaitTimeoutMs = 60000;
    private const int DefaultPollIntervalMs = 200;
    private const int ActionDelayMs = 80;

    private readonly NativeUiaAutomation _uia;
    private readonly NativeUiaElementResolver _resolver;
    private readonly NativeUiaTextReader _textReader;
    private readonly ILogger<NativeUiaBasicOperationService> _logger;

    public NativeUiaBasicOperationService(ILogger<NativeUiaBasicOperationService> logger)
        : this(new NativeUiaAutomation(), logger)
    {
    }

    internal NativeUiaBasicOperationService(NativeUiaAutomation uia, ILogger<NativeUiaBasicOperationService> logger)
    {
        _uia = uia;
        _textReader = new NativeUiaTextReader(_uia);
        _resolver = new NativeUiaElementResolver(_uia);
        _logger = logger;
    }

    public object Click(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default) =>
        ExecuteOperation("clickuia", request, activeWindowHwnd, processId, cancellationToken, ClickElement);

    public object ClickMenu(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default) =>
        ExecuteOperation("clickmenuuia", request, activeWindowHwnd, processId, cancellationToken, ClickElement);

    public object Type(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default) =>
        ExecuteOperation("typeuia", request, activeWindowHwnd, processId, cancellationToken, TypeElement);

    public object SendKeys(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default) =>
        ExecuteOperation("sendkeysuia", request, activeWindowHwnd, processId, cancellationToken, SendKeysElement);

    public object Clear(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default) =>
        ExecuteOperation("clearuia", request, activeWindowHwnd, processId, cancellationToken, ClearElement);

    public object Focus(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default) =>
        ExecuteOperation("focusuia", request, activeWindowHwnd, processId, cancellationToken, FocusElement);

    public object DoubleClick(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default) =>
        ExecuteOperation("doubleclickuia", request, activeWindowHwnd, processId, cancellationToken, DoubleClickElement);

    public object RightClick(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default) =>
        ExecuteOperation("rightclickuia", request, activeWindowHwnd, processId, cancellationToken, RightClickElement);

    public object Check(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default) =>
        ExecuteOperation("checkuia", request, activeWindowHwnd, processId, cancellationToken,
            (element, req, deadline, ct) => SetToggleElement(element, deadline, ct, wantChecked: true));

    public object Uncheck(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default) =>
        ExecuteOperation("uncheckuia", request, activeWindowHwnd, processId, cancellationToken,
            (element, req, deadline, ct) => SetToggleElement(element, deadline, ct, wantChecked: false));

    public object SelectTab(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default) =>
        ExecuteSelectTabOperation(request, activeWindowHwnd, processId, cancellationToken);

    public object ScreenshotElement(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default) =>
        ExecuteScreenshotElementOperation(request, activeWindowHwnd, processId, cancellationToken);

    public object Exists(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default)
    {
        var sw = Stopwatch.StartNew();
        var timeoutMs = ResolveTimeoutMs(request.TimeoutMs);
        var deadlineUtc = DateTime.UtcNow.AddMilliseconds(timeoutMs);

        return ExecuteQueryOperation(
            "existsuia",
            request,
            activeWindowHwnd,
            processId,
            deadlineUtc,
            cancellationToken,
            sw,
            (element, stage) => new
            {
                operation = "existsuia",
                success = true,
                exists = true,
                reason = (string?)null,
                strategy = "element-resolved",
                stage,
                elapsedMs = sw.ElapsedMilliseconds
            },
            onNotFound: (_, stage, view, resolveResult, elapsedMs) => new
            {
                operation = "existsuia",
                success = false,
                exists = false,
                reason = resolveResult.Stage ?? "element-not-found",
                stage,
                view,
                elapsedMs,
                candidates = resolveResult.Candidates,
                message = resolveResult.LastError,
                diagnostics = resolveResult.Diagnostics
            });
    }

    public object GetValue(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default)
    {
        var sw = Stopwatch.StartNew();
        var timeoutMs = ResolveTimeoutMs(request.TimeoutMs);
        var deadlineUtc = DateTime.UtcNow.AddMilliseconds(timeoutMs);

        return ExecuteQueryOperation(
            "getvalueuia",
            request,
            activeWindowHwnd,
            processId,
            deadlineUtc,
            cancellationToken,
            sw,
            (element, stage) =>
            {
                var (readSuccess, strategy, value) = ReadElementValue(element);
                if (!readSuccess)
                {
                    return new
                    {
                        operation = "getvalueuia",
                        success = false,
                        reason = "value-unavailable",
                        stage,
                        elapsedMs = sw.ElapsedMilliseconds,
                        message = "Could not read a value from the resolved element."
                    };
                }

                return new
                {
                    operation = "getvalueuia",
                    success = true,
                    reason = (string?)null,
                    strategy,
                    stage,
                    value,
                    elapsedMs = sw.ElapsedMilliseconds
                };
            },
            onNotFound: (_, stage, view, resolveResult, elapsedMs) => new
            {
                operation = "getvalueuia",
                success = false,
                reason = resolveResult.Stage ?? "element-not-found",
                stage,
                view,
                elapsedMs,
                candidates = resolveResult.Candidates,
                message = resolveResult.LastError,
                diagnostics = resolveResult.Diagnostics
            });
    }

    public object Wait(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default)
    {
        var sw = Stopwatch.StartNew();
        var timeoutMs = ResolveWaitTimeoutMs(request.TimeoutMs);
        var pollIntervalMs = Math.Clamp(request.PollIntervalMs ?? DefaultPollIntervalMs, 50, 2000);
        var deadlineUtc = DateTime.UtcNow.AddMilliseconds(timeoutMs);

        _logger.LogInformation(
            "NativeUiaBasic waituia starting. rootHwnd={RootHwnd}, processId={ProcessId}, timeoutMs={TimeoutMs}, pollIntervalMs={PollIntervalMs}",
            activeWindowHwnd,
            processId,
            timeoutMs,
            pollIntervalMs);

        try
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (!HasLocator(request))
            {
                return BuildFailure(
                    "waituia",
                    "invalid-request",
                    "Locator is required for waituia.",
                    sw.ElapsedMilliseconds);
            }

            if (string.IsNullOrWhiteSpace(request.Operation))
                request.Operation = "waituia";

            var attempts = 0;
            NativeUiaResolveResult? lastResolveResult = null;
            string? lastStage = null;
            string? lastView = null;

            while (DateTime.UtcNow <= deadlineUtc)
            {
                cancellationToken.ThrowIfCancellationRequested();
                attempts++;

                var resolveView = request.View
                                  ?? request.TreeView
                                  ?? NativeUiaElementResolver.InferDefaultView(request);

                var resolveResult = _resolver.ResolveElement(
                    request,
                    activeWindowHwnd,
                    processId,
                    deadlineUtc,
                    cancellationToken);

                lastResolveResult = resolveResult;
                lastStage = resolveResult.Stage;
                lastView = resolveView;

                if (resolveResult.Element != null && !resolveResult.IsAmbiguous)
                {
                    return new
                    {
                        operation = "waituia",
                        success = true,
                        reason = (string?)null,
                        strategy = "wait-until-found",
                        stage = lastStage,
                        view = lastView,
                        attempts,
                        elapsedMs = sw.ElapsedMilliseconds
                    };
                }

                if (resolveResult.IsAmbiguous)
                {
                    return new
                    {
                        operation = "waituia",
                        success = false,
                        reason = "ambiguous-element",
                        stage = lastStage,
                        view = lastView,
                        attempts,
                        elapsedMs = sw.ElapsedMilliseconds,
                        candidates = resolveResult.Candidates,
                        message = resolveResult.LastError,
                        diagnostics = resolveResult.Diagnostics
                    };
                }

                if (DateTime.UtcNow.AddMilliseconds(pollIntervalMs) > deadlineUtc)
                    break;

                WaitWithCancellation(pollIntervalMs, cancellationToken);
            }

            return new
            {
                operation = "waituia",
                success = false,
                reason = "wait-timeout",
                stage = lastStage,
                view = lastView,
                attempts,
                elapsedMs = sw.ElapsedMilliseconds,
                message = lastResolveResult?.LastError
                    ?? "Native UIA resolver did not find the element before the wait timeout elapsed.",
                candidates = lastResolveResult?.Candidates,
                diagnostics = lastResolveResult?.Diagnostics
            };
        }
        catch (TimeoutException ex)
        {
            return BuildFailure("waituia", "timeout", ex.Message, sw.ElapsedMilliseconds);
        }
        catch (OperationCanceledException)
        {
            return BuildFailure("waituia", "cancelled", "waituia was cancelled.", sw.ElapsedMilliseconds);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "NativeUiaBasic waituia failed with exception.");

            return new
            {
                operation = "waituia",
                success = false,
                reason = "error",
                elapsedMs = sw.ElapsedMilliseconds,
                exceptionType = ex.GetType().Name,
                message = ex.Message
            };
        }
    }

    private object ExecuteQueryOperation(
        string operation,
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        DateTime deadlineUtc,
        CancellationToken cancellationToken,
        Stopwatch sw,
        Func<IUIAutomationElement, string?, object> onFound,
        Func<string, string?, string?, NativeUiaResolveResult, long, object> onNotFound)
    {
        _logger.LogInformation(
            "NativeUiaBasic {Operation} starting. rootHwnd={RootHwnd}, processId={ProcessId}, timeoutMs={TimeoutMs}",
            operation,
            activeWindowHwnd,
            processId,
            ResolveTimeoutMs(request.TimeoutMs));

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            EnsureWithinDeadline(deadlineUtc, cancellationToken, "start");

            if (!HasLocator(request))
            {
                return BuildFailure(
                    operation,
                    "invalid-request",
                    $"Locator is required for {operation}.",
                    sw.ElapsedMilliseconds);
            }

            if (string.IsNullOrWhiteSpace(request.Operation))
                request.Operation = operation;

            var resolveOutcome = TryResolveLocatedElement(
                operation,
                request,
                activeWindowHwnd,
                processId,
                deadlineUtc,
                cancellationToken,
                sw.ElapsedMilliseconds,
                requireEnabled: false);

            if (resolveOutcome.FailureResponse != null)
            {
                if (resolveOutcome.ResolveResult != null
                    && resolveOutcome.View != null
                    && resolveOutcome.Stage != null
                    && resolveOutcome.ResolveResult.Element == null
                    && !resolveOutcome.ResolveResult.IsAmbiguous)
                {
                    return onNotFound(
                        operation,
                        resolveOutcome.Stage,
                        resolveOutcome.View,
                        resolveOutcome.ResolveResult,
                        sw.ElapsedMilliseconds);
                }

                return resolveOutcome.FailureResponse;
            }

            return onFound(resolveOutcome.Element!, resolveOutcome.Stage);
        }
        catch (TimeoutException ex)
        {
            return BuildFailure(operation, "timeout", ex.Message, sw.ElapsedMilliseconds);
        }
        catch (OperationCanceledException)
        {
            return BuildFailure(operation, "cancelled", $"{operation} was cancelled.", sw.ElapsedMilliseconds);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "NativeUiaBasic {Operation} failed with exception.", operation);

            return new
            {
                operation,
                success = false,
                reason = "error",
                elapsedMs = sw.ElapsedMilliseconds,
                exceptionType = ex.GetType().Name,
                message = ex.Message
            };
        }
    }

    private sealed class ResolveLocatedElementOutcome
    {
        public IUIAutomationElement? Element { get; init; }
        public string? Stage { get; init; }
        public string? View { get; init; }
        public NativeUiaResolveResult? ResolveResult { get; init; }
        public object? FailureResponse { get; init; }
    }

    private ResolveLocatedElementOutcome TryResolveLocatedElement(
        string operation,
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        DateTime deadlineUtc,
        CancellationToken cancellationToken,
        long elapsedMs,
        bool requireEnabled)
    {
        var resolveView = request.View
                          ?? request.TreeView
                          ?? NativeUiaElementResolver.InferDefaultView(request);

        var resolveResult = _resolver.ResolveElement(
            request,
            activeWindowHwnd,
            processId,
            deadlineUtc,
            cancellationToken);

        var resolverStage = resolveResult.Stage;

        _logger.LogInformation(
            "NativeUiaBasic {Operation} resolver stage={Stage}, view={View}, elapsedMs={ElapsedMs}",
            operation,
            resolverStage,
            resolveView,
            elapsedMs);

        if (resolveResult.Element == null)
        {
            return new ResolveLocatedElementOutcome
            {
                Stage = resolverStage,
                View = resolveView,
                ResolveResult = resolveResult,
                FailureResponse = new
                {
                    operation,
                    success = false,
                    reason = resolveResult.Stage ?? "element-not-found",
                    stage = resolverStage,
                    view = resolveView,
                    elapsedMs,
                    candidates = resolveResult.Candidates,
                    message = resolveResult.LastError,
                    diagnostics = resolveResult.Diagnostics
                }
            };
        }

        if (resolveResult.IsAmbiguous)
        {
            return new ResolveLocatedElementOutcome
            {
                Stage = resolverStage,
                View = resolveView,
                ResolveResult = resolveResult,
                FailureResponse = new
                {
                    operation,
                    success = false,
                    reason = "ambiguous-element",
                    stage = resolverStage,
                    view = resolveView,
                    elapsedMs,
                    candidates = resolveResult.Candidates,
                    message = resolveResult.LastError,
                    diagnostics = resolveResult.Diagnostics
                }
            };
        }

        if (requireEnabled
            && !_uia.GetBoolProperty(resolveResult.Element, UIA_PropertyIds.UIA_IsEnabledPropertyId, true))
        {
            return new ResolveLocatedElementOutcome
            {
                Stage = resolverStage,
                View = resolveView,
                ResolveResult = resolveResult,
                FailureResponse = BuildFailure(
                    operation,
                    "element-disabled",
                    "Target element is disabled.",
                    elapsedMs)
            };
        }

        return new ResolveLocatedElementOutcome
        {
            Element = resolveResult.Element,
            Stage = resolverStage,
            View = resolveView,
            ResolveResult = resolveResult
        };
    }

    private (bool success, string strategy, string? value) ReadElementValue(IUIAutomationElement element)
    {
        var valuePatternText = _uia.GetValuePatternText(element);
        if (!string.IsNullOrWhiteSpace(valuePatternText))
            return (true, "valuepattern", valuePatternText);

        var text = _textReader.ReadElementText(element);
        if (!string.IsNullOrWhiteSpace(text))
            return (true, "text-reader", text);

        var snapshot = _uia.CreateSnapshot(element);
        if (!string.IsNullOrWhiteSpace(snapshot.Name))
            return (true, "name", snapshot.Name);

        return (false, "value-unavailable", null);
    }

    private static void WaitWithCancellation(int pollIntervalMs, CancellationToken cancellationToken)
    {
        if (pollIntervalMs <= 0)
            return;

        cancellationToken.WaitHandle.WaitOne(pollIntervalMs);
        cancellationToken.ThrowIfCancellationRequested();
    }

    private static int ResolveWaitTimeoutMs(int? requestTimeoutMs)
    {
        var timeout = requestTimeoutMs ?? DefaultWaitTimeoutMs;
        return Math.Clamp(timeout, 500, MaxWaitTimeoutMs);
    }

    private object ExecuteOperation(
        string operation,
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken,
        Func<IUIAutomationElement, UiRequest, DateTime, CancellationToken, (bool success, string? strategy, string? reason, List<string>? attemptedStrategies, string? actual)> action)
    {
        var sw = Stopwatch.StartNew();
        var timeoutMs = ResolveTimeoutMs(request.TimeoutMs);
        var deadlineUtc = DateTime.UtcNow.AddMilliseconds(timeoutMs);

        _logger.LogInformation(
            "NativeUiaBasic {Operation} starting. rootHwnd={RootHwnd}, processId={ProcessId}, timeoutMs={TimeoutMs}",
            operation,
            activeWindowHwnd,
            processId,
            timeoutMs);

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            EnsureWithinDeadline(deadlineUtc, cancellationToken, "start");

            if (operation != "sendkeysuia" && !HasLocator(request))
            {
                return BuildFailure(
                    operation,
                    "invalid-request",
                    "Locator is required for this operation.",
                    sw.ElapsedMilliseconds);
            }

            if (operation == "typeuia" && string.IsNullOrWhiteSpace(request.Value))
            {
                return BuildFailure(
                    operation,
                    "invalid-request",
                    "Value is required for typeuia.",
                    sw.ElapsedMilliseconds);
            }

            if (operation == "sendkeysuia" && string.IsNullOrWhiteSpace(request.Value))
            {
                return BuildFailure(
                    operation,
                    "invalid-request",
                    "Value is required for sendkeysuia.",
                    sw.ElapsedMilliseconds);
            }

            IUIAutomationElement? element = null;
            string? resolverStage = null;

            if (HasLocator(request))
            {
                if (string.IsNullOrWhiteSpace(request.Operation))
                    request.Operation = operation;

                var resolveOutcome = TryResolveLocatedElement(
                    operation,
                    request,
                    activeWindowHwnd,
                    processId,
                    deadlineUtc,
                    cancellationToken,
                    sw.ElapsedMilliseconds,
                    requireEnabled: true);

                if (resolveOutcome.FailureResponse != null)
                    return resolveOutcome.FailureResponse;

                element = resolveOutcome.Element;
                resolverStage = resolveOutcome.Stage;
            }

            EnsureWithinDeadline(deadlineUtc, cancellationToken, "action");

            if (element == null && operation == "sendkeysuia")
            {
                SendKeysValue(request.Value!);
                _logger.LogInformation(
                    "NativeUiaBasic {Operation} succeeded. strategy=global-sendkeys, elapsedMs={ElapsedMs}",
                    operation,
                    sw.ElapsedMilliseconds);

                return new
                {
                    operation,
                    success = true,
                    strategy = "global-sendkeys",
                    elapsedMs = sw.ElapsedMilliseconds
                };
            }

            var (success, strategy, reason, attemptedStrategies, actual) =
                action(element!, request, deadlineUtc, cancellationToken);

            if (success)
            {
                _logger.LogInformation(
                    "NativeUiaBasic {Operation} succeeded. strategy={Strategy}, elapsedMs={ElapsedMs}",
                    operation,
                    strategy,
                    sw.ElapsedMilliseconds);

                return actual == null
                    ? new
                    {
                        operation,
                        success = true,
                        strategy,
                        stage = resolverStage,
                        elapsedMs = sw.ElapsedMilliseconds
                    }
                    : new
                    {
                        operation,
                        success = true,
                        strategy,
                        actual,
                        stage = resolverStage,
                        elapsedMs = sw.ElapsedMilliseconds
                    };
            }

            _logger.LogWarning(
                "NativeUiaBasic {Operation} failed. reason={Reason}, elapsedMs={ElapsedMs}",
                operation,
                reason,
                sw.ElapsedMilliseconds);

            return new
            {
                operation,
                success = false,
                reason = reason ?? "action-failed",
                attemptedStrategies,
                stage = resolverStage,
                elapsedMs = sw.ElapsedMilliseconds
            };
        }
        catch (TimeoutException ex)
        {
            _logger.LogWarning(
                ex,
                "NativeUiaBasic {Operation} timed out. elapsedMs={ElapsedMs}",
                operation,
                sw.ElapsedMilliseconds);

            return BuildFailure(operation, "timeout", ex.Message, sw.ElapsedMilliseconds);
        }
        catch (OperationCanceledException)
        {
            return BuildFailure(operation, "cancelled", $"{operation} was cancelled.", sw.ElapsedMilliseconds);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "NativeUiaBasic {Operation} failed with exception.", operation);

            return new
            {
                operation,
                success = false,
                reason = "error",
                elapsedMs = sw.ElapsedMilliseconds,
                exceptionType = ex.GetType().Name,
                message = ex.Message
            };
        }
    }

    private (bool success, string? strategy, string? reason, List<string>? attemptedStrategies, string? actual) ClickElement(
        IUIAutomationElement element,
        UiRequest request,
        DateTime deadlineUtc,
        CancellationToken cancellationToken)
    {
        var attempted = new List<string>();

        cancellationToken.ThrowIfCancellationRequested();
        EnsureWithinDeadline(deadlineUtc, cancellationToken, "click");

        if (_uia.TryGetInvokePattern(element, out var invoke))
        {
            attempted.Add("invoke-pattern");
            try
            {
                invoke!.Invoke();
                Thread.Sleep(ActionDelayMs);
                return (true, "invoke-pattern", null, attempted, null);
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "InvokePattern.Invoke failed.");
            }
        }

        var controlType = _uia.GetIntProperty(element, UIA_PropertyIds.UIA_ControlTypePropertyId);

        if (controlType is MenuItemControlTypeId or MenuControlTypeId
            && _uia.TryGetExpandCollapsePattern(element, out var expandCollapse))
        {
            attempted.Add("expandcollapse-expand");
            try
            {
                expandCollapse!.Expand();
                Thread.Sleep(ActionDelayMs);
                return (true, "expandcollapse-expand", null, attempted, null);
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "ExpandCollapse.Expand failed.");
            }
        }

        if ((controlType is ListItemControlTypeId
                or TabItemControlTypeId
                or MenuItemControlTypeId
                or MenuControlTypeId)
            && _uia.TryGetSelectionItemPattern(element, out var selectionItem))
        {
            attempted.Add("selectionitem-select");
            try
            {
                selectionItem!.Select();
                Thread.Sleep(ActionDelayMs);
                return (true, "selectionitem-select", null, attempted, null);
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "SelectionItem.Select failed.");
            }
        }

        if (controlType is CheckBoxControlTypeId or RadioButtonControlTypeId
            && _uia.TryGetTogglePattern(element, out var toggle))
        {
            attempted.Add("toggle-pattern");
            try
            {
                toggle!.Toggle();
                Thread.Sleep(ActionDelayMs);
                return (true, "toggle-pattern", null, attempted, null);
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "TogglePattern.Toggle failed.");
            }
        }

        attempted.Add("focus-enter");
        try
        {
            element.SetFocus();
            Thread.Sleep(40);
            NativeUiaInput.SendKey(NativeUiaInput.VirtualKeys.Return);
            Thread.Sleep(ActionDelayMs);
            return (true, "focus-enter", null, attempted, null);
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "SetFocus+Enter failed.");
        }

        attempted.Add("physical-click");
        var rect = _uia.GetBoundingRectangle(element);
        if (rect.HasValue)
        {
            var center = new Point(rect.Value.Left + rect.Value.Width / 2, rect.Value.Top + rect.Value.Height / 2);
            if (NativeUiaInput.ClickPoint(center))
            {
                Thread.Sleep(ActionDelayMs);
                return (true, "physical-click", null, attempted, null);
            }
        }

        return (false, null, "action-failed", attempted, null);
    }

    private (bool success, string? strategy, string? reason, List<string>? attemptedStrategies, string? actual) DoubleClickElement(
        IUIAutomationElement element,
        UiRequest request,
        DateTime deadlineUtc,
        CancellationToken cancellationToken)
    {
        var attempted = new List<string>();
        cancellationToken.ThrowIfCancellationRequested();
        EnsureWithinDeadline(deadlineUtc, cancellationToken, "doubleclick");

        if (_uia.TryGetInvokePattern(element, out var invoke))
        {
            attempted.Add("invoke-pattern");
            try
            {
                invoke!.Invoke();
                Thread.Sleep(ActionDelayMs);
                return (true, "invoke-pattern", null, attempted, null);
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "InvokePattern.Invoke failed for doubleclickuia.");
            }
        }

        attempted.Add("physical-doubleclick");
        var rect = _uia.GetBoundingRectangle(element);
        if (rect.HasValue)
        {
            var center = new Point(rect.Value.Left + rect.Value.Width / 2, rect.Value.Top + rect.Value.Height / 2);
            if (NativeUiaInput.DoubleClickPoint(center))
            {
                Thread.Sleep(ActionDelayMs);
                return (true, "physical-doubleclick", null, attempted, null);
            }
        }

        return (false, null, "action-failed", attempted, null);
    }

    private (bool success, string? strategy, string? reason, List<string>? attemptedStrategies, string? actual) RightClickElement(
        IUIAutomationElement element,
        UiRequest request,
        DateTime deadlineUtc,
        CancellationToken cancellationToken)
    {
        var attempted = new List<string>();
        cancellationToken.ThrowIfCancellationRequested();
        EnsureWithinDeadline(deadlineUtc, cancellationToken, "rightclick");

        attempted.Add("physical-rightclick");
        var rect = _uia.GetBoundingRectangle(element);
        if (rect.HasValue)
        {
            var center = new Point(rect.Value.Left + rect.Value.Width / 2, rect.Value.Top + rect.Value.Height / 2);
            if (NativeUiaInput.RightClickPoint(center))
            {
                Thread.Sleep(ActionDelayMs);
                return (true, "physical-rightclick", null, attempted, null);
            }
        }

        return (false, null, "action-failed", attempted, null);
    }

    private (bool success, string? strategy, string? reason, List<string>? attemptedStrategies, string? actual) SetToggleElement(
        IUIAutomationElement element,
        DateTime deadlineUtc,
        CancellationToken cancellationToken,
        bool wantChecked)
    {
        cancellationToken.ThrowIfCancellationRequested();
        EnsureWithinDeadline(deadlineUtc, cancellationToken, wantChecked ? "check" : "uncheck");

        if (!_uia.TryGetTogglePattern(element, out var toggle))
            return (false, null, "toggle-unavailable", ["toggle-pattern"], null);

        var current = toggle!.CurrentToggleState;
        for (var i = 0; i < 2; i++)
        {
            if (wantChecked && current == ToggleState.ToggleState_On)
                return (true, "toggle-pattern", null, null, "On");
            if (!wantChecked && current == ToggleState.ToggleState_Off)
                return (true, "toggle-pattern", null, null, "Off");

            toggle.Toggle();
            current = toggle.CurrentToggleState;
        }

        var actual = current switch
        {
            ToggleState.ToggleState_On => "On",
            ToggleState.ToggleState_Off => "Off",
            _ => "Indeterminate"
        };

        var success = wantChecked ? current == ToggleState.ToggleState_On : current == ToggleState.ToggleState_Off;
        return success
            ? (true, "toggle-pattern", null, null, actual)
            : (false, null, "toggle-failed", ["toggle-pattern"], actual);
    }

    private object ExecuteSelectTabOperation(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken)
    {
        const string operation = "selecttabuia";
        var sw = Stopwatch.StartNew();
        var timeoutMs = ResolveTimeoutMs(request.TimeoutMs);
        var deadlineUtc = DateTime.UtcNow.AddMilliseconds(timeoutMs);

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            EnsureWithinDeadline(deadlineUtc, cancellationToken, "start");

            if (!HasLocator(request))
            {
                return BuildFailure(
                    operation,
                    "invalid-request",
                    "Locator is required for selecttabuia.",
                    sw.ElapsedMilliseconds);
            }

            if (request.Index == null && string.IsNullOrWhiteSpace(request.Value))
            {
                return BuildFailure(
                    operation,
                    "invalid-request",
                    "Either index or value is required for selecttabuia.",
                    sw.ElapsedMilliseconds);
            }

            if (string.IsNullOrWhiteSpace(request.Operation))
                request.Operation = operation;

            var resolveOutcome = TryResolveLocatedElement(
                operation,
                request,
                activeWindowHwnd,
                processId,
                deadlineUtc,
                cancellationToken,
                sw.ElapsedMilliseconds,
                requireEnabled: true);

            if (resolveOutcome.FailureResponse != null)
                return resolveOutcome.FailureResponse;

            var tabItem = ResolveTabItem(resolveOutcome.Element!, request, out var tabMatchStrategy);
            if (tabItem == null)
            {
                return new
                {
                    operation,
                    success = false,
                    reason = "tab-not-found",
                    stage = resolveOutcome.Stage,
                    elapsedMs = sw.ElapsedMilliseconds,
                    message = "Could not find a matching TabItem for the requested tab."
                };
            }

            var (success, strategy, reason, attemptedStrategies, actual) =
                SelectTabItemElement(tabItem, deadlineUtc, cancellationToken);

            if (success)
            {
                return new
                {
                    operation,
                    success = true,
                    strategy,
                    actual,
                    tabMatchStrategy,
                    stage = resolveOutcome.Stage,
                    elapsedMs = sw.ElapsedMilliseconds
                };
            }

            return new
            {
                operation,
                success = false,
                reason = reason ?? "action-failed",
                attemptedStrategies,
                tabMatchStrategy,
                stage = resolveOutcome.Stage,
                elapsedMs = sw.ElapsedMilliseconds
            };
        }
        catch (TimeoutException ex)
        {
            return BuildFailure(operation, "timeout", ex.Message, sw.ElapsedMilliseconds);
        }
        catch (OperationCanceledException)
        {
            return BuildFailure(operation, "cancelled", $"{operation} was cancelled.", sw.ElapsedMilliseconds);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "NativeUiaBasic selecttabuia failed with exception.");

            return new
            {
                operation,
                success = false,
                reason = "error",
                elapsedMs = sw.ElapsedMilliseconds,
                exceptionType = ex.GetType().Name,
                message = ex.Message
            };
        }
    }

    private object ExecuteScreenshotElementOperation(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken)
    {
        const string operation = "screenshotelementuia";
        var sw = Stopwatch.StartNew();
        var timeoutMs = ResolveTimeoutMs(request.TimeoutMs);
        var deadlineUtc = DateTime.UtcNow.AddMilliseconds(timeoutMs);

        return ExecuteQueryOperation(
            operation,
            request,
            activeWindowHwnd,
            processId,
            deadlineUtc,
            cancellationToken,
            sw,
            (element, stage) =>
            {
                var outputPath = string.IsNullOrWhiteSpace(request.Value)
                    ? null
                    : request.Value;

                var capture = NativeUiaScreenshot.CaptureElement(element, _uia, outputPath);
                if (!capture.success)
                {
                    return new
                    {
                        operation,
                        success = false,
                        reason = "screenshot-failed",
                        stage,
                        elapsedMs = sw.ElapsedMilliseconds,
                        message = capture.error
                    };
                }

                return new
                {
                    operation,
                    success = true,
                    reason = (string?)null,
                    strategy = "bounding-rect",
                    stage,
                    screenshot = capture.base64,
                    path = capture.path,
                    width = capture.width,
                    height = capture.height,
                    elapsedMs = sw.ElapsedMilliseconds
                };
            },
            onNotFound: (_, stage, view, resolveResult, elapsedMs) => new
            {
                operation,
                success = false,
                reason = resolveResult.Stage ?? "element-not-found",
                stage,
                view,
                elapsedMs,
                candidates = resolveResult.Candidates,
                message = resolveResult.LastError,
                diagnostics = resolveResult.Diagnostics
            });
    }

    private IUIAutomationElement? ResolveTabItem(IUIAutomationElement element, UiRequest request, out string tabMatchStrategy)
    {
        tabMatchStrategy = "direct-tabitem";

        var controlType = _uia.GetIntProperty(element, UIA_PropertyIds.UIA_ControlTypePropertyId);
        if (controlType == TabItemControlTypeId)
            return element;

        var tabItems = _uia.FindAllDescendants(
            element,
            _uia.ControlTypeCondition(TabItemControlTypeId),
            limit: 200);

        if (tabItems.Count == 0 && controlType != TabControlTypeId)
        {
            var tabContainer = _uia.WalkAncestor(
                element,
                el => _uia.GetIntProperty(el, UIA_PropertyIds.UIA_ControlTypePropertyId) == TabControlTypeId);

            if (tabContainer != null)
            {
                tabItems = _uia.FindAllDescendants(
                    tabContainer,
                    _uia.ControlTypeCondition(TabItemControlTypeId),
                    limit: 200);
            }
        }

        if (request.Index is >= 0)
        {
            tabMatchStrategy = "tab-index";
            return request.Index.Value < tabItems.Count ? tabItems[request.Index.Value] : null;
        }

        tabMatchStrategy = "tab-name";
        var matchMode = request.Locator?.MatchMode ?? "exact";
        return tabItems.FirstOrDefault(tab =>
            NativeUiaText.Matches(
                _uia.GetStringProperty(tab, UIA_PropertyIds.UIA_NamePropertyId),
                request.Value,
                matchMode));
    }

    private (bool success, string? strategy, string? reason, List<string>? attemptedStrategies, string? actual) SelectTabItemElement(
        IUIAutomationElement tabItem,
        DateTime deadlineUtc,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        EnsureWithinDeadline(deadlineUtc, cancellationToken, "selecttab");

        if (_uia.TryGetSelectionItemPattern(tabItem, out var selectionItem))
        {
            try
            {
                selectionItem!.Select();
                Thread.Sleep(ActionDelayMs);
                return (true, "selectionitem-select", null, null, null);
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "SelectionItem.Select failed for selecttabuia.");
            }
        }

        return ClickElement(tabItem, new UiRequest { Operation = "selecttabuia" }, deadlineUtc, cancellationToken);
    }

    private (bool success, string? strategy, string? reason, List<string>? attemptedStrategies, string? actual) TypeElement(
        IUIAutomationElement element,
        UiRequest request,
        DateTime deadlineUtc,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        EnsureWithinDeadline(deadlineUtc, cancellationToken, "type");

        var requestedText = WebUtility.HtmlDecode(request.Value ?? string.Empty);
        FocusElementInternal(element);

        if (TrySetWritableValue(element, requestedText, out var setStrategy))
        {
            Thread.Sleep(ActionDelayMs);
            var actual = ReadEditableValue(element);
            if (NativeUiaText.ValuesEquivalent(actual, requestedText))
                return (true, setStrategy ?? "valuepattern-setvalue", null, null, actual);

            _logger.LogDebug(
                "ValuePattern set value but verification failed. actual={Actual}, requested={Requested}",
                actual,
                requestedText);
        }

        NativeUiaInput.SendChord(
            NativeUiaInput.VirtualKeys.Control,
            NativeUiaInput.VirtualKeys.A);
        NativeUiaInput.SendKey(NativeUiaInput.VirtualKeys.Backspace);
        Thread.Sleep(ActionDelayMs);
        NativeUiaInput.TypeText(requestedText);
        Thread.Sleep(ActionDelayMs);

        var keyboardActual = ReadEditableValue(element);
        if (NativeUiaText.ValuesEquivalent(keyboardActual, requestedText))
            return (true, "keyboard-type", null, null, keyboardActual);

        return (true, "keyboard-type", null, null, null);
    }

    private (bool success, string? strategy, string? reason, List<string>? attemptedStrategies, string? actual) SendKeysElement(
        IUIAutomationElement element,
        UiRequest request,
        DateTime deadlineUtc,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        EnsureWithinDeadline(deadlineUtc, cancellationToken, "sendkeys");

        FocusElementInternal(element);
        SendKeysValue(request.Value!);
        Thread.Sleep(ActionDelayMs);

        return (true, "sendkeys", null, null, null);
    }

    private (bool success, string? strategy, string? reason, List<string>? attemptedStrategies, string? actual) ClearElement(
        IUIAutomationElement element,
        UiRequest request,
        DateTime deadlineUtc,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        EnsureWithinDeadline(deadlineUtc, cancellationToken, "clear");

        if (TrySetWritableValue(element, string.Empty, out _))
        {
            Thread.Sleep(ActionDelayMs);
            var actual = ReadEditableValue(element);
            if (string.IsNullOrEmpty(NativeUiaText.Normalize(actual)))
                return (true, "valuepattern-empty", null, null, actual);
        }

        FocusElementInternal(element);
        NativeUiaInput.SendChord(
            NativeUiaInput.VirtualKeys.Control,
            NativeUiaInput.VirtualKeys.A);
        NativeUiaInput.SendKey(NativeUiaInput.VirtualKeys.Backspace);
        Thread.Sleep(ActionDelayMs);

        var keyboardActual = ReadEditableValue(element);
        if (string.IsNullOrEmpty(NativeUiaText.Normalize(keyboardActual)))
            return (true, "keyboard-clear", null, null, keyboardActual);

        return (true, "keyboard-clear", null, null, keyboardActual);
    }

    private (bool success, string? strategy, string? reason, List<string>? attemptedStrategies, string? actual) FocusElement(
        IUIAutomationElement element,
        UiRequest request,
        DateTime deadlineUtc,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        EnsureWithinDeadline(deadlineUtc, cancellationToken, "focus");

        if (TryFocusElement(element, out var strategy))
            return (true, strategy, null, null, null);

        return (false, null, "action-failed", ["uia-setfocus", "native-window-focus", "physical-click"], null);
    }

    private bool TryFocusElement(IUIAutomationElement element, out string strategy)
    {
        try
        {
            element.SetFocus();
            Thread.Sleep(ActionDelayMs);
            strategy = "uia-setfocus";
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "SetFocus failed.");
        }

        var hwnd = _uia.GetIntProperty(element, UIA_PropertyIds.UIA_NativeWindowHandlePropertyId);
        if (hwnd > 0 && NativeUiaInput.FocusWindow(new IntPtr(hwnd)))
        {
            Thread.Sleep(ActionDelayMs);
            strategy = "native-window-focus";
            return true;
        }

        var rect = _uia.GetBoundingRectangle(element);
        if (rect.HasValue)
        {
            var center = new Point(rect.Value.Left + rect.Value.Width / 2, rect.Value.Top + rect.Value.Height / 2);
            if (NativeUiaInput.ClickPoint(center))
            {
                Thread.Sleep(ActionDelayMs);
                strategy = "physical-click";
                return true;
            }
        }

        strategy = string.Empty;
        return false;
    }

    private void FocusElementInternal(IUIAutomationElement element)
    {
        if (!TryFocusElement(element, out _))
        {
            try
            {
                element.SetFocus();
            }
            catch
            {
                // best effort
            }
        }
    }

    private bool TrySetWritableValue(IUIAutomationElement element, string value, out string? strategy)
    {
        strategy = null;
        try
        {
            if (element.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId) is IUIAutomationValuePattern valuePattern
                && valuePattern.CurrentIsReadOnly == 0)
            {
                valuePattern.SetValue(value);
                strategy = "valuepattern-setvalue";
                return true;
            }

            var innerEdit = _uia.FindFirstDescendant(
                element,
                _uia.PropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId, 50004));
            if (innerEdit?.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId) is IUIAutomationValuePattern editValue
                && editValue.CurrentIsReadOnly == 0)
            {
                editValue.SetValue(value);
                strategy = "edit-valuepattern-setvalue";
                return true;
            }
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "TrySetWritableValue failed.");
        }

        return false;
    }

    private string ReadEditableValue(IUIAutomationElement element)
    {
        var value = _uia.GetValuePatternText(element);
        if (!string.IsNullOrWhiteSpace(value))
            return value;

        return _textReader.ReadElementText(element);
    }

    private static void SendKeysValue(string value)
    {
        WinForms.SendKeys.SendWait(value);
    }

    private static bool HasLocator(UiRequest request) =>
        request.Hwnd is > 0
        || request.Locator?.Hwnd is > 0
        || request.Locator?.Handle is > 0
        || !string.IsNullOrWhiteSpace(request.Locator?.AutomationId)
        || !string.IsNullOrWhiteSpace(request.Locator?.Name)
        || !string.IsNullOrWhiteSpace(request.Locator?.ClassName)
        || !string.IsNullOrWhiteSpace(request.Locator?.ControlType);

    private static int ResolveTimeoutMs(int? requestTimeoutMs)
    {
        var timeout = requestTimeoutMs ?? DefaultTimeoutMs;
        return Math.Clamp(timeout, 500, MaxTimeoutMs);
    }

    private static void EnsureWithinDeadline(DateTime deadlineUtc, CancellationToken cancellationToken, string stage)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (DateTime.UtcNow > deadlineUtc)
            throw new TimeoutException($"{stage} exceeded timeout.");
    }

    private static object BuildFailure(string operation, string reason, string message, long elapsedMs) => new
    {
        operation,
        success = false,
        reason,
        elapsedMs,
        message
    };
}
