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
    private const int MenuItemControlTypeId = 50010;
    private const int MenuControlTypeId = 50011;
    private const int CheckBoxControlTypeId = 50002;
    private const int RadioButtonControlTypeId = 50013;
    private const int DefaultTimeoutMs = 5000;
    private const int MaxTimeoutMs = 15000;
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

                var resolveView = request.View
                                  ?? request.TreeView
                                  ?? NativeUiaElementResolver.InferDefaultView(request);

                var resolveResult = _resolver.ResolveElement(
                    request,
                    activeWindowHwnd,
                    processId,
                    deadlineUtc,
                    cancellationToken);

                resolverStage = resolveResult.Stage;

                _logger.LogInformation(
                    "NativeUiaBasic {Operation} resolver stage={Stage}, view={View}, elapsedMs={ElapsedMs}",
                    operation,
                    resolverStage,
                    resolveView,
                    sw.ElapsedMilliseconds);

                if (resolveResult.Element == null)
                {
                    var view = request.View
                               ?? request.TreeView
                               ?? NativeUiaElementResolver.InferDefaultView(request);

                    return new
                    {
                        operation,
                        success = false,
                        reason = resolveResult.Stage ?? "element-not-found",
                        stage = resolverStage,
                        view,
                        elapsedMs = sw.ElapsedMilliseconds,
                        candidates = resolveResult.Candidates,
                        message = resolveResult.LastError
                    };
                }

                if (resolveResult.IsAmbiguous)
                {
                    return new
                    {
                        operation,
                        success = false,
                        reason = "ambiguous-element",
                        stage = resolverStage,
                        elapsedMs = sw.ElapsedMilliseconds,
                        candidates = resolveResult.Candidates,
                        message = resolveResult.LastError
                    };
                }

                element = resolveResult.Element;

                if (!_uia.GetBoolProperty(element, UIA_PropertyIds.UIA_IsEnabledPropertyId, true))
                {
                    return BuildFailure(
                        operation,
                        "element-disabled",
                        "Target element is disabled.",
                        sw.ElapsedMilliseconds);
                }
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
