using System.Diagnostics;
using System.Drawing;
using System.Net;
using DesktopAutomationDriver.Models.Request;
using Interop.UIAutomationClient;

namespace DesktopAutomationDriver.Services.NativeUia;

/// <summary>
/// pywinauto-style native UIA-only ComboBox selection. Bounded by request timeoutMs.
/// </summary>
internal sealed class NativeUiaComboBoxService : INativeUiaComboBoxService
{
    private const int ComboBoxControlTypeId = 50003;
    private const int ButtonControlTypeId = 50000;
    private const int DefaultTimeoutMs = 8000;
    private const int MaxTimeoutMs = 15000;
    private const int ExpandDelayMs = 120;
    private const int ActionDelayMs = 80;
    private const int MaxBoundedScrollPages = 6;

    private readonly NativeUiaAutomation _uia;
    private readonly NativeUiaElementResolver _resolver;
    private readonly NativeUiaDropdownFinder _dropdownFinder;
    private readonly NativeUiaTextReader _textReader;
    private readonly ILogger<NativeUiaComboBoxService> _logger;

    public NativeUiaComboBoxService(ILogger<NativeUiaComboBoxService> logger)
        : this(new NativeUiaAutomation(), logger)
    {
    }

    internal NativeUiaComboBoxService(NativeUiaAutomation uia, ILogger<NativeUiaComboBoxService> logger)
    {
        _uia = uia;
        _textReader = new NativeUiaTextReader(_uia);
        _resolver = new NativeUiaElementResolver(_uia);
        _dropdownFinder = new NativeUiaDropdownFinder(_uia, _textReader);
        _logger = logger;
    }

    public object SelectComboBox(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default)
    {
        var sw = Stopwatch.StartNew();
        var operation = string.IsNullOrWhiteSpace(request.Operation) ? "select" : request.Operation;
        var timeoutMs = ResolveTimeoutMs(request.TimeoutMs);
        var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);

        _logger.LogInformation(
            "NativeUiaComboBox {Operation} starting. rootHwnd={RootHwnd}, processId={ProcessId}, timeoutMs={TimeoutMs}",
            operation,
            activeWindowHwnd,
            processId,
            timeoutMs);

        cancellationToken.ThrowIfCancellationRequested();

        if (IsUiaOnlyOperation(operation) && !HasSearchContext(request, activeWindowHwnd, processId))
        {
            return BuildFailure(
                operation,
                request,
                "no-search-context",
                "No active window hwnd or processId. Launch/attach and switchwindow first.",
                sw.ElapsedMilliseconds);
        }

        try
        {
            if (!TryEnsureWithinDeadline(deadline, cancellationToken, sw, operation, "start", out var startTimeout))
                return startTimeout!;

            var resolveResult = _resolver.ResolveComboBox(
                request,
                activeWindowHwnd,
                processId,
                cancellationToken,
                deadline,
                allowDesktopFallback: !IsUiaOnlyOperation(operation));

            if (resolveResult.Element == null)
            {
                _logger.LogWarning(
                    "NativeUiaComboBox {Operation} resolve failed. stage={Stage}, elapsedMs={ElapsedMs}",
                    operation,
                    resolveResult.Stage,
                    sw.ElapsedMilliseconds);

                return BuildFailure(
                    operation,
                    request,
                    resolveResult.Stage ?? "combo-not-found",
                    resolveResult.LastError ?? "ComboBox not found.",
                    sw.ElapsedMilliseconds,
                    resolveResult.Candidates);
            }

            if (resolveResult.IsAmbiguous)
            {
                return BuildFailure(
                    operation,
                    request,
                    "ambiguous-combobox",
                    resolveResult.LastError ?? "Multiple ComboBoxes matched.",
                    sw.ElapsedMilliseconds,
                    resolveResult.Candidates);
            }

            var combo = resolveResult.Element;
            var comboSnapshot = _uia.CreateSnapshot(combo);
            string? expandedBy = null;
            string? selectionStrategy = null;
            IUIAutomationElement? matchedItem = null;

            try
            {
                if (!_uia.GetBoolProperty(combo, UIA_PropertyIds.UIA_IsEnabledPropertyId, true))
                {
                    return BuildFailure(
                        operation,
                        request,
                        "combo-disabled",
                        "ComboBox is disabled.",
                        sw.ElapsedMilliseconds,
                        comboSnapshot: comboSnapshot);
                }

                var requestedValue = request.Index.HasValue && string.IsNullOrWhiteSpace(request.Value)
                    ? null
                    : WebUtility.HtmlDecode(request.Value ?? string.Empty).Trim();

                if (request.Index.HasValue && string.IsNullOrWhiteSpace(requestedValue))
                {
                    expandedBy = ExpandComboBox(combo, cancellationToken, deadline);
                    Thread.Sleep(ExpandDelayMs);
                    matchedItem = SelectByIndex(combo, activeWindowHwnd, processId, request.Index.Value, cancellationToken, deadline);
                    requestedValue = matchedItem == null ? $"index:{request.Index.Value}" : _textReader.ReadElementText(matchedItem);
                    selectionStrategy = matchedItem == null ? null : "index-select";
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(requestedValue))
                        return BuildFailure(
                            operation,
                            request,
                            "invalid-request",
                            "Either 'value' or 'index' is required for ComboBox select.",
                            sw.ElapsedMilliseconds,
                            comboSnapshot: comboSnapshot);

                    if (TrySetEditableValue(combo, requestedValue, cancellationToken, deadline, out selectionStrategy))
                    {
                        var directActual = _textReader.ReadComboBoxValue(combo);
                        if (NativeUiaText.ValuesEquivalent(directActual, requestedValue))
                        {
                            return BuildSuccess(
                                operation,
                                requestedValue,
                                directActual,
                                verified: true,
                                selectionStrategy ?? "valuepattern-setvalue",
                                expandedBy: "ValuePattern.SetValue",
                                comboSnapshot,
                                matchedItemSnapshot: null,
                                sw.ElapsedMilliseconds);
                        }
                    }

                    expandedBy = ExpandComboBox(combo, cancellationToken, deadline);
                    Thread.Sleep(ExpandDelayMs);

                    matchedItem = FindDropdownItem(
                        combo,
                        activeWindowHwnd,
                        processId,
                        requestedValue,
                        request,
                        cancellationToken,
                        deadline);

                    if (matchedItem != null)
                        selectionStrategy = ActivateItem(combo, matchedItem, cancellationToken, deadline);

                    if ((matchedItem == null || selectionStrategy == null)
                        && request.AllowKeyboardFallback == true
                        && DateTime.UtcNow < deadline)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (TryKeyboardTypeahead(combo, requestedValue, cancellationToken, deadline))
                            selectionStrategy = "keyboard-typeahead";
                    }

                    if (matchedItem == null
                        && DateTime.UtcNow < deadline
                        && TryBoundedScrollSearch(
                            combo,
                            activeWindowHwnd,
                            processId,
                            requestedValue,
                            request,
                            cancellationToken,
                            deadline,
                            out matchedItem,
                            out selectionStrategy))
                    {
                        // bounded scroll found item
                    }
                }

                if (!TryEnsureWithinDeadline(deadline, cancellationToken, sw, operation, "verify", out var verifyTimeout))
                    return verifyTimeout!;

                var actual = _textReader.ReadComboBoxValue(combo);
                var requested = requestedValue ?? request.Value;
                var verified = !string.IsNullOrWhiteSpace(requested)
                               && NativeUiaText.ValuesEquivalent(actual, requested);
                if (request.Index.HasValue && string.IsNullOrWhiteSpace(request.Value))
                    verified = matchedItem != null;

                if (verified || !string.IsNullOrWhiteSpace(selectionStrategy))
                {
                    _logger.LogInformation(
                        "NativeUiaComboBox {Operation} succeeded. strategy={Strategy}, verified={Verified}, elapsedMs={ElapsedMs}",
                        operation,
                        selectionStrategy ?? "native-uia-unverified",
                        verified,
                        sw.ElapsedMilliseconds);

                    return BuildSuccess(
                        operation,
                        requested,
                        actual,
                        verified,
                        selectionStrategy ?? "native-uia-unverified",
                        expandedBy,
                        comboSnapshot,
                        matchedItem == null ? null : _uia.CreateSnapshot(matchedItem),
                        sw.ElapsedMilliseconds);
                }

                var candidates = _dropdownFinder.CollectVisibleItemTexts(
                    combo, activeWindowHwnd, processId, cancellationToken, deadline, 50)
                    .Cast<object>()
                    .ToList();

                return BuildFailure(
                    operation,
                    request,
                    "item-not-found",
                    $"ComboBox item '{requested}' not found.",
                    sw.ElapsedMilliseconds,
                    candidates,
                    comboSnapshot,
                    expandedBy);
            }
            finally
            {
                CollapseComboBox(combo);
            }
        }
        catch (TimeoutException ex)
        {
            _logger.LogWarning(
                ex,
                "NativeUiaComboBox {Operation} timed out. elapsedMs={ElapsedMs}",
                operation,
                sw.ElapsedMilliseconds);

            return BuildFailure(
                operation,
                request,
                "timeout",
                ex.Message,
                sw.ElapsedMilliseconds);
        }
    }

    public object FindComboBox(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default)
    {
        var sw = Stopwatch.StartNew();
        const string operation = "findcomboboxuia";
        var timeoutMs = ResolveTimeoutMs(request.TimeoutMs);
        var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);

        _logger.LogInformation(
            "NativeUiaComboBox {Operation} starting. rootHwnd={RootHwnd}, processId={ProcessId}, timeoutMs={TimeoutMs}",
            operation,
            activeWindowHwnd,
            processId,
            timeoutMs);

        cancellationToken.ThrowIfCancellationRequested();

        if (!HasSearchContext(request, activeWindowHwnd, processId))
        {
            _logger.LogWarning(
                "NativeUiaComboBox {Operation} failed fast: no-search-context. elapsedMs={ElapsedMs}",
                operation,
                sw.ElapsedMilliseconds);

            return new
            {
                operation,
                found = false,
                success = false,
                stage = "no-search-context",
                error = "No active window hwnd or processId. Launch/attach and switchwindow first.",
                rootHwnd = activeWindowHwnd,
                processId,
                timeoutMs,
                elapsedMs = sw.ElapsedMilliseconds
            };
        }

        try
        {
            var resolveResult = _resolver.ResolveComboBox(
                request,
                activeWindowHwnd,
                processId,
                cancellationToken,
                deadline,
                allowDesktopFallback: false);

            if (resolveResult.Element == null)
            {
                _logger.LogWarning(
                    "NativeUiaComboBox {Operation} not found. stage={Stage}, elapsedMs={ElapsedMs}",
                    operation,
                    resolveResult.Stage,
                    sw.ElapsedMilliseconds);

                return new
                {
                    operation,
                    found = false,
                    success = false,
                    stage = resolveResult.Stage ?? "combo-not-found",
                    error = resolveResult.LastError ?? "ComboBox not found.",
                    ambiguous = resolveResult.IsAmbiguous,
                    candidates = resolveResult.Candidates,
                    rootHwnd = activeWindowHwnd,
                    processId,
                    timeoutMs,
                    elapsedMs = sw.ElapsedMilliseconds
                };
            }

            var snapshot = _uia.CreateSnapshot(resolveResult.Element);
            _logger.LogInformation(
                "NativeUiaComboBox {Operation} found ComboBox. automationId={AutomationId}, elapsedMs={ElapsedMs}",
                operation,
                snapshot.AutomationId,
                sw.ElapsedMilliseconds);

            return new
            {
                operation,
                found = true,
                success = true,
                strategy = "native-uia-resolve",
                ambiguous = resolveResult.IsAmbiguous,
                comboBox = NativeUiaDiagnostics.ComboSummary(snapshot),
                candidates = resolveResult.Candidates,
                rootHwnd = activeWindowHwnd,
                processId,
                timeoutMs,
                elapsedMs = sw.ElapsedMilliseconds
            };
        }
        catch (TimeoutException ex)
        {
            _logger.LogWarning(
                ex,
                "NativeUiaComboBox {Operation} timed out. elapsedMs={ElapsedMs}",
                operation,
                sw.ElapsedMilliseconds);

            return new
            {
                operation,
                found = false,
                success = false,
                stage = "timeout",
                error = ex.Message,
                rootHwnd = activeWindowHwnd,
                processId,
                timeoutMs,
                elapsedMs = sw.ElapsedMilliseconds
            };
        }
    }

    private static bool IsUiaOnlyOperation(string operation) =>
        string.Equals(operation, "selectcomboboxuia", StringComparison.OrdinalIgnoreCase)
        || string.Equals(operation, "findcomboboxuia", StringComparison.OrdinalIgnoreCase);

    private static bool HasSearchContext(UiRequest request, IntPtr? activeWindowHwnd, int? processId)
    {
        if (request.Hwnd is > 0 || request.Locator?.Hwnd is > 0 || request.Locator?.Handle is > 0)
            return true;

        if (activeWindowHwnd is > 0)
            return true;

        return processId.HasValue;
    }

    private static bool TryEnsureWithinDeadline(
        DateTime deadline,
        CancellationToken cancellationToken,
        Stopwatch sw,
        string operation,
        string stage,
        out object? timeoutFailure)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (DateTime.UtcNow <= deadline)
        {
            timeoutFailure = null;
            return true;
        }

        timeoutFailure = new Dictionary<string, object?>
        {
            ["operation"] = operation,
            ["success"] = false,
            ["stage"] = "timeout",
            ["error"] = $"ComboBox {operation} timed out during {stage} after {sw.ElapsedMilliseconds} ms.",
            ["elapsedMs"] = sw.ElapsedMilliseconds
        };
        return false;
    }

    public object InspectComboBox(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default)
    {
        var timeoutMs = ResolveTimeoutMs(request.TimeoutMs);
        var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);
        cancellationToken.ThrowIfCancellationRequested();

        var resolveResult = _resolver.ResolveComboBox(
            request,
            activeWindowHwnd,
            processId,
            cancellationToken,
            deadline,
            allowDesktopFallback: true);
        if (resolveResult.Element == null)
        {
            return new
            {
                operation = "inspectcombobox",
                found = false,
                stage = resolveResult.Stage,
                error = resolveResult.LastError,
                candidates = resolveResult.Candidates
            };
        }

        var combo = resolveResult.Element;
        var snapshot = _uia.CreateSnapshot(combo);
        var expandedBy = ExpandComboBox(combo, cancellationToken, deadline);
        Thread.Sleep(ExpandDelayMs);

        try
        {
            var itemsPreview = _dropdownFinder.BuildItemsPreview(
                combo, activeWindowHwnd, processId, cancellationToken, deadline, 50);

            return new
            {
                operation = "inspectcombobox",
                found = true,
                ambiguous = resolveResult.IsAmbiguous,
                comboBox = new
                {
                    name = snapshot.Name,
                    automationId = snapshot.AutomationId,
                    className = snapshot.ClassName,
                    frameworkId = snapshot.FrameworkId,
                    rectangle = NativeUiaDiagnostics.ToRectangleObject(snapshot.BoundingRectangle),
                    isEnabled = snapshot.IsEnabled,
                    isOffscreen = snapshot.IsOffscreen,
                    expandCollapseSupported = snapshot.SupportedPatterns.Contains("ExpandCollapse"),
                    valuePatternSupported = snapshot.SupportedPatterns.Contains("Value"),
                    currentValue = _textReader.ReadComboBoxValue(combo),
                    expandState = GetExpandState(combo)
                },
                expandedBy,
                itemsPreview,
                candidates = resolveResult.Candidates
            };
        }
        finally
        {
            CollapseComboBox(combo);
        }
    }

    private static int ResolveTimeoutMs(int? requestTimeoutMs)
    {
        var timeout = requestTimeoutMs ?? DefaultTimeoutMs;
        return Math.Clamp(timeout, 500, MaxTimeoutMs);
    }

    private static void EnsureWithinDeadline(
        DateTime deadline,
        CancellationToken cancellationToken,
        Stopwatch sw,
        string operation,
        string stage)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (DateTime.UtcNow > deadline)
        {
            throw new TimeoutException(
                $"ComboBox {operation} timed out during {stage} after {sw.ElapsedMilliseconds} ms.");
        }
    }

    private bool TrySetEditableValue(
        IUIAutomationElement combo,
        string requestedValue,
        CancellationToken cancellationToken,
        DateTime deadline,
        out string? strategy)
    {
        strategy = null;
        cancellationToken.ThrowIfCancellationRequested();
        if (DateTime.UtcNow >= deadline)
            return false;

        try
        {
            if (combo.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId) is IUIAutomationValuePattern valuePattern
                && valuePattern.CurrentIsReadOnly == 0)
            {
                valuePattern.SetValue(requestedValue);
                Thread.Sleep(ActionDelayMs);
                strategy = "valuepattern-setvalue";
                return true;
            }

            var innerEdit = _uia.FindFirstDescendant(
                combo,
                _uia.PropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId, 50004));
            if (innerEdit?.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId) is IUIAutomationValuePattern editValue
                && editValue.CurrentIsReadOnly == 0)
            {
                editValue.SetValue(requestedValue);
                Thread.Sleep(ActionDelayMs);
                strategy = "edit-valuepattern-setvalue";
                return true;
            }
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "TrySetEditableValue failed.");
        }

        return false;
    }

    private string ExpandComboBox(IUIAutomationElement combo, CancellationToken cancellationToken, DateTime deadline)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (_uia.TryGetExpandCollapsePattern(combo, out var expandCollapse))
        {
            try
            {
                expandCollapse!.Expand();
                return "expandcollapse";
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "ExpandCollapsePattern.Expand failed.");
            }
        }

        var openButton = FindOpenButton(combo);
        if (openButton != null && _uia.TryGetInvokePattern(openButton, out var openInvoke))
        {
            try
            {
                openInvoke!.Invoke();
                return "open-button-invoke";
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "Open button invoke failed.");
            }
        }

        if (_uia.TryGetInvokePattern(combo, out var invoke))
        {
            try
            {
                invoke!.Invoke();
                return "invoke";
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "ComboBox Invoke failed.");
            }
        }

        var rect = _uia.GetBoundingRectangle(combo);
        if (rect.HasValue)
        {
            var clickPoint = ComboBoxRightEdgePoint(rect.Value);
            if (NativeUiaInput.ClickPoint(clickPoint))
                return "physical-right-edge-click";
        }

        FocusComboBox(combo);
        NativeUiaInput.SendChord(NativeUiaInput.VirtualKeys.Menu, NativeUiaInput.VirtualKeys.Down);
        Thread.Sleep(ActionDelayMs);
        return "alt-down";
    }

    private IUIAutomationElement? FindOpenButton(IUIAutomationElement combo)
    {
        foreach (var child in _uia.FindAllDescendants(
                     combo,
                     _uia.PropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId, ButtonControlTypeId),
                     20))
        {
            var name = _uia.GetStringProperty(child, UIA_PropertyIds.UIA_NamePropertyId);
            if (name.Contains("Open", StringComparison.OrdinalIgnoreCase)
                || name.Contains("Drop", StringComparison.OrdinalIgnoreCase))
                return child;
        }

        return null;
    }

    private IUIAutomationElement? FindDropdownItem(
        IUIAutomationElement combo,
        IntPtr? activeWindowHwnd,
        int? processId,
        string requestedValue,
        UiRequest request,
        CancellationToken cancellationToken,
        DateTime deadline)
    {
        if (request.Index.HasValue && string.IsNullOrWhiteSpace(request.Value))
            return SelectByIndex(combo, activeWindowHwnd, processId, request.Index.Value, cancellationToken, deadline);

        foreach (var candidate in _dropdownFinder.EnumerateItemCandidates(
                     combo, activeWindowHwnd, processId, cancellationToken, deadline))
        {
            var texts = _textReader.ReadCandidateTexts(candidate);
            if (texts.Any(t => ItemMatches(t, requestedValue, request)))
                return candidate;
        }

        return null;
    }

    private IUIAutomationElement? SelectByIndex(
        IUIAutomationElement combo,
        IntPtr? activeWindowHwnd,
        int? processId,
        int index,
        CancellationToken cancellationToken,
        DateTime deadline)
    {
        if (index < 0)
            throw new ArgumentOutOfRangeException(nameof(index), "ComboBox index must be >= 0.");

        var items = _dropdownFinder
            .EnumerateItemCandidates(combo, activeWindowHwnd, processId, cancellationToken, deadline)
            .ToList();

        return index < items.Count ? items[index] : null;
    }

    private string? ActivateItem(
        IUIAutomationElement combo,
        IUIAutomationElement item,
        CancellationToken cancellationToken,
        DateTime deadline)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (DateTime.UtcNow >= deadline)
            return null;

        if (_uia.TryGetSelectionItemPattern(item, out var selectionItem))
        {
            try
            {
                selectionItem!.Select();
                Thread.Sleep(ActionDelayMs);
                return "selectionitem-select";
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "SelectionItem.Select failed.");
            }
        }

        if (_uia.TryGetInvokePattern(item, out var invoke))
        {
            try
            {
                invoke!.Invoke();
                Thread.Sleep(ActionDelayMs);
                return "invoke";
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "Invoke failed.");
            }
        }

        var controlType = _uia.GetIntProperty(item, UIA_PropertyIds.UIA_ControlTypePropertyId);
        if (controlType is 50002 or 50013 && _uia.TryGetTogglePattern(item, out var toggle))
        {
            try
            {
                toggle!.Toggle();
                Thread.Sleep(ActionDelayMs);
                return "toggle";
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "Toggle failed.");
            }
        }

        var rect = _uia.GetBoundingRectangle(item);
        if (rect.HasValue)
        {
            var center = new Point(rect.Value.Left + rect.Value.Width / 2, rect.Value.Top + rect.Value.Height / 2);
            var leftCenter = new Point(rect.Value.Left + Math.Min(Math.Max(rect.Value.Width / 10, 6), 18), center.Y);

            if (NativeUiaInput.ClickPoint(center))
            {
                Thread.Sleep(ActionDelayMs);
                return "physical-click-center";
            }

            if (NativeUiaInput.ClickPoint(leftCenter))
            {
                Thread.Sleep(ActionDelayMs);
                return "physical-click-leftcenter";
            }
        }

        try
        {
            item.SetFocus();
            Thread.Sleep(40);
            NativeUiaInput.SendKey(NativeUiaInput.VirtualKeys.Return);
            Thread.Sleep(ActionDelayMs);
            return "focus-enter";
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Focus+Enter failed.");
        }

        try
        {
            item.SetFocus();
            NativeUiaInput.SendKey(NativeUiaInput.VirtualKeys.Space);
            Thread.Sleep(ActionDelayMs);
            return "focus-space";
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Focus+Space failed.");
        }

        return null;
    }

    private bool TryKeyboardTypeahead(
        IUIAutomationElement combo,
        string requestedValue,
        CancellationToken cancellationToken,
        DateTime deadline)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (DateTime.UtcNow >= deadline)
            return false;

        try
        {
            FocusComboBox(combo);
            NativeUiaInput.TypeText(requestedValue);
            Thread.Sleep(ActionDelayMs);
            NativeUiaInput.SendKey(NativeUiaInput.VirtualKeys.Return);
            Thread.Sleep(ActionDelayMs);
            return NativeUiaText.ValuesEquivalent(_textReader.ReadComboBoxValue(combo), requestedValue);
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Keyboard typeahead failed.");
            return false;
        }
    }

    private bool TryBoundedScrollSearch(
        IUIAutomationElement combo,
        IntPtr? activeWindowHwnd,
        int? processId,
        string requestedValue,
        UiRequest request,
        CancellationToken cancellationToken,
        DateTime deadline,
        out IUIAutomationElement? matchedItem,
        out string? strategy)
    {
        matchedItem = null;
        strategy = null;
        string? previousBatch = null;

        for (var page = 0; page < MaxBoundedScrollPages && DateTime.UtcNow < deadline; page++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            matchedItem = FindDropdownItem(
                combo, activeWindowHwnd, processId, requestedValue, request, cancellationToken, deadline);
            if (matchedItem != null)
            {
                strategy = ActivateItem(combo, matchedItem, cancellationToken, deadline);
                if (strategy != null)
                    return true;
            }

            var batch = string.Join(
                "|",
                _dropdownFinder.CollectVisibleItemTexts(combo, activeWindowHwnd, processId, cancellationToken, deadline, 20));
            if (string.Equals(batch, previousBatch, StringComparison.Ordinal))
                break;

            previousBatch = batch;
            NativeUiaInput.WheelDown(3);
            Thread.Sleep(ActionDelayMs);
        }

        matchedItem = null;
        strategy = null;
        return false;
    }

    private void CollapseComboBox(IUIAutomationElement combo)
    {
        try
        {
            if (_uia.TryGetExpandCollapsePattern(combo, out var pattern))
            {
                pattern!.Collapse();
                return;
            }
        }
        catch
        {
            // continue to ESC fallback
        }

        try
        {
            FocusComboBox(combo);
            NativeUiaInput.SendKey(NativeUiaInput.VirtualKeys.Escape, release: false);
            NativeUiaInput.SendKey(NativeUiaInput.VirtualKeys.Escape, release: true);
        }
        catch
        {
            // best effort cleanup
        }
    }

    private void FocusComboBox(IUIAutomationElement combo)
    {
        try
        {
            combo.SetFocus();
        }
        catch
        {
            var hwnd = _uia.GetIntProperty(combo, UIA_PropertyIds.UIA_NativeWindowHandlePropertyId);
            if (hwnd > 0)
                NativeUiaInput.FocusWindow(new IntPtr(hwnd));
        }
    }

    private static Point ComboBoxRightEdgePoint(Rectangle rect)
    {
        var offset = Math.Min(Math.Max(rect.Width / 8, 8), 20);
        return new Point(rect.Right - offset, rect.Top + rect.Height / 2);
    }

    private string GetExpandState(IUIAutomationElement combo)
    {
        if (!_uia.TryGetExpandCollapsePattern(combo, out var pattern) || pattern == null)
            return "unsupported";

        try
        {
            return pattern.CurrentExpandCollapseState switch
            {
                ExpandCollapseState.ExpandCollapseState_Expanded => "expanded",
                ExpandCollapseState.ExpandCollapseState_Collapsed => "collapsed",
                ExpandCollapseState.ExpandCollapseState_PartiallyExpanded => "partially-expanded",
                _ => "unknown"
            };
        }
        catch
        {
            return "unknown";
        }
    }

    private static bool ItemMatches(string candidate, string requested, UiRequest request)
    {
        var matchMode = string.IsNullOrWhiteSpace(request.MatchMode) ? "exact" : request.MatchMode!;
        return NativeUiaText.Matches(candidate, requested, matchMode);
    }

    private static object BuildSuccess(
        string operation,
        string? requested,
        string actual,
        bool verified,
        string strategy,
        string? expandedBy,
        NativeUiaElementSnapshot comboSnapshot,
        NativeUiaElementSnapshot? matchedItemSnapshot,
        long elapsedMs) => new Dictionary<string, object?>
    {
        ["operation"] = operation,
        ["controlType"] = "ComboBox",
        ["selected"] = requested,
        ["success"] = true,
        ["strategy"] = strategy,
        ["expandedBy"] = expandedBy,
        ["matchedItem"] = matchedItemSnapshot == null
            ? null
            : NativeUiaDiagnostics.CandidateDiagnostic(0, matchedItemSnapshot),
        ["actualValue"] = actual,
        ["actual"] = actual,
        ["verified"] = verified,
        ["comboBox"] = NativeUiaDiagnostics.ComboSummary(comboSnapshot),
        ["combo"] = NativeUiaDiagnostics.ComboSummary(comboSnapshot),
        ["elapsedMs"] = elapsedMs
    };

    private static object BuildFailure(
        string operation,
        UiRequest request,
        string stage,
        string error,
        long elapsedMs,
        IEnumerable<object>? candidates = null,
        NativeUiaElementSnapshot? comboSnapshot = null,
        string? expandedBy = null) => new Dictionary<string, object?>
    {
        ["operation"] = operation,
        ["controlType"] = "ComboBox",
        ["success"] = false,
        ["error"] = error,
        ["stage"] = stage,
        ["requested"] = request.Value ?? (request.Index.HasValue ? $"index:{request.Index}" : null),
        ["comboBox"] = comboSnapshot == null ? null : NativeUiaDiagnostics.ComboSummary(comboSnapshot),
        ["combo"] = comboSnapshot == null ? null : NativeUiaDiagnostics.ComboSummary(comboSnapshot),
        ["expandedBy"] = expandedBy,
        ["candidates"] = candidates?.ToList() ?? [],
        ["elapsedMs"] = elapsedMs
    };
}
