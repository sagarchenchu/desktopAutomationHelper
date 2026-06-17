using System.Drawing;
using System.Net;
using DesktopAutomationDriver.Models.Request;
using Interop.UIAutomationClient;

namespace DesktopAutomationDriver.Services.NativeUia;

/// <summary>
/// pywinauto-style native UIA ComboBox selection engine (no FlaUI).
/// </summary>
internal sealed class NativeUiaComboBoxService : INativeUiaComboBoxService
{
    private const int ComboBoxControlTypeId = 50003;
    private const int EditControlTypeId = 50004;
    private const int ExpandDelayMs = 250;
    private const int ActionDelayMs = 150;
    private const int PagedSearchMaxPages = 300;
    private const int PagedSearchVisibleLimit = 20;
    private const int DefaultTimeoutMs = 20000;

    private static readonly int[] ItemControlTypeIds =
    [
        50007, // ListItem
        50010, // MenuItem
        50020, // Text
        50029, // DataItem
        50025, // TreeItem
        50002, // CheckBox
        50013  // RadioButton
    ];

    private static readonly int[] ContainerControlTypeIds =
    [
        50008, // List
        50009, // Menu
        50024, // Tree
        50033, // Pane
        50032, // Window
        50028, // DataGrid
        50003  // ComboBox
    ];

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

    public object SelectComboBoxNativeUia(UiRequest request, IntPtr? activeWindowHwnd, int? processId)
    {
        var operation = string.IsNullOrWhiteSpace(request.Operation) ? "selectcomboboxuia" : request.Operation;
        var timeoutMs = request.TimeoutMs ?? DefaultTimeoutMs;
        var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);
        var diagnostics = new List<object>();

        var resolveResult = _resolver.ResolveComboBox(request, activeWindowHwnd, processId);
        if (resolveResult.Element == null)
        {
            return BuildFailureResponse(
                operation,
                request,
                stage: resolveResult.Stage ?? "combo-not-found",
                lastError: resolveResult.LastError ?? "ComboBox not found.",
                diagnostics,
                candidateContainers: resolveResult.Candidates);
        }

        if (resolveResult.IsAmbiguous)
        {
            return new Dictionary<string, object?>
            {
                ["operation"] = operation,
                ["success"] = false,
                ["stage"] = "ambiguous-combobox",
                ["requested"] = request.Value,
                ["lastError"] = resolveResult.LastError,
                ["candidateContainers"] = resolveResult.Candidates,
                ["diagnostics"] = diagnostics
            };
        }

        var combo = resolveResult.Element;
        var comboSnapshot = _uia.CreateSnapshot(combo);
        var beforeValue = _textReader.ReadComboBoxValue(combo);

        if (!_uia.GetBoolProperty(combo, UIA_PropertyIds.UIA_IsEnabledPropertyId, true))
        {
            return BuildFailureResponse(operation, request, "combo-disabled", "ComboBox is disabled.", diagnostics, comboSnapshot);
        }

        if (_uia.GetBoolProperty(combo, UIA_PropertyIds.UIA_IsOffscreenPropertyId))
        {
            return BuildFailureResponse(operation, request, "combo-offscreen", "ComboBox is offscreen.", diagnostics, comboSnapshot);
        }

        string? openStrategy = null;
        string? itemMatchStrategy = null;
        string? selectionStrategy = null;
        IUIAutomationElement? matchedItem = null;
        string? requestedValue = request.Index.HasValue && string.IsNullOrWhiteSpace(request.Value)
            ? null
            : WebUtility.HtmlDecode(request.Value ?? string.Empty).Trim();

        if (request.Index.HasValue && string.IsNullOrWhiteSpace(requestedValue))
        {
            matchedItem = SelectByIndex(combo, request.Index.Value, diagnostics, out selectionStrategy);
            requestedValue = matchedItem == null ? $"index:{request.Index.Value}" : _textReader.ReadElementText(matchedItem);
            itemMatchStrategy = "index";
        }
        else
        {
            if (string.IsNullOrWhiteSpace(requestedValue))
            {
                return BuildFailureResponse(operation, request, "missing-value", "Either value or index is required.", diagnostics, comboSnapshot);
            }

            if (TrySetEditableValue(combo, requestedValue, diagnostics, out selectionStrategy))
            {
                var directActual = _textReader.ReadComboBoxValue(combo);
                if (NativeUiaText.ValuesEquivalent(directActual, requestedValue))
                {
                    return BuildSuccessResponse(
                        operation,
                        requestedValue,
                        directActual,
                        verified: true,
                        openStrategy: "ValuePattern.SetValue",
                        itemMatchStrategy: "direct-value",
                        selectionStrategy: selectionStrategy ?? "ValuePattern.SetValue",
                        beforeValue,
                        comboSnapshot,
                        null,
                        null,
                        [],
                        diagnostics);
                }
            }

            openStrategy = ExpandComboBox(combo, diagnostics);
            Thread.Sleep(ExpandDelayMs);

            matchedItem = FindDropdownItem(
                combo,
                activeWindowHwnd,
                processId,
                requestedValue,
                request,
                diagnostics,
                deadline,
                out itemMatchStrategy);

            if (matchedItem != null)
                selectionStrategy = ActivateItem(combo, matchedItem, diagnostics, out _);

            if (matchedItem == null || selectionStrategy == null)
            {
                if (request.AllowKeyboardFallback == true
                    && DateTime.UtcNow < deadline
                    && TryKeyboardTypeahead(combo, requestedValue, diagnostics))
                {
                    selectionStrategy = "native-uia-keyboard-typeahead";
                }
                else if (DateTime.UtcNow < deadline
                         && TryPagedVisibleSearch(
                             combo,
                             activeWindowHwnd,
                             processId,
                             requestedValue,
                             request,
                             diagnostics,
                             deadline,
                             out matchedItem,
                             out selectionStrategy))
                {
                    itemMatchStrategy ??= "paged-visible-search";
                }
            }
        }

        var actual = _textReader.ReadComboBoxValue(combo);
        var verified = !string.IsNullOrWhiteSpace(requestedValue)
                       && NativeUiaText.ValuesEquivalent(actual, requestedValue);
        if (request.Index.HasValue && string.IsNullOrWhiteSpace(request.Value))
            verified = matchedItem != null;

        var container = _dropdownFinder.FindBestContainer(combo, activeWindowHwnd, processId);
        var containerSnapshot = container == null ? null : _uia.CreateSnapshot(container);
        var visibleItems = _dropdownFinder.CollectVisibleItemTexts(combo, activeWindowHwnd, processId, 25);

        if (!verified && string.IsNullOrWhiteSpace(selectionStrategy))
        {
            return BuildFailureResponse(
                operation,
                request,
                stage: matchedItem == null ? "item-not-found" : "selection-not-verified",
                lastError: matchedItem == null
                    ? $"Item '{requestedValue}' was not found in dropdown."
                    : $"Selection did not verify. actual='{actual}'.",
                diagnostics,
                comboSnapshot,
                openStrategy,
                containerSnapshot,
                visibleItems,
                matchedItem);
        }

        return BuildSuccessResponse(
            operation,
            requestedValue,
            actual,
            verified,
            openStrategy,
            itemMatchStrategy,
            selectionStrategy ?? "native-uia-unverified",
            beforeValue,
            comboSnapshot,
            containerSnapshot,
            matchedItem == null ? null : _uia.CreateSnapshot(matchedItem),
            visibleItems,
            diagnostics);
    }

    public object InspectComboBoxNativeUia(UiRequest request, IntPtr? activeWindowHwnd, int? processId)
    {
        var operation = "inspectcomboboxuia";
        var resolveResult = _resolver.ResolveComboBox(request, activeWindowHwnd, processId);

        if (resolveResult.Element == null)
        {
            return new Dictionary<string, object?>
            {
                ["operation"] = operation,
                ["success"] = false,
                ["stage"] = resolveResult.Stage ?? "combo-not-found",
                ["lastError"] = resolveResult.LastError,
                ["candidates"] = resolveResult.Candidates
            };
        }

        var combo = resolveResult.Element;
        var comboSnapshot = _uia.CreateSnapshot(combo);
        var container = _dropdownFinder.FindBestContainer(combo, activeWindowHwnd, processId);
        var containerSnapshot = container == null ? null : _uia.CreateSnapshot(container);
        var visibleItems = _dropdownFinder.CollectVisibleItemTexts(combo, activeWindowHwnd, processId, 50);
        var itemCandidates = _dropdownFinder.CollectItemDiagnostics(combo, activeWindowHwnd, processId, 50);

        return new Dictionary<string, object?>
        {
            ["operation"] = operation,
            ["success"] = true,
            ["combo"] = NativeUiaDiagnostics.ComboSummary(comboSnapshot),
            ["currentValue"] = _textReader.ReadComboBoxValue(combo),
            ["expandState"] = GetExpandState(combo),
            ["supportedPatterns"] = comboSnapshot.SupportedPatterns,
            ["dropdownCandidates"] = containerSnapshot == null ? null : NativeUiaDiagnostics.ComboSummary(containerSnapshot),
            ["visibleItems"] = visibleItems,
            ["itemCandidates"] = itemCandidates,
            ["ambiguous"] = resolveResult.IsAmbiguous,
            ["candidates"] = resolveResult.Candidates
        };
    }

    private static object BuildSuccessResponse(
        string operation,
        string? requested,
        string actual,
        bool verified,
        string? openStrategy,
        string? itemMatchStrategy,
        string selectionStrategy,
        string beforeValue,
        NativeUiaElementSnapshot comboSnapshot,
        NativeUiaElementSnapshot? containerSnapshot,
        NativeUiaElementSnapshot? matchedItemSnapshot,
        List<string>? visibleItems,
        List<object> diagnostics)
    {
        return new Dictionary<string, object?>
        {
            ["operation"] = operation,
            ["success"] = true,
            ["requested"] = requested,
            ["actual"] = actual,
            ["verified"] = verified,
            ["strategy"] = selectionStrategy,
            ["openStrategy"] = openStrategy,
            ["itemMatchStrategy"] = itemMatchStrategy,
            ["selectionStrategy"] = selectionStrategy,
            ["beforeValue"] = beforeValue,
            ["afterValue"] = actual,
            ["combo"] = NativeUiaDiagnostics.ComboSummary(comboSnapshot),
            ["comboBox"] = NativeUiaDiagnostics.ComboSummary(comboSnapshot),
            ["dropdown"] = containerSnapshot == null
                ? new { found = false }
                : new
                {
                    found = true,
                    controlType = containerSnapshot.ControlType,
                    rectangle = NativeUiaDiagnostics.ToRectangleObject(containerSnapshot.BoundingRectangle)
                },
            ["matchedItem"] = matchedItemSnapshot == null
                ? null
                : NativeUiaDiagnostics.CandidateDiagnostic(0, matchedItemSnapshot),
            ["visibleItemsSample"] = visibleItems ?? [],
            ["diagnostics"] = diagnostics
        };
    }

    private object BuildFailureResponse(
        string operation,
        UiRequest request,
        string stage,
        string lastError,
        List<object> diagnostics,
        NativeUiaElementSnapshot? comboSnapshot = null,
        string? openStrategy = null,
        NativeUiaElementSnapshot? containerSnapshot = null,
        List<string>? visibleItems = null,
        IUIAutomationElement? matchedItem = null,
        List<object>? candidateContainers = null)
    {
        return new Dictionary<string, object?>
        {
            ["operation"] = operation,
            ["success"] = false,
            ["requested"] = request.Value ?? (request.Index.HasValue ? $"index:{request.Index}" : null),
            ["stage"] = stage,
            ["openStrategy"] = openStrategy,
            ["dropdownFound"] = containerSnapshot != null,
            ["visibleItemsSample"] = visibleItems ?? [],
            ["candidateContainers"] = candidateContainers ?? [],
            ["candidateItems"] = matchedItem == null ? [] : diagnostics,
            ["combo"] = comboSnapshot == null ? null : NativeUiaDiagnostics.ComboSummary(comboSnapshot),
            ["lastError"] = lastError,
            ["diagnostics"] = diagnostics
        };
    }

    private bool TrySetEditableValue(IUIAutomationElement combo, string requestedValue, List<object> diagnostics, out string? strategy)
    {
        strategy = null;
        try
        {
            if (combo.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId) is not IUIAutomationValuePattern valuePattern)
                return false;

            if (valuePattern.CurrentIsReadOnly != 0)
                return false;

            valuePattern.SetValue(requestedValue);
            Thread.Sleep(ActionDelayMs);
            strategy = "native-uia-valuepattern-setvalue";
            diagnostics.Add(Attempt(strategy, true, null, _textReader.ReadComboBoxValue(combo)));
            return true;
        }
        catch (Exception ex)
        {
            diagnostics.Add(Attempt("native-uia-valuepattern-setvalue", false, ex.Message, _textReader.ReadComboBoxValue(combo)));
            return false;
        }
    }

    private string? GetExpandState(IUIAutomationElement combo)
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

    private string GetComboBoxCurrentValue(IUIAutomationElement combo) =>
        _textReader.ReadComboBoxValue(combo);

    private string ExpandComboBox(IUIAutomationElement combo, List<object> attempts)
    {
        if (_uia.TryGetExpandCollapsePattern(combo, out var expandCollapse))
        {
            try
            {
                expandCollapse!.Expand();
                attempts.Add(Attempt("ExpandCollapsePattern.Expand", true, null, GetComboBoxCurrentValue(combo)));
                return "ExpandCollapsePattern.Expand";
            }
            catch (Exception ex)
            {
                attempts.Add(Attempt("ExpandCollapsePattern.Expand", false, ex.Message, GetComboBoxCurrentValue(combo)));
            }
        }

        if (_uia.TryGetInvokePattern(combo, out var invoke))
        {
            try
            {
                invoke!.Invoke();
                attempts.Add(Attempt("InvokePattern.Invoke", true, null, GetComboBoxCurrentValue(combo)));
                return "InvokePattern.Invoke";
            }
            catch (Exception ex)
            {
                attempts.Add(Attempt("InvokePattern.Invoke", false, ex.Message, GetComboBoxCurrentValue(combo)));
            }
        }

        var rect = _uia.GetBoundingRectangle(combo);
        if (rect.HasValue)
        {
            var clickPoint = ComboBoxRightEdgePoint(rect.Value);
            var clicked = NativeUiaInput.ClickPoint(clickPoint);
            attempts.Add(Attempt("physical-click-right-edge", clicked, clicked ? null : "click failed", GetComboBoxCurrentValue(combo)));
            if (clicked)
                return "physical-click-right-edge";
        }

        FocusComboBox(combo);
        NativeUiaInput.SendChord(NativeUiaInput.VirtualKeys.Menu, NativeUiaInput.VirtualKeys.Down);
        Thread.Sleep(ActionDelayMs);
        attempts.Add(Attempt("Alt+Down", true, null, GetComboBoxCurrentValue(combo)));

        FocusComboBox(combo);
        NativeUiaInput.SendKey(NativeUiaInput.VirtualKeys.F4);
        Thread.Sleep(ActionDelayMs);
        attempts.Add(Attempt("F4", true, null, GetComboBoxCurrentValue(combo)));
        return "F4";
    }

    private IUIAutomationElement? FindDropdownItem(
        IUIAutomationElement combo,
        IntPtr? activeWindowHwnd,
        int? processId,
        string requestedValue,
        UiRequest request,
        List<object> attempts,
        DateTime deadline,
        out string? itemMatchStrategy)
    {
        itemMatchStrategy = null;
        foreach (var candidate in _dropdownFinder.EnumerateItemCandidates(combo, activeWindowHwnd, processId))
        {
            if (DateTime.UtcNow >= deadline)
                break;

            var snapshot = _uia.CreateSnapshot(candidate);
            var texts = _textReader.ReadCandidateTexts(candidate);
            var isMatch = texts.Any(t => ItemMatches(t, requestedValue, request));

            attempts.Add(new Dictionary<string, object?>
            {
                ["candidate"] = snapshot.ToDiagnosticObject(),
                ["texts"] = texts,
                ["match"] = isMatch
            });

            if (isMatch)
            {
                itemMatchStrategy = texts.Any(t => NativeUiaText.Matches(t, requestedValue, "exact"))
                    ? "exact-name"
                    : "contains-name";
                return candidate;
            }
        }

        return null;
    }

    private IEnumerable<IUIAutomationElement> EnumerateItemCandidates(
        IUIAutomationElement combo,
        IntPtr? activeWindowHwnd,
        int? processId,
        UiRequest request) =>
        _dropdownFinder.EnumerateItemCandidates(combo, activeWindowHwnd, processId);

    private List<IUIAutomationElement> BuildDropdownSearchRoots(
        IUIAutomationElement combo,
        IntPtr? activeWindowHwnd,
        int? processId) =>
        _dropdownFinder.BuildSearchRoots(combo, activeWindowHwnd, processId);

    private static bool ItemMatches(string candidate, string requested, UiRequest request)
    {
        var matchMode = string.IsNullOrWhiteSpace(request.MatchMode) ? "exact" : request.MatchMode!;
        return NativeUiaText.Matches(candidate, requested, matchMode);
    }

    private string? ActivateItem(
        IUIAutomationElement combo,
        IUIAutomationElement item,
        List<object> attempts,
        out string strategy)
    {
        strategy = "native-uia-unverified";
        var before = GetComboBoxCurrentValue(combo);

        if (_uia.TryGetSelectionItemPattern(item, out var selectionItem))
        {
            try
            {
                selectionItem!.Select();
                Thread.Sleep(ActionDelayMs);
                var after = GetComboBoxCurrentValue(combo);
                var success = NativeUiaText.ValuesEquivalent(after, _uia.GetElementText(item));
                attempts.Add(Attempt("native-uia-selectionitem", success, null, after, before, success));
                if (success)
                {
                    strategy = "native-uia-selectionitem";
                    return strategy;
                }
            }
            catch (Exception ex)
            {
                attempts.Add(Attempt("native-uia-selectionitem", false, ex.Message, GetComboBoxCurrentValue(combo), before, false));
            }
        }

        if (_uia.TryGetInvokePattern(item, out var invoke))
        {
            try
            {
                invoke!.Invoke();
                Thread.Sleep(ActionDelayMs);
                var after = GetComboBoxCurrentValue(combo);
                var success = NativeUiaText.ValuesEquivalent(after, _uia.GetElementText(item));
                attempts.Add(Attempt("native-uia-invoke", success, null, after, before, success));
                if (success)
                {
                    strategy = "native-uia-invoke";
                    return strategy;
                }
            }
            catch (Exception ex)
            {
                attempts.Add(Attempt("native-uia-invoke", false, ex.Message, GetComboBoxCurrentValue(combo), before, false));
            }
        }

        var controlType = _uia.GetIntProperty(item, UIA_PropertyIds.UIA_ControlTypePropertyId);
        if (controlType is 50002 or 50013 && _uia.TryGetTogglePattern(item, out var toggle))
        {
            try
            {
                toggle!.Toggle();
                Thread.Sleep(ActionDelayMs);
                var after = GetComboBoxCurrentValue(combo);
                attempts.Add(Attempt("native-uia-toggle", true, null, after, before, true));
                strategy = "native-uia-toggle";
                return strategy;
            }
            catch (Exception ex)
            {
                attempts.Add(Attempt("native-uia-toggle", false, ex.Message, GetComboBoxCurrentValue(combo), before, false));
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
                var after = GetComboBoxCurrentValue(combo);
                var success = NativeUiaText.ValuesEquivalent(after, _uia.GetElementText(item));
                attempts.Add(Attempt("native-uia-physical-click", success, null, after, before, success));
                if (success)
                {
                    strategy = "native-uia-physical-click";
                    return strategy;
                }
            }

            if (NativeUiaInput.ClickPoint(leftCenter))
            {
                Thread.Sleep(ActionDelayMs);
                var after = GetComboBoxCurrentValue(combo);
                var success = NativeUiaText.ValuesEquivalent(after, _uia.GetElementText(item));
                attempts.Add(Attempt("native-uia-physical-click-left", success, null, after, before, success));
                if (success)
                {
                    strategy = "native-uia-physical-click-left";
                    return strategy;
                }
            }
        }

        try
        {
            item.SetFocus();
            Thread.Sleep(50);
            NativeUiaInput.SendKey(NativeUiaInput.VirtualKeys.Return);
            Thread.Sleep(ActionDelayMs);
            var afterEnter = GetComboBoxCurrentValue(combo);
            attempts.Add(Attempt("native-uia-focus-enter", true, null, afterEnter, before, true));
            strategy = "native-uia-focus-enter";
            return strategy;
        }
        catch (Exception ex)
        {
            attempts.Add(Attempt("native-uia-focus-enter", false, ex.Message, GetComboBoxCurrentValue(combo), before, false));
        }

        try
        {
            item.SetFocus();
            NativeUiaInput.SendKey(NativeUiaInput.VirtualKeys.Space);
            Thread.Sleep(ActionDelayMs);
            var afterSpace = GetComboBoxCurrentValue(combo);
            attempts.Add(Attempt("native-uia-focus-space", true, null, afterSpace, before, true));
            strategy = "native-uia-focus-space";
            return strategy;
        }
        catch (Exception ex)
        {
            attempts.Add(Attempt("native-uia-focus-space", false, ex.Message, GetComboBoxCurrentValue(combo), before, false));
        }

        return null;
    }

    private bool TryKeyboardTypeahead(IUIAutomationElement combo, string requestedValue, List<object> attempts)
    {
        try
        {
            FocusComboBox(combo);
            NativeUiaInput.TypeText(requestedValue);
            Thread.Sleep(ActionDelayMs);
            NativeUiaInput.SendKey(NativeUiaInput.VirtualKeys.Return);
            Thread.Sleep(ActionDelayMs);
            var after = GetComboBoxCurrentValue(combo);
            var success = NativeUiaText.ValuesEquivalent(after, requestedValue);
            attempts.Add(Attempt("native-uia-keyboard-typeahead", success, null, after, null, success));
            return success;
        }
        catch (Exception ex)
        {
            attempts.Add(Attempt("native-uia-keyboard-typeahead", false, ex.Message, GetComboBoxCurrentValue(combo)));
            return false;
        }
    }

    private bool TryPagedVisibleSearch(
        IUIAutomationElement combo,
        IntPtr? activeWindowHwnd,
        int? processId,
        string requestedValue,
        UiRequest request,
        List<object> attempts,
        DateTime deadline,
        out IUIAutomationElement? matchedItem,
        out string strategy)
    {
        matchedItem = null;
        strategy = "native-uia-paged-visible-search";
        string? previousBatch = null;

        for (var page = 0; page < PagedSearchMaxPages && DateTime.UtcNow < deadline; page++)
        {
            var visibleItems = _dropdownFinder.CollectVisibleItemTexts(combo, activeWindowHwnd, processId, PagedSearchVisibleLimit);
            var batch = string.Join("|", visibleItems);
            attempts.Add(new Dictionary<string, object?> { ["pagedBatch"] = batch, ["page"] = page });

            matchedItem = FindDropdownItem(combo, activeWindowHwnd, processId, requestedValue, request, attempts, deadline, out _);
            if (matchedItem != null)
            {
                var activation = ActivateItem(combo, matchedItem, attempts, out strategy);
                if (activation != null)
                    return true;
            }

            if (string.Equals(batch, previousBatch, StringComparison.Ordinal))
                break;

            previousBatch = batch;
            NativeUiaInput.WheelDown(5);
            Thread.Sleep(ActionDelayMs);
            NativeUiaInput.SendKey(NativeUiaInput.VirtualKeys.Next);
            Thread.Sleep(ActionDelayMs);
        }

        matchedItem = null;
        return false;
    }

    private List<string> CollectVisibleItemTexts(
        IUIAutomationElement combo,
        IntPtr? activeWindowHwnd,
        int? processId,
        UiRequest request,
        int limit) =>
        _dropdownFinder.CollectVisibleItemTexts(combo, activeWindowHwnd, processId, limit);

    private IUIAutomationElement? SelectByIndex(
        IUIAutomationElement combo,
        int index,
        List<object> attempts,
        out string? strategy)
    {
        strategy = null;
        if (index < 0)
            throw new ArgumentOutOfRangeException(nameof(index), "ComboBox index must be >= 0.");

        ExpandComboBox(combo, attempts);
        Thread.Sleep(ExpandDelayMs);

        var items = _dropdownFinder.EnumerateItemCandidates(combo, null, null).ToList();
        if (index >= items.Count)
            return null;

        var item = items[index];
        strategy = ActivateItem(combo, item, attempts, out var activationStrategy);
        return strategy == null ? null : item;
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

    private bool IsExpanded(IUIAutomationElement combo)
    {
        if (!_uia.TryGetExpandCollapsePattern(combo, out var pattern) || pattern == null)
            return false;

        try
        {
            return pattern.CurrentExpandCollapseState == ExpandCollapseState.ExpandCollapseState_Expanded;
        }
        catch
        {
            return false;
        }
    }

    private static Dictionary<string, object?> Attempt(
        string strategy,
        bool success,
        string? error,
        string? afterValue,
        string? beforeValue = null,
        bool? verified = null)
    {
        return new Dictionary<string, object?>
        {
            ["strategy"] = strategy,
            ["success"] = success,
            ["error"] = error,
            ["beforeValue"] = beforeValue,
            ["afterValue"] = afterValue,
            ["verified"] = verified
        };
    }

    private static bool SameElement(IUIAutomationElement left, IUIAutomationElement right)
    {
        return string.Equals(RuntimeKey(left), RuntimeKey(right), StringComparison.Ordinal);
    }

    private static string RuntimeKey(IUIAutomationElement element)
    {
        try
        {
            var runtimeId = element.GetRuntimeId();
            return runtimeId == null ? element.GetHashCode().ToString() : string.Join(".", runtimeId);
        }
        catch
        {
            return Guid.NewGuid().ToString("N");
        }
    }
}
