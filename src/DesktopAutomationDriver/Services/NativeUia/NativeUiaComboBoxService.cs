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
    private readonly NativeUiaElementFinder _finder;
    private readonly ILogger<NativeUiaComboBoxService> _logger;

    public NativeUiaComboBoxService(ILogger<NativeUiaComboBoxService> logger)
        : this(new NativeUiaAutomation(), logger)
    {
    }

    internal NativeUiaComboBoxService(NativeUiaAutomation uia, ILogger<NativeUiaComboBoxService> logger)
    {
        _uia = uia;
        _finder = new NativeUiaElementFinder(_uia);
        _logger = logger;
    }

    public object SelectComboBox(UiRequest request, IntPtr? activeWindowHwnd, int? processId)
    {
        var operation = string.IsNullOrWhiteSpace(request.Operation) ? "select" : request.Operation;
        var timeoutMs = request.TimeoutMs ?? DefaultTimeoutMs;
        var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);
        var attempts = new List<object>();

        var requestedValue = request.Index.HasValue && string.IsNullOrWhiteSpace(request.Value)
            ? null
            : WebUtility.HtmlDecode(request.Value ?? string.Empty).Trim();

        var combo = ResolveComboBox(request, activeWindowHwnd, processId);
        if (combo == null)
            throw new InvalidOperationException("Native UIA ComboBox resolver could not find a ComboBox for the locator.");

        var comboSnapshot = _uia.CreateSnapshot(combo);
        var beforeValue = GetComboBoxCurrentValue(combo);

        IUIAutomationElement? matchedItem = null;
        string? strategy = null;

        if (request.Index.HasValue && string.IsNullOrWhiteSpace(requestedValue))
        {
            matchedItem = SelectByIndex(combo, request.Index.Value, attempts, out strategy);
            requestedValue = matchedItem == null ? $"index:{request.Index.Value}" : _uia.GetElementText(matchedItem);
        }
        else
        {
            if (string.IsNullOrWhiteSpace(requestedValue))
                throw new ArgumentException("Either 'value' or 'index' is required for ComboBox select.");

            if (!ExpandComboBox(combo, attempts))
            {
                _logger.LogWarning(
                    "Native UIA ComboBox expand did not report success; continuing with item search.");
            }

            Thread.Sleep(ExpandDelayMs);

            matchedItem = FindDropdownItem(
                combo,
                activeWindowHwnd,
                processId,
                requestedValue,
                request,
                attempts,
                deadline);

            if (matchedItem != null)
            {
                strategy = ActivateItem(combo, matchedItem, attempts, out _);
            }

            if (matchedItem == null || strategy == null)
            {
                if (request.AllowKeyboardFallback == true
                    && DateTime.UtcNow < deadline
                    && TryKeyboardTypeahead(combo, requestedValue, attempts))
                {
                    strategy = "native-uia-keyboard-typeahead";
                }
                else if (DateTime.UtcNow < deadline
                         && TryPagedVisibleSearch(
                             combo,
                             activeWindowHwnd,
                             processId,
                             requestedValue,
                             request,
                             attempts,
                             deadline,
                             out matchedItem,
                             out var pagedStrategy))
                {
                    strategy = pagedStrategy;
                }
            }
        }

        var actual = GetComboBoxCurrentValue(combo);
        var verified = !string.IsNullOrWhiteSpace(requestedValue)
                       && NativeUiaText.ValuesEquivalent(actual, requestedValue);

        if (request.Index.HasValue && string.IsNullOrWhiteSpace(request.Value))
            verified = matchedItem != null;

        var success = verified || (!string.IsNullOrWhiteSpace(strategy) && matchedItem != null);

        var response = new Dictionary<string, object?>
        {
            ["operation"] = operation,
            ["success"] = success,
            ["strategy"] = strategy ?? "native-uia-unverified",
            ["requested"] = requestedValue,
            ["actual"] = actual,
            ["verified"] = verified,
            ["comboBox"] = comboSnapshot.ToDiagnosticObject(),
            ["item"] = matchedItem == null ? null : _uia.CreateSnapshot(matchedItem).ToDiagnosticObject(),
            ["attempts"] = attempts,
            ["dropdown"] = new
            {
                beforeValue,
                afterValue = actual,
                expanded = IsExpanded(combo)
            }
        };

        if (!success)
        {
            response["visibleItems"] = CollectVisibleItemTexts(combo, activeWindowHwnd, processId, request, 50);
            response["dropdownCandidates"] = attempts
                .Where(a => a is Dictionary<string, object?> d && d.ContainsKey("candidate"))
                .Take(25)
                .ToList();
        }

        if (!success)
        {
            throw new InvalidOperationException(
                $"Native UIA ComboBox selection failed. requested='{requestedValue}', actual='{actual}'.");
        }

        return response;
    }

    private IUIAutomationElement? ResolveComboBox(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId)
    {
        var locator = ToNativeLocator(request, processId);
        var searchRoots = BuildSearchRoots(activeWindowHwnd, processId);

        foreach (var root in searchRoots)
        {
            var found = _finder.FindFirst(root, locator.AsComboBoxLocator());
            if (found != null)
                return found;

            if (!string.IsNullOrWhiteSpace(locator.AutomationId) || !string.IsNullOrWhiteSpace(locator.Name))
            {
                var relaxedFound = _finder.FindFirst(root, locator.WithoutControlType());
                var ancestor = relaxedFound == null ? null : PromoteToComboBox(relaxedFound);
                if (ancestor != null)
                    return ancestor;
            }
        }

        return null;
    }

    private IUIAutomationElement? PromoteToComboBox(IUIAutomationElement element)
    {
        var controlType = _uia.GetIntProperty(element, UIA_PropertyIds.UIA_ControlTypePropertyId);
        if (controlType == ComboBoxControlTypeId)
            return element;

        if (controlType == EditControlTypeId)
        {
            return _uia.WalkAncestor(
                element,
                e => _uia.GetIntProperty(e, UIA_PropertyIds.UIA_ControlTypePropertyId) == ComboBoxControlTypeId);
        }

        return null;
    }

    private List<IUIAutomationElement> BuildSearchRoots(IntPtr? activeWindowHwnd, int? processId)
    {
        var roots = new List<IUIAutomationElement>();

        if (activeWindowHwnd is > 0)
        {
            var activeRoot = _uia.FromHandle(activeWindowHwnd.Value);
            if (activeRoot != null)
                roots.Add(activeRoot);
        }

        var foreground = NativeUiaInput.ForegroundWindowHandle();
        if (foreground != IntPtr.Zero)
        {
            var fgRoot = _uia.FromHandle(foreground);
            if (fgRoot != null && !roots.Any(r => SameElement(r, fgRoot)))
                roots.Add(fgRoot);
        }

        roots.Add(_uia.Root);
        return roots;
    }

    private static NativeUiaLocator ToNativeLocator(UiRequest request, int? processId)
    {
        return new NativeUiaLocator
        {
            Name = request.Locator?.Name,
            AutomationId = request.Locator?.AutomationId,
            ClassName = request.Locator?.ClassName ?? request.ClassName,
            ControlType = request.Locator?.ControlType,
            Value = null,
            Hwnd = request.Hwnd ?? request.Locator?.Hwnd,
            ProcessId = request.ProcessId ?? request.Locator?.ProcessId ?? processId,
            FoundIndex = request.FoundIndex,
            MatchMode = string.IsNullOrWhiteSpace(request.MatchMode) ? "exact" : request.MatchMode!
        };
    }

    private string GetComboBoxCurrentValue(IUIAutomationElement combo)
    {
        var valuePattern = _uia.GetValuePatternText(combo);
        if (!string.IsNullOrWhiteSpace(valuePattern))
            return valuePattern;

        var legacyValue = _uia.GetLegacyAccessibleValue(combo);
        if (!string.IsNullOrWhiteSpace(legacyValue))
            return legacyValue;

        var name = _uia.GetStringProperty(combo, UIA_PropertyIds.UIA_NamePropertyId);
        if (!string.IsNullOrWhiteSpace(name))
            return name;

        var selectedItem = FindSelectedItem(combo);
        if (selectedItem != null)
            return _uia.GetElementText(selectedItem);

        var innerEdit = _uia.FindFirstDescendant(
            combo,
            _uia.PropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId, EditControlTypeId));
        if (innerEdit != null)
        {
            var editValue = _uia.GetValuePatternText(innerEdit);
            if (!string.IsNullOrWhiteSpace(editValue))
                return editValue;
        }

        return _uia.GetElementText(combo);
    }

    private IUIAutomationElement? FindSelectedItem(IUIAutomationElement combo)
    {
        var items = _uia.FindAllDescendants(
            combo,
            _uia.PropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId, 50007),
            200);

        foreach (var item in items)
        {
            if (_uia.TryGetSelectionItemPattern(item, out var selection) && selection!.CurrentIsSelected != 0)
                return item;
        }

        return null;
    }

    private bool ExpandComboBox(IUIAutomationElement combo, List<object> attempts)
    {
        if (_uia.TryGetExpandCollapsePattern(combo, out var expandCollapse))
        {
            try
            {
                expandCollapse!.Expand();
                attempts.Add(Attempt("native-uia-expandcollapse", true, null, GetComboBoxCurrentValue(combo)));
                return true;
            }
            catch (Exception ex)
            {
                attempts.Add(Attempt("native-uia-expandcollapse", false, ex.Message, GetComboBoxCurrentValue(combo)));
            }
        }

        if (_uia.TryGetInvokePattern(combo, out var invoke))
        {
            try
            {
                invoke!.Invoke();
                attempts.Add(Attempt("native-uia-invoke-expand", true, null, GetComboBoxCurrentValue(combo)));
                return true;
            }
            catch (Exception ex)
            {
                attempts.Add(Attempt("native-uia-invoke-expand", false, ex.Message, GetComboBoxCurrentValue(combo)));
            }
        }

        var rect = _uia.GetBoundingRectangle(combo);
        if (rect.HasValue)
        {
            var clickPoint = ComboBoxRightEdgePoint(rect.Value);
            var clicked = NativeUiaInput.ClickPoint(clickPoint);
            attempts.Add(Attempt("native-uia-physical-expand", clicked, clicked ? null : "click failed", GetComboBoxCurrentValue(combo)));
            if (clicked)
                return true;
        }

        FocusComboBox(combo);
        NativeUiaInput.SendChord(NativeUiaInput.VirtualKeys.Menu, NativeUiaInput.VirtualKeys.Down);
        Thread.Sleep(ActionDelayMs);
        attempts.Add(Attempt("native-uia-alt-down-expand", true, null, GetComboBoxCurrentValue(combo)));

        FocusComboBox(combo);
        NativeUiaInput.SendKey(NativeUiaInput.VirtualKeys.F4);
        Thread.Sleep(ActionDelayMs);
        attempts.Add(Attempt("native-uia-f4-expand", true, null, GetComboBoxCurrentValue(combo)));
        return true;
    }

    private IUIAutomationElement? FindDropdownItem(
        IUIAutomationElement combo,
        IntPtr? activeWindowHwnd,
        int? processId,
        string requestedValue,
        UiRequest request,
        List<object> attempts,
        DateTime deadline)
    {
        foreach (var candidate in EnumerateItemCandidates(combo, activeWindowHwnd, processId, request))
        {
            if (DateTime.UtcNow >= deadline)
                break;

            var snapshot = _uia.CreateSnapshot(candidate);
            var texts = CandidateTexts(candidate);
            var isMatch = texts.Any(t => ItemMatches(t, requestedValue, request));

            attempts.Add(new Dictionary<string, object?>
            {
                ["candidate"] = snapshot.ToDiagnosticObject(),
                ["texts"] = texts,
                ["match"] = isMatch
            });

            if (isMatch)
                return candidate;
        }

        return null;
    }

    private IEnumerable<IUIAutomationElement> EnumerateItemCandidates(
        IUIAutomationElement combo,
        IntPtr? activeWindowHwnd,
        int? processId,
        UiRequest request)
    {
        var seen = new HashSet<string>();
        foreach (var root in BuildDropdownSearchRoots(combo, activeWindowHwnd, processId))
        {
            foreach (var typeId in ItemControlTypeIds)
            {
                foreach (var element in _uia.FindAllDescendants(
                             root,
                             _uia.PropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId, typeId),
                             500))
                {
                    var key = RuntimeKey(element);
                    if (!seen.Add(key))
                        continue;

                    yield return element;
                }
            }
        }
    }

    private List<IUIAutomationElement> BuildDropdownSearchRoots(
        IUIAutomationElement combo,
        IntPtr? activeWindowHwnd,
        int? processId)
    {
        var roots = new List<IUIAutomationElement> { combo };

        if (activeWindowHwnd is > 0)
        {
            var activeRoot = _uia.FromHandle(activeWindowHwnd.Value);
            if (activeRoot != null)
                roots.Add(activeRoot);
        }

        var desktopChildren = _uia.FindAllChildren(_uia.Root, _uia.TrueCondition(), 200);
        roots.AddRange(desktopChildren);

        if (processId.HasValue)
        {
            var sameProcess = _uia.FindAllDescendants(
                _uia.Root,
                _uia.PropertyCondition(UIA_PropertyIds.UIA_ProcessIdPropertyId, processId.Value),
                500);
            roots.AddRange(sameProcess.Where(e =>
                ContainerControlTypeIds.Contains(_uia.GetIntProperty(e, UIA_PropertyIds.UIA_ControlTypePropertyId))));
        }

        var foreground = NativeUiaInput.ForegroundWindowHandle();
        if (foreground != IntPtr.Zero)
        {
            var fgRoot = _uia.FromHandle(foreground);
            if (fgRoot != null)
                roots.Add(fgRoot);
        }

        return roots.DistinctBy(RuntimeKey).ToList();
    }

    private List<string> CandidateTexts(IUIAutomationElement element)
    {
        return new List<string>
        {
            _uia.GetStringProperty(element, UIA_PropertyIds.UIA_NamePropertyId),
            _uia.GetValuePatternText(element),
            _uia.GetTextPatternText(element),
            _uia.GetLegacyAccessibleName(element),
            _uia.GetLegacyAccessibleValue(element),
            _uia.GetStringProperty(element, UIA_PropertyIds.UIA_AutomationIdPropertyId)
        }.Where(t => !string.IsNullOrWhiteSpace(t)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    }

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
            var visibleItems = CollectVisibleItemTexts(combo, activeWindowHwnd, processId, request, PagedSearchVisibleLimit);
            var batch = string.Join("|", visibleItems);
            attempts.Add(new Dictionary<string, object?> { ["pagedBatch"] = batch, ["page"] = page });

            matchedItem = FindDropdownItem(combo, activeWindowHwnd, processId, requestedValue, request, attempts, deadline);
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
        int limit)
    {
        return EnumerateItemCandidates(combo, activeWindowHwnd, processId, request)
            .Select(_uia.GetElementText)
            .Where(t => !string.IsNullOrWhiteSpace(t))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(limit)
            .ToList();
    }

    private IUIAutomationElement? SelectByIndex(
        IUIAutomationElement combo,
        int index,
        List<object> attempts,
        out string? strategy)
    {
        strategy = null;
        if (index < 0)
            throw new ArgumentOutOfRangeException(nameof(index), "ComboBox index must be >= 0.");

        if (!ExpandComboBox(combo, attempts))
            _logger.LogDebug("Expand before index selection did not report success.");

        Thread.Sleep(ExpandDelayMs);

        var items = EnumerateItemCandidates(combo, null, null, new UiRequest()).ToList();
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
