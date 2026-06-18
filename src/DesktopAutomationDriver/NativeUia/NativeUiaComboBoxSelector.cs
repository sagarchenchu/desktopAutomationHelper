using System.Diagnostics;
using System.Drawing;
using System.Net;
using DesktopAutomationDriver.Models.Request;
using Interop.UIAutomationClient;
using Microsoft.Extensions.Logging;

namespace DesktopAutomationDriver.NativeUia;

/// <summary>
/// pywinauto-style native UIA ComboBox selector for <c>selectcomboboxuia</c>.
/// Does not use FlaUI AutomationElement or FlaUI patterns.
/// </summary>
internal sealed class NativeUiaComboBoxSelector
{
    private const int ExpandDelayMs = 250;
    private const int ActionDelayMs = 120;
    private const int DefaultTimeoutMs = 20000;
    private const int MaxTimeoutMs = 60000;

    private readonly IUIAutomation _automation;
    private readonly NativeUiaFinder _finder;
    private readonly ILogger _logger;

    public NativeUiaComboBoxSelector(ILogger logger)
    {
        _logger = logger;
        try
        {
            _automation = new CUIAutomation8();
        }
        catch
        {
            _automation = new CUIAutomation();
        }

        _finder = new NativeUiaFinder(_automation, logger);
    }

    public object Select(
        UiRequest request,
        IntPtr? rootHwnd,
        int? processId,
        CancellationToken cancellationToken = default)
    {
        var sw = Stopwatch.StartNew();
        var strategiesTried = new List<string> { "resolve-combobox" };
        var timeoutMs = ResolveTimeoutMs(request.TimeoutMs);
        var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);
        var matchMode = request.MatchMode ?? request.Locator?.MatchMode ?? "exact";

        cancellationToken.ThrowIfCancellationRequested();

        var root = BuildRoot(rootHwnd);
        var locator = request.Locator ?? new UiLocator();
        if (string.IsNullOrWhiteSpace(locator.ControlType))
            locator.ControlType = "ComboBox";

        var resolve = _finder.FindOne(
            root,
            locator,
            request.ParentLocator,
            processId,
            timeoutMs,
            includeOffscreen: locator.IncludeOffscreen ?? true);

        if (!resolve.Found || resolve.Element == null)
        {
            return ToFailure(
                request,
                "failed",
                resolve.LastError,
                strategiesTried,
                sw.ElapsedMilliseconds,
                comboCandidates: resolve.Candidates);
        }

        var combo = resolve.Element;
        var comboSnapshot = _finder.Snapshot(combo);
        var requestedValue = request.Index.HasValue && string.IsNullOrWhiteSpace(request.Value)
            ? null
            : WebUtility.HtmlDecode(request.Value ?? string.Empty).Trim();

        try
        {
            if (NativeUiaFinder.SafeBoolProperty(combo, NativeUiaConstants.UIA_IsEnabledPropertyId) == false)
            {
                return ToFailure(
                    request,
                    "combo-disabled",
                    "ComboBox is disabled.",
                    strategiesTried,
                    sw.ElapsedMilliseconds,
                    comboSnapshot);
            }

            // Strategy 5: editable ValuePattern
            if (!string.IsNullOrWhiteSpace(requestedValue)
                && TryValuePatternSetValue(combo, requestedValue, strategiesTried, out var valueStrategy)
                && VerifyComboValue(combo, requestedValue, matchMode, out var valueActual, out var valueReason))
            {
                return ToSuccess(
                    request,
                    requestedValue,
                    valueActual,
                    valueStrategy,
                    true,
                    valueReason,
                    comboSnapshot,
                    null,
                    strategiesTried,
                    sw.ElapsedMilliseconds);
            }

            strategiesTried.Add("expandcollapse-expand");
            var expandedBy = TryExpand(combo);
            Thread.Sleep(ExpandDelayMs);

            IUIAutomationElement? matchedItem = null;
            string? selectionStrategy = null;

            if (request.Index.HasValue && string.IsNullOrWhiteSpace(requestedValue))
            {
                var items = FindComboBoxDropdownItems(combo, processId);
                strategiesTried.Add("find-dropdown-items");
                if (request.Index.Value >= 0 && request.Index.Value < items.Count)
                {
                    matchedItem = items[request.Index.Value].Element;
                    requestedValue = items[request.Index.Value].Snapshot.Name;
                }
            }
            else if (!string.IsNullOrWhiteSpace(requestedValue))
            {
                matchedItem = SelectWithVisibleItems(
                    combo,
                    rootHwnd,
                    processId,
                    requestedValue,
                    matchMode,
                    request,
                    strategiesTried,
                    cancellationToken,
                    deadline,
                    out selectionStrategy);
            }
            else
            {
                return ToFailure(
                    request,
                    "invalid-request",
                    "Either value or index is required.",
                    strategiesTried,
                    sw.ElapsedMilliseconds,
                    comboSnapshot);
            }

            if (matchedItem != null && selectionStrategy == null)
            {
                selectionStrategy = TryActivateItem(matchedItem, strategiesTried);
            }

            if (matchedItem == null
                && request.AllowKeyboardFallback == true
                && !string.IsNullOrWhiteSpace(requestedValue)
                && DateTime.UtcNow < deadline)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (TryKeyboardTypeahead(combo, requestedValue))
                {
                    strategiesTried.Add("keyboard-typeahead-enter");
                    selectionStrategy = "keyboard-typeahead-enter";
                }
            }

            strategiesTried.Add("verify-valuepattern");
            var actual = ReadComboValue(combo);
            var verified = !string.IsNullOrWhiteSpace(requestedValue)
                           && NativeUiaText.TextMatches(actual, requestedValue, matchMode);
            if (request.Index.HasValue && string.IsNullOrWhiteSpace(request.Value))
                verified = matchedItem != null;

            var verificationReason = verified
                ? "ValuePattern or selected text matched requested value"
                : selectionStrategy != null
                    ? "selection action succeeded but final value could not be verified"
                    : "All strategies failed";

            if (selectionStrategy != null)
            {
                return ToSuccess(
                    request,
                    requestedValue ?? request.Index?.ToString() ?? "",
                    actual,
                    selectionStrategy,
                    verified,
                    verificationReason,
                    comboSnapshot,
                    matchedItem == null ? null : _finder.Snapshot(matchedItem),
                    strategiesTried,
                    sw.ElapsedMilliseconds,
                    FindComboBoxDropdownItems(combo, processId).Select(c => c.Snapshot).ToList());
            }

            return ToFailure(
                request,
                "failed",
                verificationReason,
                strategiesTried,
                sw.ElapsedMilliseconds,
                comboSnapshot,
                visibleItems: FindComboBoxDropdownItems(combo, processId).Select(c => c.Snapshot).ToList());
        }
        finally
        {
            TryCollapse(combo);
        }
    }

    public object FindOnly(
        UiRequest request,
        IntPtr? rootHwnd,
        int? processId,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var timeoutMs = ResolveTimeoutMs(request.TimeoutMs);
        var root = BuildRoot(rootHwnd);
        var locator = request.Locator ?? new UiLocator();
        if (string.IsNullOrWhiteSpace(locator.ControlType))
            locator.ControlType = "ComboBox";

        var resolve = _finder.FindOne(root, locator, request.ParentLocator, processId, timeoutMs);
        if (!resolve.Found || resolve.Element == null)
        {
            return new
            {
                operation = "findcomboboxuia",
                found = false,
                strategy = resolve.Strategy,
                error = resolve.LastError,
                candidates = resolve.Candidates
            };
        }

        var combo = resolve.Element;
        var snapshot = _finder.Snapshot(combo);
        TryExpand(combo);
        Thread.Sleep(ExpandDelayMs);

        try
        {
            var items = FindComboBoxDropdownItems(combo, processId).Select(c => c.Snapshot).Take(50).ToList();
            return new
            {
                operation = "findcomboboxuia",
                found = true,
                strategy = resolve.Strategy,
                ambiguous = resolve.Ambiguous,
                comboBox = snapshot,
                itemsPreview = items,
                candidates = resolve.Candidates
            };
        }
        finally
        {
            TryCollapse(combo);
        }
    }

    private IUIAutomationElement? SelectWithVisibleItems(
        IUIAutomationElement combo,
        IntPtr? rootHwnd,
        int? processId,
        string requestedValue,
        string matchMode,
        UiRequest request,
        List<string> strategiesTried,
        CancellationToken cancellationToken,
        DateTime deadline,
        out string? selectionStrategy)
    {
        selectionStrategy = null;
        strategiesTried.Add("find-dropdown-items");

        var maxPages = Math.Clamp(request.MaxAttempts ?? 300, 1, 300);
        var delayMs = Math.Clamp(request.DelayMs ?? 150, 50, 2000);
        var seenBatches = new HashSet<string>();

        for (var page = 0; page < maxPages && DateTime.UtcNow < deadline; page++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var items = FindComboBoxDropdownItems(combo, processId);
            var batchKey = string.Join("|", items.Take(20).Select(i => i.Snapshot.RuntimeId));
            if (!seenBatches.Add(batchKey) && page > 0)
                break;

            var matched = MatchItem(items, requestedValue, null, matchMode);
            if (matched != null)
            {
                selectionStrategy = TryActivateItem(matched.Element, strategiesTried);
                return matched.Element;
            }

            if (page < maxPages - 1)
            {
                strategiesTried.Add("paged-visible-dropdown-search");
                NativeUiaInput.WheelDown(3);
                Thread.Sleep(delayMs);
            }
        }

        return null;
    }

    private List<NativeUiaCandidate> FindComboBoxDropdownItems(
        IUIAutomationElement combo,
        int? processId)
    {
        var itemTypeIds = new[]
        {
            NativeUiaConstants.UIA_ListItemControlTypeId,
            NativeUiaConstants.UIA_TextControlTypeId,
            NativeUiaConstants.UIA_DataItemControlTypeId,
            NativeUiaConstants.UIA_CheckBoxControlTypeId,
            NativeUiaConstants.UIA_CustomControlTypeId,
            NativeUiaConstants.UIA_MenuItemControlTypeId,
            NativeUiaConstants.UIA_RadioButtonControlTypeId,
            NativeUiaConstants.UIA_ButtonControlTypeId
        };

        var results = new List<NativeUiaCandidate>();
        var seen = new HashSet<string>();

        foreach (var typeId in itemTypeIds)
        {
            foreach (var element in FindDescendantsByControlType(combo, typeId, 200))
            {
                AddItemCandidate(element, results, seen, processId, "combo-descendant");
            }
        }

        var comboRect = GetRectangle(combo);
        if (comboRect.HasValue)
        {
            foreach (var container in FindPopupContainers(comboRect.Value, processId))
            {
                foreach (var typeId in itemTypeIds)
                {
                    foreach (var element in FindDescendantsByControlType(container, typeId, 250))
                    {
                        AddItemCandidate(element, results, seen, processId, "desktop-popup");
                    }
                }
            }
        }

        return results.OrderByDescending(c => c.Score).ToList();
    }

    private void AddItemCandidate(
        IUIAutomationElement element,
        List<NativeUiaCandidate> results,
        HashSet<string> seen,
        int? processId,
        string reason)
    {
        var snapshot = _finder.Snapshot(element);
        if (string.IsNullOrWhiteSpace(snapshot.MatchText))
            return;

        if (processId.HasValue && snapshot.ProcessId != processId.Value)
            return;

        if (!seen.Add(snapshot.RuntimeId))
            return;

        var score = 50;
        if (snapshot.IsOffscreen != true) score += 20;
        if (snapshot.IsEnabled == true) score += 10;

        results.Add(new NativeUiaCandidate
        {
            Element = element,
            Snapshot = snapshot,
            Score = score,
            Reason = reason
        });
    }

    private IEnumerable<IUIAutomationElement> FindPopupContainers(Rectangle comboRect, int? processId)
    {
        var desktop = _finder.GetDesktopRoot();
        var containerTypes = new[]
        {
            NativeUiaConstants.UIA_ListControlTypeId,
            NativeUiaConstants.UIA_WindowControlTypeId,
            NativeUiaConstants.UIA_PaneControlTypeId,
            NativeUiaConstants.UIA_CustomControlTypeId
        };

        foreach (var typeId in containerTypes)
        {
            foreach (var container in FindDescendantsByControlType(desktop, typeId, 80))
            {
                var rect = GetRectangle(container);
                if (!rect.HasValue)
                    continue;

                if (processId.HasValue)
                {
                    var pid = NativeUiaFinder.SafeIntProperty(container, NativeUiaConstants.UIA_ProcessIdPropertyId);
                    if (pid != processId.Value)
                        continue;
                }

                if (RectsNearCombo(comboRect, rect.Value))
                    yield return container;
            }
        }
    }

    private static bool RectsNearCombo(Rectangle combo, Rectangle candidate)
    {
        var horizontalOverlap = combo.Left <= candidate.Right && candidate.Left <= combo.Right;
        var verticalNear = Math.Abs(candidate.Top - combo.Bottom) < 80
                           || candidate.IntersectsWith(combo);
        return horizontalOverlap && verticalNear;
    }

    private List<IUIAutomationElement> FindDescendantsByControlType(
        IUIAutomationElement root,
        int controlTypeId,
        int limit)
    {
        var results = new List<IUIAutomationElement>();
        try
        {
            var condition = _automation.CreatePropertyCondition(
                NativeUiaConstants.UIA_ControlTypePropertyId,
                controlTypeId);
            var arr = root.FindAll(TreeScope.TreeScope_Descendants, condition);
            var count = Math.Min(arr.Length, limit);
            for (var i = 0; i < count; i++)
            {
                try { results.Add(arr.GetElement(i)); }
                catch { /* stale */ }
            }
        }
        catch { /* ignore */ }

        return results;
    }

    private NativeUiaCandidate? MatchItem(
        List<NativeUiaCandidate> items,
        string? value,
        int? index,
        string matchMode)
    {
        if (index.HasValue)
            return index.Value >= 0 && index.Value < items.Count ? items[index.Value] : null;

        if (string.IsNullOrWhiteSpace(value))
            return null;

        return items.FirstOrDefault(item =>
            NativeUiaText.TextMatches(item.Snapshot.Name, value, matchMode)
            || NativeUiaText.TextMatches(item.Snapshot.Value, value, matchMode)
            || NativeUiaText.TextMatches(item.Snapshot.Text, value, matchMode)
            || NativeUiaText.TextMatches(item.Snapshot.MatchText, value, matchMode));
    }

    private string? TryActivateItem(IUIAutomationElement item, List<string> strategiesTried)
    {
        try
        {
            if (item.GetCurrentPattern(NativeUiaConstants.UIA_SelectionItemPatternId)
                is IUIAutomationSelectionItemPattern selectionItem)
            {
                selectionItem.Select();
                Thread.Sleep(ActionDelayMs);
                strategiesTried.Add("uia-selectionitem-select");
                return "uia-selectionitem-select";
            }
        }
        catch (Exception ex) { _logger.LogDebug(ex, "SelectionItem failed"); }

        try
        {
            if (item.GetCurrentPattern(NativeUiaConstants.UIA_InvokePatternId) is IUIAutomationInvokePattern invoke)
            {
                invoke.Invoke();
                Thread.Sleep(ActionDelayMs);
                strategiesTried.Add("uia-invoke-item");
                return "uia-invoke-item";
            }
        }
        catch (Exception ex) { _logger.LogDebug(ex, "Invoke failed"); }

        try
        {
            if (item.GetCurrentPattern(NativeUiaConstants.UIA_TogglePatternId) is IUIAutomationTogglePattern toggle)
            {
                if (toggle.CurrentToggleState == ToggleState.ToggleState_Off)
                    toggle.Toggle();
                Thread.Sleep(ActionDelayMs);
                strategiesTried.Add("uia-toggle-item");
                return "uia-toggle-item";
            }
        }
        catch (Exception ex) { _logger.LogDebug(ex, "Toggle failed"); }

        var rect = GetRectangle(item);
        if (rect.HasValue)
        {
            foreach (var (point, name) in ClickPoints(rect.Value))
            {
                if (NativeUiaInput.ClickPoint(point))
                {
                    Thread.Sleep(ActionDelayMs);
                    strategiesTried.Add(name);
                    return name;
                }
            }
        }

        try
        {
            item.SetFocus();
            Thread.Sleep(80);
            NativeUiaInput.SendKey(NativeUiaInput.VirtualKeys.Return);
            Thread.Sleep(ActionDelayMs);
            strategiesTried.Add("focus-enter");
            return "focus-enter";
        }
        catch (Exception ex) { _logger.LogDebug(ex, "focus-enter failed"); }

        return null;
    }

    private static IEnumerable<(Point point, string name)> ClickPoints(Rectangle rect)
    {
        yield return (new Point(rect.Left + 8, rect.Top + rect.Height / 2), "physical-click-leftcenter");
        yield return (new Point(rect.Left + rect.Width / 2, rect.Top + rect.Height / 2), "physical-click-center");
        yield return (new Point(rect.Right - 8, rect.Top + rect.Height / 2), "physical-click-rightcenter");
    }

    private bool TryValuePatternSetValue(
        IUIAutomationElement combo,
        string value,
        List<string> strategiesTried,
        out string strategy)
    {
        strategy = "";
        try
        {
            if (combo.GetCurrentPattern(NativeUiaConstants.UIA_ValuePatternId) is IUIAutomationValuePattern valuePattern
                && valuePattern.CurrentIsReadOnly == 0)
            {
                valuePattern.SetValue(value);
                Thread.Sleep(ActionDelayMs);
                strategy = "uia-valuepattern-setvalue";
                strategiesTried.Add(strategy);
                return true;
            }

            var edit = FindDescendantsByControlType(combo, NativeUiaConstants.UIA_EditControlTypeId, 5)
                .FirstOrDefault();
            if (edit?.GetCurrentPattern(NativeUiaConstants.UIA_ValuePatternId) is IUIAutomationValuePattern editValue
                && editValue.CurrentIsReadOnly == 0)
            {
                editValue.SetValue(value);
                Thread.Sleep(ActionDelayMs);
                strategy = "uia-edit-valuepattern-setvalue";
                strategiesTried.Add(strategy);
                return true;
            }
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "ValuePattern set failed");
        }

        return false;
    }

    private string TryExpand(IUIAutomationElement combo)
    {
        try
        {
            if (combo.GetCurrentPattern(NativeUiaConstants.UIA_ExpandCollapsePatternId)
                is IUIAutomationExpandCollapsePattern expandCollapse)
            {
                expandCollapse.Expand();
                return "expandcollapse-expand";
            }
        }
        catch (Exception ex) { _logger.LogDebug(ex, "ExpandCollapse failed"); }

        var openButton = FindDescendantsByControlType(combo, NativeUiaConstants.UIA_ButtonControlTypeId, 10)
            .FirstOrDefault(b =>
                NativeUiaFinder.SafeStringProperty(b, NativeUiaConstants.UIA_NamePropertyId)
                    .Contains("Open", StringComparison.OrdinalIgnoreCase));

        if (openButton?.GetCurrentPattern(NativeUiaConstants.UIA_InvokePatternId) is IUIAutomationInvokePattern invoke)
        {
            invoke.Invoke();
            return "open-button-invoke";
        }

        if (combo.GetCurrentPattern(NativeUiaConstants.UIA_InvokePatternId) is IUIAutomationInvokePattern comboInvoke)
        {
            comboInvoke.Invoke();
            return "combo-invoke";
        }

        var rect = GetRectangle(combo);
        if (rect.HasValue)
        {
            NativeUiaInput.ClickPoint(new Point(rect.Value.Right - 10, rect.Value.Top + rect.Value.Height / 2));
            return "physical-right-edge-click";
        }

        return "none";
    }

    private void TryCollapse(IUIAutomationElement combo)
    {
        try
        {
            if (combo.GetCurrentPattern(NativeUiaConstants.UIA_ExpandCollapsePatternId)
                is IUIAutomationExpandCollapsePattern expandCollapse)
            {
                expandCollapse.Collapse();
                return;
            }
        }
        catch { /* ignore */ }

        NativeUiaInput.SendKey(NativeUiaInput.VirtualKeys.Escape, release: true);
    }

    private bool TryKeyboardTypeahead(IUIAutomationElement combo, string text)
    {
        try
        {
            combo.SetFocus();
            Thread.Sleep(80);
            NativeUiaInput.SendChord(NativeUiaInput.VirtualKeys.Control, NativeUiaInput.VirtualKeys.A);
            Thread.Sleep(50);
            NativeUiaInput.TypeText(text);
            Thread.Sleep(120);
            NativeUiaInput.SendKey(NativeUiaInput.VirtualKeys.Return);
            Thread.Sleep(ActionDelayMs);
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "keyboard typeahead failed");
            return false;
        }
    }

    private bool VerifyComboValue(
        IUIAutomationElement combo,
        string requestedValue,
        string matchMode,
        out string actualValue,
        out string reason)
    {
        actualValue = ReadComboValue(combo);
        if (NativeUiaText.TextMatches(actualValue, requestedValue, matchMode))
        {
            reason = "ValuePattern matched requested value";
            return true;
        }

        reason = "verification failed";
        return false;
    }

    private string ReadComboValue(IUIAutomationElement combo)
    {
        try
        {
            if (combo.GetCurrentPattern(NativeUiaConstants.UIA_ValuePatternId) is IUIAutomationValuePattern valuePattern)
            {
                var value = valuePattern.CurrentValue;
                if (!string.IsNullOrWhiteSpace(value))
                    return NativeUiaText.Normalize(value);
            }
        }
        catch { /* ignore */ }

        var edit = FindDescendantsByControlType(combo, NativeUiaConstants.UIA_EditControlTypeId, 3).FirstOrDefault();
        if (edit != null)
        {
            try
            {
                if (edit.GetCurrentPattern(NativeUiaConstants.UIA_ValuePatternId) is IUIAutomationValuePattern editValue)
                {
                    var value = editValue.CurrentValue;
                    if (!string.IsNullOrWhiteSpace(value))
                        return NativeUiaText.Normalize(value);
                }
            }
            catch { /* ignore */ }
        }

        var name = NativeUiaFinder.SafeStringProperty(combo, NativeUiaConstants.UIA_NamePropertyId);
        return NativeUiaText.Normalize(name);
    }

    private IUIAutomationElement BuildRoot(IntPtr? rootHwnd)
    {
        if (rootHwnd is > 0)
        {
            var fromHwnd = _finder.ElementFromHwnd(rootHwnd.Value);
            if (fromHwnd != null)
                return fromHwnd;
        }

        return _finder.GetDesktopRoot();
    }

    private static Rectangle? GetRectangle(IUIAutomationElement element)
    {
        try
        {
            var value = element.GetCurrentPropertyValue(NativeUiaConstants.UIA_BoundingRectanglePropertyId);
            if (value is not double[] rect || rect.Length < 4)
                return null;
            return new Rectangle(
                (int)Math.Round(rect[0]),
                (int)Math.Round(rect[1]),
                (int)Math.Round(rect[2]),
                (int)Math.Round(rect[3]));
        }
        catch
        {
            return null;
        }
    }

    private static int ResolveTimeoutMs(int? requestTimeoutMs)
    {
        var timeout = requestTimeoutMs ?? DefaultTimeoutMs;
        return Math.Clamp(timeout, 500, MaxTimeoutMs);
    }

    private static object ToSuccess(
        UiRequest request,
        string requested,
        string actual,
        string strategy,
        bool verified,
        string verificationReason,
        NativeUiaElementSnapshot combo,
        NativeUiaElementSnapshot? selectedItem,
        List<string> strategiesTried,
        long elapsedMs,
        List<NativeUiaElementSnapshot>? candidateItems = null) =>
        new
        {
            operation = "selectcomboboxuia",
            success = true,
            requestedValue = requested,
            requestedIndex = request.Index,
            actualValue = actual,
            strategy,
            verified,
            verificationReason,
            comboBox = combo,
            selectedItem,
            strategiesTried,
            candidateItems = candidateItems ?? new List<NativeUiaElementSnapshot>(),
            elapsedMs
        };

    private static object ToFailure(
        UiRequest request,
        string strategy,
        string error,
        List<string> strategiesTried,
        long elapsedMs,
        NativeUiaElementSnapshot? combo = null,
        List<NativeUiaElementSnapshot>? comboCandidates = null,
        List<NativeUiaElementSnapshot>? visibleItems = null) =>
        new
        {
            operation = "selectcomboboxuia",
            success = false,
            requestedValue = request.Value ?? "",
            requestedIndex = request.Index,
            actualValue = "",
            strategy,
            verified = false,
            verificationReason = error,
            comboBox = combo,
            strategiesTried,
            candidateItems = visibleItems ?? new List<NativeUiaElementSnapshot>(),
            elapsedMs,
            diagnostics = new
            {
                comboCandidates = comboCandidates ?? new List<NativeUiaElementSnapshot>(),
                visibleItems = visibleItems ?? new List<NativeUiaElementSnapshot>(),
                lastError = error
            }
        };
}
