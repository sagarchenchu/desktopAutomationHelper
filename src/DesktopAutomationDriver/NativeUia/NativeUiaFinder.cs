using System.Drawing;
using DesktopAutomationDriver.Models.Request;
using Interop.UIAutomationClient;
using Microsoft.Extensions.Logging;

namespace DesktopAutomationDriver.NativeUia;

internal sealed class NativeUiaFinder
{
    private readonly IUIAutomation _automation;
    private readonly ILogger _logger;

    public NativeUiaFinder(IUIAutomation automation, ILogger logger)
    {
        _automation = automation;
        _logger = logger;
    }

    public IUIAutomationElement GetDesktopRoot() => _automation.GetRootElement();

    public IUIAutomationElement? ElementFromHwnd(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero)
            return null;
        try
        {
            return _automation.ElementFromHandle(hwnd);
        }
        catch
        {
            return null;
        }
    }

    public NativeUiaFindResult FindOne(
        IUIAutomationElement root,
        UiLocator locator,
        UiLocator? parentLocator,
        int? processId,
        int timeoutMs,
        bool includeOffscreen = true)
    {
        var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);
        NativeUiaFindResult? last = null;

        while (DateTime.UtcNow < deadline)
        {
            var searchRoot = root;
            if (parentLocator != null)
            {
                var parentResult = FindOneInternal(root, parentLocator, processId, includeOffscreen, comboOnly: false);
                if (!parentResult.Found || parentResult.Element == null)
                {
                    last = new NativeUiaFindResult
                    {
                        Found = false,
                        Strategy = "parent-not-found",
                        LastError = parentResult.LastError,
                        Candidates = parentResult.Candidates
                    };
                }
                else
                {
                    searchRoot = parentResult.Element;
                }
            }

            if (searchRoot != null)
            {
                var result = FindOneInternal(searchRoot, locator, processId, includeOffscreen, comboOnly: true);
                if (result.Found)
                    return result;
                last = result;
            }

            Thread.Sleep(150);
        }

        return last ?? new NativeUiaFindResult
        {
            Found = false,
            Strategy = "timeout",
            LastError = "Element not found within timeout."
        };
    }

    public List<NativeUiaCandidate> FindAll(
        IUIAutomationElement root,
        UiLocator locator,
        int? processId,
        int maxDepth = 25,
        bool includeOffscreen = true)
    {
        var raw = CollectElements(root, maxDepth);
        var candidates = new List<NativeUiaCandidate>();

        foreach (var element in raw)
        {
            var snapshot = Snapshot(element);
            if (!MatchesLocator(element, snapshot, locator, processId, includeOffscreen, out var score, out var reason))
                continue;

            candidates.Add(new NativeUiaCandidate
            {
                Element = element,
                Snapshot = snapshot,
                Score = score,
                Reason = reason
            });
        }

        return candidates.OrderByDescending(c => c.Score).ToList();
    }

    public NativeUiaElementSnapshot Snapshot(IUIAutomationElement element)
    {
        var name = SafeStringProperty(element, NativeUiaConstants.UIA_NamePropertyId);
        var automationId = SafeStringProperty(element, NativeUiaConstants.UIA_AutomationIdPropertyId);
        var className = SafeStringProperty(element, NativeUiaConstants.UIA_ClassNamePropertyId);
        var value = TryGetValuePatternValue(element);
        var text = TryGetTextPatternText(element);
        var controlTypeId = SafeIntProperty(element, NativeUiaConstants.UIA_ControlTypePropertyId) ?? 0;

        return new NativeUiaElementSnapshot
        {
            Name = name,
            AutomationId = automationId,
            ClassName = className,
            ControlTypeId = controlTypeId,
            ControlType = NativeUiaText.ControlTypeName(controlTypeId),
            ProcessId = SafeIntProperty(element, NativeUiaConstants.UIA_ProcessIdPropertyId),
            NativeWindowHandle = SafeIntProperty(element, NativeUiaConstants.UIA_NativeWindowHandlePropertyId),
            FrameworkId = SafeStringProperty(element, NativeUiaConstants.UIA_FrameworkIdPropertyId),
            IsEnabled = SafeBoolProperty(element, NativeUiaConstants.UIA_IsEnabledPropertyId),
            IsOffscreen = SafeBoolProperty(element, NativeUiaConstants.UIA_IsOffscreenPropertyId),
            BoundingRectangle = SafeRectangleObject(element),
            Value = value,
            Text = text,
            RuntimeId = SafeRuntimeId(element),
            MatchText = NativeUiaText.Normalize($"{name} {value} {text}")
        };
    }

    private NativeUiaFindResult FindOneInternal(
        IUIAutomationElement root,
        UiLocator locator,
        int? processId,
        bool includeOffscreen,
        bool comboOnly)
    {
        if (locator.Hwnd is > 0 || locator.Handle is > 0)
        {
            var hwnd = locator.Hwnd ?? locator.Handle!.Value;
            var fromHwnd = ElementFromHwnd(new IntPtr(hwnd));
            if (fromHwnd != null)
            {
                var promoted = PromoteToComboBox(fromHwnd) ?? fromHwnd;
                var snap = Snapshot(promoted);
                if (!comboOnly || snap.ControlTypeId == NativeUiaConstants.UIA_ComboBoxControlTypeId)
                {
                    return new NativeUiaFindResult
                    {
                        Found = true,
                        Element = promoted,
                        Strategy = "hwnd-direct",
                        Candidates = [snap]
                    };
                }
            }
        }

        var all = FindAll(root, locator, processId, includeOffscreen: includeOffscreen);
        if (comboOnly)
        {
            all = all.Where(c => c.Snapshot.ControlTypeId == NativeUiaConstants.UIA_ComboBoxControlTypeId).ToList();
            if (all.Count == 0)
                all = FindRelaxedComboCandidates(root, locator, processId, includeOffscreen);
        }

        var snapshots = all.Select(c => c.Snapshot).ToList();
        if (all.Count == 0)
        {
            return new NativeUiaFindResult
            {
                Found = false,
                Strategy = "not-found",
                LastError = "No matching elements found.",
                Candidates = snapshots
            };
        }

        var index = locator.FoundIndex ?? locator.Index;
        if (index.HasValue)
        {
            if (index.Value >= 0 && index.Value < all.Count)
            {
                return new NativeUiaFindResult
                {
                    Found = true,
                    Element = all[index.Value].Element,
                    Strategy = "found-index",
                    Candidates = snapshots
                };
            }

            return new NativeUiaFindResult
            {
                Found = false,
                Strategy = "index-out-of-range",
                LastError = $"foundIndex {index.Value} out of range ({all.Count} candidates).",
                Candidates = snapshots
            };
        }

        if (all.Count > 1)
        {
            var best = all[0];
            return new NativeUiaFindResult
            {
                Found = true,
                Element = best.Element,
                Strategy = "best-candidate",
                Ambiguous = true,
                LastError = $"Multiple matches ({all.Count}); selected highest score.",
                Candidates = snapshots
            };
        }

        return new NativeUiaFindResult
        {
            Found = true,
            Element = all[0].Element,
            Strategy = "unique-match",
            Candidates = snapshots
        };
    }

    private List<NativeUiaCandidate> FindRelaxedComboCandidates(
        IUIAutomationElement root,
        UiLocator locator,
        int? processId,
        bool includeOffscreen)
    {
        var relaxed = new UiLocator
        {
            AutomationId = locator.AutomationId,
            Name = locator.Name,
            ClassName = locator.ClassName,
            ProcessId = locator.ProcessId,
            MatchMode = locator.MatchMode
        };

        var candidates = FindAll(root, relaxed, processId, includeOffscreen: includeOffscreen)
            .Where(c => c.Snapshot.ControlTypeId == NativeUiaConstants.UIA_ComboBoxControlTypeId
                        || c.Snapshot.ControlTypeId == NativeUiaConstants.UIA_EditControlTypeId)
            .ToList();

        foreach (var candidate in candidates.ToList())
        {
            if (candidate.Snapshot.ControlTypeId == NativeUiaConstants.UIA_EditControlTypeId)
            {
                var comboParent = WalkAncestor(
                    candidate.Element,
                    e => SafeIntProperty(e, NativeUiaConstants.UIA_ControlTypePropertyId)
                         == NativeUiaConstants.UIA_ComboBoxControlTypeId);
                if (comboParent != null)
                {
                    candidates.Add(new NativeUiaCandidate
                    {
                        Element = comboParent,
                        Snapshot = Snapshot(comboParent),
                        Score = candidate.Score + 20,
                        Reason = "edit-parent-combobox"
                    });
                }
            }
        }

        return candidates
            .Where(c => c.Snapshot.ControlTypeId == NativeUiaConstants.UIA_ComboBoxControlTypeId)
            .GroupBy(c => c.Snapshot.RuntimeId)
            .Select(g => g.OrderByDescending(x => x.Score).First())
            .OrderByDescending(c => c.Score)
            .ToList();
    }

    private static bool MatchesLocator(
        IUIAutomationElement element,
        NativeUiaElementSnapshot snapshot,
        UiLocator locator,
        int? processId,
        bool includeOffscreen,
        out int score,
        out string reason)
    {
        score = 0;
        reason = "";
        var reasons = new List<string>();

        if (processId.HasValue && snapshot.ProcessId != processId.Value)
            return false;

        if (!string.IsNullOrWhiteSpace(locator.ControlType))
        {
            var expectedId = ParseControlType(locator.ControlType);
            if (expectedId.HasValue && snapshot.ControlTypeId != expectedId.Value)
                return false;
            score += 80;
            reasons.Add("controlType");
        }

        if (!string.IsNullOrWhiteSpace(locator.ClassName))
        {
            if (!NativeUiaText.TextMatches(snapshot.ClassName, locator.ClassName, locator.ClassNameMatchMode ?? locator.MatchMode))
                return false;
            score += 60;
            reasons.Add("className");
        }

        if (!string.IsNullOrWhiteSpace(locator.AutomationId))
        {
            if (!NativeUiaText.TextMatches(snapshot.AutomationId, locator.AutomationId, locator.AutomationIdMatchMode ?? locator.MatchMode))
                return false;
            score += 100;
            reasons.Add("automationId");
        }

        if (!string.IsNullOrWhiteSpace(locator.Name))
        {
            if (!NativeUiaText.TextMatches(snapshot.Name, locator.Name, locator.NameMatchMode ?? locator.MatchMode))
                return false;
            score += 70;
            reasons.Add("name");
        }

        if (!string.IsNullOrWhiteSpace(locator.Value))
        {
            if (!NativeUiaText.TextMatches(snapshot.Value, locator.Value, locator.ValueMatchMode ?? locator.MatchMode)
                && !NativeUiaText.TextMatches(snapshot.MatchText, locator.Value, locator.ValueMatchMode ?? locator.MatchMode))
                return false;
            score += 50;
            reasons.Add("value");
        }

        if (!string.IsNullOrWhiteSpace(locator.FrameworkId)
            && !string.Equals(snapshot.FrameworkId, locator.FrameworkId, StringComparison.OrdinalIgnoreCase))
            return false;

        if (locator.Enabled == true && snapshot.IsEnabled == false)
            return false;

        if (locator.Visible == true && snapshot.IsOffscreen == true)
            return false;

        if (!includeOffscreen && snapshot.IsOffscreen == true)
        {
            score -= 50;
            reasons.Add("offscreen-penalty");
        }
        else if (snapshot.IsOffscreen != true)
        {
            score += 30;
            reasons.Add("visible");
        }

        if (snapshot.IsEnabled == true)
        {
            score += 30;
            reasons.Add("enabled");
        }

        if (!string.IsNullOrWhiteSpace(locator.BestMatch))
        {
            var best = NativeUiaText.Normalize(locator.BestMatch);
            if (NativeUiaText.TextMatches(snapshot.Name, best, "contains"))
                score += 25;
        }

        reason = string.Join(",", reasons);
        return true;
    }

    private List<IUIAutomationElement> CollectElements(IUIAutomationElement root, int maxDepth)
    {
        var results = new List<IUIAutomationElement>();
        CollectRecursive(root, 0, maxDepth, results);
        return results;
    }

    private void CollectRecursive(
        IUIAutomationElement element,
        int depth,
        int maxDepth,
        List<IUIAutomationElement> results)
    {
        if (depth > maxDepth)
            return;

        results.Add(element);

        IUIAutomationElement? child;
        try
        {
            child = _automation.RawViewWalker.GetFirstChildElement(element);
        }
        catch
        {
            return;
        }

        while (child != null)
        {
            CollectRecursive(child, depth + 1, maxDepth, results);
            try
            {
                child = _automation.RawViewWalker.GetNextSiblingElement(child);
            }
            catch
            {
                break;
            }
        }
    }

    private IUIAutomationElement? PromoteToComboBox(IUIAutomationElement element)
    {
        var typeId = SafeIntProperty(element, NativeUiaConstants.UIA_ControlTypePropertyId);
        if (typeId == NativeUiaConstants.UIA_ComboBoxControlTypeId)
            return element;

        if (typeId == NativeUiaConstants.UIA_EditControlTypeId)
            return WalkAncestor(element, e =>
                SafeIntProperty(e, NativeUiaConstants.UIA_ControlTypePropertyId)
                == NativeUiaConstants.UIA_ComboBoxControlTypeId);

        return null;
    }

    private IUIAutomationElement? WalkAncestor(
        IUIAutomationElement element,
        Func<IUIAutomationElement, bool> predicate,
        int maxDepth = 12)
    {
        var current = element;
        for (var i = 0; i < maxDepth; i++)
        {
            if (predicate(current))
                return current;
            try
            {
                current = _automation.RawViewWalker.GetParentElement(current);
                if (current == null)
                    return null;
            }
            catch
            {
                return null;
            }
        }

        return null;
    }

    private static int? ParseControlType(string controlType)
    {
        if (string.IsNullOrWhiteSpace(controlType))
            return null;

        return controlType.Trim().ToLowerInvariant() switch
        {
            "combobox" => NativeUiaConstants.UIA_ComboBoxControlTypeId,
            "edit" => NativeUiaConstants.UIA_EditControlTypeId,
            "listitem" => NativeUiaConstants.UIA_ListItemControlTypeId,
            "list" => NativeUiaConstants.UIA_ListControlTypeId,
            "button" => NativeUiaConstants.UIA_ButtonControlTypeId,
            "checkbox" => NativeUiaConstants.UIA_CheckBoxControlTypeId,
            "text" => NativeUiaConstants.UIA_TextControlTypeId,
            "pane" => NativeUiaConstants.UIA_PaneControlTypeId,
            "window" => NativeUiaConstants.UIA_WindowControlTypeId,
            _ => null
        };
    }

    internal static string SafeStringProperty(IUIAutomationElement element, int propertyId)
    {
        try
        {
            var value = element.GetCurrentPropertyValue(propertyId);
            return value?.ToString()?.Trim() ?? "";
        }
        catch
        {
            return "";
        }
    }

    internal static int? SafeIntProperty(IUIAutomationElement element, int propertyId)
    {
        try
        {
            var value = element.GetCurrentPropertyValue(propertyId);
            return value switch
            {
                int i => i,
                short s => s,
                long l => (int)l,
                _ => int.TryParse(value?.ToString(), out var parsed) ? parsed : null
            };
        }
        catch
        {
            return null;
        }
    }

    internal static bool? SafeBoolProperty(IUIAutomationElement element, int propertyId)
    {
        try
        {
            var value = element.GetCurrentPropertyValue(propertyId);
            return value switch
            {
                bool b => b,
                int i => i != 0,
                _ => bool.TryParse(value?.ToString(), out var parsed) ? parsed : null
            };
        }
        catch
        {
            return null;
        }
    }

    private static string TryGetValuePatternValue(IUIAutomationElement element)
    {
        try
        {
            if (element.GetCurrentPattern(NativeUiaConstants.UIA_ValuePatternId) is IUIAutomationValuePattern valuePattern)
                return valuePattern.CurrentValue ?? "";
        }
        catch { }
        return "";
    }

    private static string TryGetTextPatternText(IUIAutomationElement element)
    {
        try
        {
            if (element.GetCurrentPattern(NativeUiaConstants.UIA_TextPatternId) is IUIAutomationTextPattern textPattern)
                return textPattern.DocumentRange.GetText(-1) ?? "";
        }
        catch { }
        return "";
    }

    private static object? SafeRectangleObject(IUIAutomationElement element)
    {
        try
        {
            var value = element.GetCurrentPropertyValue(NativeUiaConstants.UIA_BoundingRectanglePropertyId);
            if (value is not double[] rect || rect.Length < 4)
                return null;

            return new
            {
                left = (int)Math.Round(rect[0]),
                top = (int)Math.Round(rect[1]),
                width = (int)Math.Round(rect[2]),
                height = (int)Math.Round(rect[3])
            };
        }
        catch
        {
            return null;
        }
    }

    private static string SafeRuntimeId(IUIAutomationElement element)
    {
        try
        {
            var ids = element.GetRuntimeId();
            return ids == null ? "" : string.Join(".", ids);
        }
        catch
        {
            return "";
        }
    }
}
