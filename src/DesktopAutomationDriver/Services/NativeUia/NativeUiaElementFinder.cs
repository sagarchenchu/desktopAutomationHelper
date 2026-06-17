using Interop.UIAutomationClient;

namespace DesktopAutomationDriver.Services.NativeUia;

/// <summary>
/// pywinauto-style element finder over native UIA COM.
/// </summary>
internal sealed class NativeUiaElementFinder
{
    private readonly NativeUiaAutomation _uia;

    public NativeUiaElementFinder(NativeUiaAutomation uia)
    {
        _uia = uia;
    }

    public List<IUIAutomationElement> FindElements(
        IUIAutomationElement root,
        NativeUiaLocator locator,
        bool descendants = true,
        int limit = 5000)
    {
        if (locator.Hwnd is > 0)
        {
            var fromHandle = _uia.FromHandle(new IntPtr(locator.Hwnd.Value));
            if (fromHandle != null && ElementMatches(fromHandle, locator))
                return [fromHandle];
        }

        var conditions = BuildBroadConditions(locator);
        var condition = conditions.Count == 0
            ? _uia.TrueCondition()
            : _uia.And(conditions.ToArray());

        var candidates = descendants
            ? _uia.FindAllDescendants(root, condition, limit)
            : _uia.FindAllChildren(root, condition, limit);

        var filtered = candidates
            .Where(e => ElementMatches(e, locator))
            .ToList();

        if (filtered.Count > 0)
            return ApplyFoundIndex(filtered, locator.FoundIndex);

        var relaxed = TryRelaxedFallback(root, locator, descendants, limit);
        return ApplyFoundIndex(relaxed, locator.FoundIndex);
    }

    public IUIAutomationElement? FindFirst(
        IUIAutomationElement root,
        NativeUiaLocator locator,
        bool descendants = true,
        int limit = 5000)
    {
        var matches = FindElements(root, locator, descendants, limit);
        return matches.Count > 0 ? matches[0] : null;
    }

    private List<IUIAutomationElement> TryRelaxedFallback(
        IUIAutomationElement root,
        NativeUiaLocator locator,
        bool descendants,
        int limit)
    {
        var results = new List<IUIAutomationElement>();

        if (!string.IsNullOrWhiteSpace(locator.AutomationId))
        {
            results.AddRange(FindElements(root, locator.AutomationIdOnly(), descendants, limit));
            if (results.Count > 0)
                return results.DistinctBy(RuntimeKey).ToList();
        }

        if (!string.IsNullOrWhiteSpace(locator.Name))
        {
            results.AddRange(FindElements(root, locator.NameOnly(), descendants, limit));
            if (results.Count > 0)
                return results.DistinctBy(RuntimeKey).ToList();
        }

        if (!string.IsNullOrWhiteSpace(locator.ControlType))
        {
            results.AddRange(FindElements(root, locator.ControlTypeOnly(), descendants, limit));
        }

        return results.DistinctBy(RuntimeKey).ToList();
    }

    private List<IUIAutomationCondition> BuildBroadConditions(NativeUiaLocator locator)
    {
        var conditions = new List<IUIAutomationCondition>();

        var controlTypeId = NativeUiaText.ParseControlTypeId(locator.ControlType);
        if (controlTypeId.HasValue)
        {
            conditions.Add(_uia.PropertyCondition(
                UIA_PropertyIds.UIA_ControlTypePropertyId,
                controlTypeId.Value));
        }

        if (!string.IsNullOrWhiteSpace(locator.ClassName))
        {
            conditions.Add(_uia.PropertyCondition(
                UIA_PropertyIds.UIA_ClassNamePropertyId,
                locator.ClassName));
        }

        if (!string.IsNullOrWhiteSpace(locator.AutomationId))
        {
            conditions.Add(_uia.PropertyCondition(
                UIA_PropertyIds.UIA_AutomationIdPropertyId,
                locator.AutomationId));
        }

        if (!string.IsNullOrWhiteSpace(locator.Name))
        {
            conditions.Add(_uia.PropertyCondition(
                UIA_PropertyIds.UIA_NamePropertyId,
                locator.Name));
        }

        if (locator.ProcessId.HasValue)
        {
            conditions.Add(_uia.PropertyCondition(
                UIA_PropertyIds.UIA_ProcessIdPropertyId,
                locator.ProcessId.Value));
        }

        return conditions;
    }

    private bool ElementMatches(IUIAutomationElement element, NativeUiaLocator locator)
    {
        if (locator.ProcessId.HasValue)
        {
            var pid = _uia.GetIntProperty(element, UIA_PropertyIds.UIA_ProcessIdPropertyId);
            if (pid != locator.ProcessId.Value)
                return false;
        }

        var controlTypeId = NativeUiaText.ParseControlTypeId(locator.ControlType);
        if (controlTypeId.HasValue)
        {
            var actualType = _uia.GetIntProperty(element, UIA_PropertyIds.UIA_ControlTypePropertyId);
            if (actualType != controlTypeId.Value)
                return false;
        }

        if (!MatchesString(
                _uia.GetStringProperty(element, UIA_PropertyIds.UIA_ClassNamePropertyId),
                locator.ClassName,
                locator.MatchMode))
            return false;

        if (!MatchesString(
                _uia.GetStringProperty(element, UIA_PropertyIds.UIA_AutomationIdPropertyId),
                locator.AutomationId,
                locator.MatchMode))
            return false;

        if (!MatchesString(
                _uia.GetStringProperty(element, UIA_PropertyIds.UIA_NamePropertyId),
                locator.Name,
                locator.MatchMode))
            return false;

        if (!string.IsNullOrWhiteSpace(locator.Value))
        {
            var value = _uia.GetElementText(element);
            if (!MatchesString(value, locator.Value, locator.MatchMode))
                return false;
        }

        return true;
    }

    private static bool MatchesString(string? actual, string? expected, string matchMode)
    {
        if (string.IsNullOrWhiteSpace(expected))
            return true;

        return NativeUiaText.Matches(actual, expected, matchMode);
    }

    private static List<IUIAutomationElement> ApplyFoundIndex(
        List<IUIAutomationElement> elements,
        int? foundIndex)
    {
        if (!foundIndex.HasValue)
            return elements;

        if (foundIndex.Value < 0 || foundIndex.Value >= elements.Count)
            return [];

        return [elements[foundIndex.Value]];
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
