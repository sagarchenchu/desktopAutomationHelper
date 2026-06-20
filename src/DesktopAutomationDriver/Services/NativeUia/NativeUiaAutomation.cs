using System.Drawing;
using System.Runtime.InteropServices;
using Interop.UIAutomationClient;

namespace DesktopAutomationDriver.Services.NativeUia;

/// <summary>
/// Thin wrapper around native UIA COM interfaces for live property and pattern access.
/// </summary>
internal sealed class NativeUiaAutomation
{
    private readonly IUIAutomation _automation;

    public NativeUiaAutomation()
    {
        try
        {
            _automation = new CUIAutomation8();
        }
        catch
        {
            _automation = new CUIAutomation();
        }
    }

    public IUIAutomation Automation => _automation;

    public IUIAutomationElement Root => _automation.GetRootElement();

    public IUIAutomationCondition ControlViewCondition => _automation.ControlViewCondition;

    public IUIAutomationCondition ContentViewCondition => _automation.ContentViewCondition;

    public IUIAutomationElement? FromHandle(IntPtr hwnd)
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

    public IUIAutomationCondition TrueCondition() => _automation.CreateTrueCondition();

    public IUIAutomationCondition PropertyCondition(int propertyId, object value) =>
        _automation.CreatePropertyCondition(propertyId, value);

    public IUIAutomationCondition And(params IUIAutomationCondition[] conditions)
    {
        if (conditions.Length == 0)
            return TrueCondition();
        if (conditions.Length == 1)
            return conditions[0];
        return _automation.CreateAndConditionFromArray(conditions);
    }

    public IUIAutomationCondition Or(params IUIAutomationCondition[] conditions)
    {
        if (conditions.Length == 0)
            return TrueCondition();
        if (conditions.Length == 1)
            return conditions[0];
        return _automation.CreateOrConditionFromArray(conditions);
    }

    public List<IUIAutomationElement> FindAllDescendants(
        IUIAutomationElement root,
        IUIAutomationCondition condition,
        int limit = 5000)
    {
        var results = new List<IUIAutomationElement>();

        try
        {
            var arr = root.FindAll(TreeScope.TreeScope_Descendants, condition);
            var count = Math.Min(arr.Length, limit);

            for (var i = 0; i < count; i++)
            {
                try
                {
                    results.Add(arr.GetElement(i));
                }
                catch
                {
                    // ignore stale element
                }
            }
        }
        catch
        {
            // ignore search failures
        }

        return results;
    }

    public List<IUIAutomationElement> FindAllChildren(
        IUIAutomationElement root,
        IUIAutomationCondition condition,
        int limit = 5000)
    {
        var results = new List<IUIAutomationElement>();

        try
        {
            var arr = root.FindAll(TreeScope.TreeScope_Children, condition);
            var count = Math.Min(arr.Length, limit);

            for (var i = 0; i < count; i++)
            {
                try
                {
                    results.Add(arr.GetElement(i));
                }
                catch
                {
                    // ignore stale element
                }
            }
        }
        catch
        {
            // ignore search failures
        }

        return results;
    }

    public IUIAutomationElement[] GetChildren(IUIAutomationElement root, int maxChildren = 200)
    {
        var results = new List<IUIAutomationElement>(maxChildren);

        try
        {
            var arr = root.FindAll(TreeScope.TreeScope_Children, TrueCondition());
            var count = Math.Min(arr.Length, maxChildren);

            for (var i = 0; i < count; i++)
            {
                try
                {
                    var child = arr.GetElement(i);
                    if (child != null)
                        results.Add(child);
                }
                catch
                {
                    // ignore stale element
                }
            }
        }
        catch
        {
            // ignore search failures
        }

        return results.ToArray();
    }

    public IUIAutomationElement? FindFirstDescendant(
        IUIAutomationElement root,
        IUIAutomationCondition condition)
    {
        try
        {
            return root.FindFirst(TreeScope.TreeScope_Descendants, condition);
        }
        catch
        {
            return null;
        }
    }

    public string GetStringProperty(IUIAutomationElement element, int propertyId)
    {
        try
        {
            var value = element.GetCurrentPropertyValue(propertyId);
            return value switch
            {
                null => string.Empty,
                string s => s,
                _ => value.ToString() ?? string.Empty
            };
        }
        catch
        {
            return string.Empty;
        }
    }

    public int GetIntProperty(IUIAutomationElement element, int propertyId)
    {
        try
        {
            var value = element.GetCurrentPropertyValue(propertyId);
            return value switch
            {
                int i => i,
                short s => s,
                long l => (int)l,
                _ => 0
            };
        }
        catch
        {
            return 0;
        }
    }

    public bool GetBoolProperty(IUIAutomationElement element, int propertyId, bool defaultValue = false)
    {
        try
        {
            var value = element.GetCurrentPropertyValue(propertyId);
            return value switch
            {
                bool b => b,
                int i => i != 0,
                _ => defaultValue
            };
        }
        catch
        {
            return defaultValue;
        }
    }

    public Rectangle? GetBoundingRectangle(IUIAutomationElement element)
    {
        try
        {
            var value = element.GetCurrentPropertyValue(UIA_PropertyIds.UIA_BoundingRectanglePropertyId);
            if (value is not double[] rect || rect.Length < 4)
                return null;

            var left = (int)Math.Round(rect[0]);
            var top = (int)Math.Round(rect[1]);
            var width = (int)Math.Round(rect[2]);
            var height = (int)Math.Round(rect[3]);
            return new Rectangle(left, top, width, height);
        }
        catch
        {
            return null;
        }
    }

    public NativeUiaElementSnapshot CreateSnapshot(IUIAutomationElement element)
    {
        var controlTypeId = GetIntProperty(element, UIA_PropertyIds.UIA_ControlTypePropertyId);
        var patterns = new List<string>();

        foreach (var patternId in new[]
                 {
                     UIA_PatternIds.UIA_ExpandCollapsePatternId,
                     UIA_PatternIds.UIA_SelectionItemPatternId,
                     UIA_PatternIds.UIA_InvokePatternId,
                     UIA_PatternIds.UIA_ValuePatternId,
                     UIA_PatternIds.UIA_LegacyIAccessiblePatternId,
                     UIA_PatternIds.UIA_TextPatternId,
                     UIA_PatternIds.UIA_TogglePatternId,
                     UIA_PatternIds.UIA_ScrollPatternId
                 })
        {
            try
            {
                if (element.GetCurrentPattern(patternId) != null)
                    patterns.Add(PatternName(patternId));
            }
            catch
            {
                // ignore
            }
        }

        return new NativeUiaElementSnapshot
        {
            Name = GetStringProperty(element, UIA_PropertyIds.UIA_NamePropertyId),
            AutomationId = GetStringProperty(element, UIA_PropertyIds.UIA_AutomationIdPropertyId),
            ClassName = GetStringProperty(element, UIA_PropertyIds.UIA_ClassNamePropertyId),
            ControlType = NativeUiaText.ControlTypeName(controlTypeId),
            FrameworkId = GetStringProperty(element, UIA_PropertyIds.UIA_FrameworkIdPropertyId),
            Value = GetValuePatternText(element),
            LegacyName = GetLegacyAccessibleName(element),
            LegacyValue = GetLegacyAccessibleValue(element),
            ProcessId = GetIntProperty(element, UIA_PropertyIds.UIA_ProcessIdPropertyId),
            NativeWindowHandle = GetIntProperty(element, UIA_PropertyIds.UIA_NativeWindowHandlePropertyId),
            IsEnabled = GetBoolProperty(element, UIA_PropertyIds.UIA_IsEnabledPropertyId, true),
            IsOffscreen = GetBoolProperty(element, UIA_PropertyIds.UIA_IsOffscreenPropertyId),
            BoundingRectangle = GetBoundingRectangle(element),
            RuntimeId = GetRuntimeIdString(element),
            SupportedPatterns = patterns
        };
    }

    public string GetElementText(IUIAutomationElement element)
    {
        var name = GetStringProperty(element, UIA_PropertyIds.UIA_NamePropertyId);
        if (!string.IsNullOrWhiteSpace(name))
            return name;

        var value = GetValuePatternText(element);
        if (!string.IsNullOrWhiteSpace(value))
            return value;

        var text = GetTextPatternText(element);
        if (!string.IsNullOrWhiteSpace(text))
            return text;

        var legacyName = GetLegacyAccessibleName(element);
        if (!string.IsNullOrWhiteSpace(legacyName))
            return legacyName;

        return GetLegacyAccessibleValue(element);
    }

    public string GetValuePatternText(IUIAutomationElement element)
    {
        try
        {
            if (element.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId) is not IUIAutomationValuePattern valuePattern)
                return string.Empty;

            return valuePattern.CurrentValue ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    public string GetTextPatternText(IUIAutomationElement element)
    {
        try
        {
            if (element.GetCurrentPattern(UIA_PatternIds.UIA_TextPatternId) is not IUIAutomationTextPattern textPattern)
                return string.Empty;

            return textPattern.DocumentRange.GetText(-1) ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    public string GetLegacyAccessibleName(IUIAutomationElement element)
    {
        try
        {
            if (element.GetCurrentPattern(UIA_PatternIds.UIA_LegacyIAccessiblePatternId)
                is not IUIAutomationLegacyIAccessiblePattern legacy)
                return string.Empty;

            return legacy.CurrentName ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    public string GetLegacyAccessibleValue(IUIAutomationElement element)
    {
        try
        {
            if (element.GetCurrentPattern(UIA_PatternIds.UIA_LegacyIAccessiblePatternId)
                is not IUIAutomationLegacyIAccessiblePattern legacy)
                return string.Empty;

            return legacy.CurrentValue ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    public bool TryGetExpandCollapsePattern(
        IUIAutomationElement element,
        out IUIAutomationExpandCollapsePattern? pattern)
    {
        pattern = null;
        try
        {
            pattern = element.GetCurrentPattern(UIA_PatternIds.UIA_ExpandCollapsePatternId)
                as IUIAutomationExpandCollapsePattern;
            return pattern != null;
        }
        catch
        {
            return false;
        }
    }

    public bool TryGetSelectionItemPattern(
        IUIAutomationElement element,
        out IUIAutomationSelectionItemPattern? pattern)
    {
        pattern = null;
        try
        {
            pattern = element.GetCurrentPattern(UIA_PatternIds.UIA_SelectionItemPatternId)
                as IUIAutomationSelectionItemPattern;
            return pattern != null;
        }
        catch
        {
            return false;
        }
    }

    public bool TryGetInvokePattern(
        IUIAutomationElement element,
        out IUIAutomationInvokePattern? pattern)
    {
        pattern = null;
        try
        {
            pattern = element.GetCurrentPattern(UIA_PatternIds.UIA_InvokePatternId)
                as IUIAutomationInvokePattern;
            return pattern != null;
        }
        catch
        {
            return false;
        }
    }

    public bool TryGetTogglePattern(
        IUIAutomationElement element,
        out IUIAutomationTogglePattern? pattern)
    {
        pattern = null;
        try
        {
            pattern = element.GetCurrentPattern(UIA_PatternIds.UIA_TogglePatternId)
                as IUIAutomationTogglePattern;
            return pattern != null;
        }
        catch
        {
            return false;
        }
    }

    public bool TryGetGridPattern(
        IUIAutomationElement element,
        out IUIAutomationGridPattern? pattern)
    {
        pattern = null;
        try
        {
            pattern = element.GetCurrentPattern(UIA_PatternIds.UIA_GridPatternId)
                as IUIAutomationGridPattern;
            return pattern != null;
        }
        catch
        {
            return false;
        }
    }

    public bool TryGetTablePattern(
        IUIAutomationElement element,
        out IUIAutomationTablePattern? pattern)
    {
        pattern = null;
        try
        {
            pattern = element.GetCurrentPattern(UIA_PatternIds.UIA_TablePatternId)
                as IUIAutomationTablePattern;
            return pattern != null;
        }
        catch
        {
            return false;
        }
    }

    public IUIAutomationCondition ControlTypeCondition(int controlTypeId) =>
        PropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId, controlTypeId);

    public IUIAutomationElement? WalkAncestor(
        IUIAutomationElement element,
        Func<IUIAutomationElement, bool> predicate,
        int maxDepth = 12)
    {
        var current = element;
        for (var depth = 0; depth < maxDepth; depth++)
        {
            try
            {
                if (predicate(current))
                    return current;

                var parent = _automation.RawViewWalker.GetParentElement(current);
                if (parent == null)
                    return null;

                current = parent;
            }
            catch
            {
                return null;
            }
        }

        return null;
    }

    private static string GetRuntimeIdString(IUIAutomationElement element)
    {
        try
        {
            var value = element.GetRuntimeId();
            return value == null ? string.Empty : string.Join(".", value);
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string PatternName(int patternId) => patternId switch
    {
        UIA_PatternIds.UIA_ExpandCollapsePatternId => "ExpandCollapse",
        UIA_PatternIds.UIA_SelectionItemPatternId => "SelectionItem",
        UIA_PatternIds.UIA_InvokePatternId => "Invoke",
        UIA_PatternIds.UIA_ValuePatternId => "Value",
        UIA_PatternIds.UIA_LegacyIAccessiblePatternId => "LegacyIAccessible",
        UIA_PatternIds.UIA_TextPatternId => "Text",
        UIA_PatternIds.UIA_TogglePatternId => "Toggle",
        UIA_PatternIds.UIA_ScrollPatternId => "Scroll",
        _ => $"Pattern({patternId})"
    };
}
