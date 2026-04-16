using System.Drawing;
using System.Drawing.Imaging;
using System.Text.RegularExpressions;
using DesktopAutomationDriver.Models.Request;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;

namespace DesktopAutomationDriver.Services;

/// <summary>
/// FlaUI-backed implementation of <see cref="IAutomationService"/>.
/// Runs on Windows and requires UI Automation (UIA2 or UIA3).
/// </summary>
public class AutomationService : IAutomationService
{
    private readonly ILogger<AutomationService> _logger;

    // WebDriver special key codes → Windows virtual keys
    private static readonly Dictionary<string, VirtualKeyShort> SpecialKeys = new()
    {
        ["\uE001"] = VirtualKeyShort.CANCEL,
        ["\uE002"] = VirtualKeyShort.HELP,
        ["\uE003"] = VirtualKeyShort.BACK,
        ["\uE004"] = VirtualKeyShort.TAB,
        ["\uE005"] = VirtualKeyShort.CLEAR,
        ["\uE006"] = VirtualKeyShort.RETURN,
        ["\uE007"] = VirtualKeyShort.RETURN,
        ["\uE008"] = VirtualKeyShort.LSHIFT,
        ["\uE009"] = VirtualKeyShort.LCONTROL,
        ["\uE00A"] = VirtualKeyShort.ALT,
        ["\uE00B"] = VirtualKeyShort.PAUSE,
        ["\uE00C"] = VirtualKeyShort.ESCAPE,
        ["\uE00D"] = VirtualKeyShort.SPACE,
        ["\uE00E"] = VirtualKeyShort.PRIOR,
        ["\uE00F"] = VirtualKeyShort.NEXT,
        ["\uE010"] = VirtualKeyShort.END,
        ["\uE011"] = VirtualKeyShort.HOME,
        ["\uE012"] = VirtualKeyShort.LEFT,
        ["\uE013"] = VirtualKeyShort.UP,
        ["\uE014"] = VirtualKeyShort.RIGHT,
        ["\uE015"] = VirtualKeyShort.DOWN,
        ["\uE016"] = VirtualKeyShort.INSERT,
        ["\uE017"] = VirtualKeyShort.DELETE,
        ["\uE018"] = VirtualKeyShort.OEM_1,
        ["\uE019"] = VirtualKeyShort.OEM_PLUS,
        ["\uE031"] = VirtualKeyShort.F1,
        ["\uE032"] = VirtualKeyShort.F2,
        ["\uE033"] = VirtualKeyShort.F3,
        ["\uE034"] = VirtualKeyShort.F4,
        ["\uE035"] = VirtualKeyShort.F5,
        ["\uE036"] = VirtualKeyShort.F6,
        ["\uE037"] = VirtualKeyShort.F7,
        ["\uE038"] = VirtualKeyShort.F8,
        ["\uE039"] = VirtualKeyShort.F9,
        ["\uE03A"] = VirtualKeyShort.F10,
        ["\uE03B"] = VirtualKeyShort.F11,
        ["\uE03C"] = VirtualKeyShort.F12,
    };

    public AutomationService(ILogger<AutomationService> logger)
    {
        _logger = logger;
    }

    /// <inheritdoc/>
    public string FindElement(AutomationSession session, FindElementRequest request, string? parentElementId = null)
    {
        var root = GetSearchRoot(session, parentElementId);
        var condition = BuildCondition(session, request);
        var element = root.FindFirstDescendant(condition)
            ?? throw new InvalidOperationException(
                $"Element not found using strategy '{request.Using}' with value '{request.Value}'.");
        return session.CacheElement(element);
    }

    /// <inheritdoc/>
    public IList<string> FindElements(AutomationSession session, FindElementRequest request, string? parentElementId = null)
    {
        var root = GetSearchRoot(session, parentElementId);
        var condition = BuildCondition(session, request);
        var elements = root.FindAllDescendants(condition);
        return elements.Select(session.CacheElement).ToList();
    }

    /// <inheritdoc/>
    public void Click(AutomationSession session, string elementId)
    {
        var element = GetElement(session, elementId);
        _logger.LogDebug("Clicking element {ElementId}", elementId);
        element.Click();
    }

    /// <inheritdoc/>
    public void DoubleClick(AutomationSession session, string elementId)
    {
        var element = GetElement(session, elementId);
        _logger.LogDebug("Double-clicking element {ElementId}", elementId);
        element.DoubleClick();
    }

    /// <inheritdoc/>
    public void RightClick(AutomationSession session, string elementId)
    {
        var element = GetElement(session, elementId);
        _logger.LogDebug("Right-clicking element {ElementId}", elementId);
        element.RightClick();
    }

    /// <inheritdoc/>
    public void SendKeys(AutomationSession session, string elementId, string[] keys)
    {
        var element = GetElement(session, elementId);
        element.Focus();

        foreach (var key in keys)
        {
            if (SpecialKeys.TryGetValue(key, out var vk))
            {
                Keyboard.Press(vk);
                Keyboard.Release(vk);
            }
            else
            {
                Keyboard.Type(key);
            }
        }
    }

    /// <inheritdoc/>
    public void Clear(AutomationSession session, string elementId)
    {
        var element = GetElement(session, elementId);
        var textBox = element.AsTextBox();
        textBox.Text = string.Empty;
    }

    /// <inheritdoc/>
    public string GetText(AutomationSession session, string elementId)
    {
        var element = GetElement(session, elementId);

        // Try TextBox value first, fall back to Name property.
        // Limit text retrieval to 1 MB to prevent excessive memory use.
        const int MaxTextLength = 1_048_576;
        try
        {
            if (element.Patterns.Text.IsSupported)
                return element.Patterns.Text.Pattern.DocumentRange.GetText(MaxTextLength);
        }
        catch { /* fall through */ }

        try
        {
            if (element.Patterns.Value.IsSupported)
                return element.Patterns.Value.Pattern.Value ?? string.Empty;
        }
        catch { /* fall through */ }

        return element.Name ?? string.Empty;
    }

    /// <inheritdoc/>
    public string? GetAttribute(AutomationSession session, string elementId, string attributeName)
    {
        var element = GetElement(session, elementId);

        return attributeName.ToLowerInvariant() switch
        {
            "name" => element.Name,
            "automationid" or "automation id" => element.AutomationId,
            "classname" or "class name" => element.ClassName,
            "isenabled" or "enabled" => element.IsEnabled.ToString().ToLowerInvariant(),
            "isoffscreen" or "displayed" => (!element.IsOffscreen).ToString().ToLowerInvariant(),
            "controltype" or "tag name" => element.ControlType.ToString(),
            "value" or "value.value" => TryGetValue(element),
            "helptext" => element.HelpText,
            "processident" or "processid" => element.Properties.ProcessId.ValueOrDefault.ToString(),
            "boundingrectangle" or "rect" => element.BoundingRectangle.ToString(),
            "acceleratorkey" => element.Properties.AcceleratorKey.ValueOrDefault,
            "accesskey" => element.Properties.AccessKey.ValueOrDefault,
            _ => null
        };
    }

    /// <inheritdoc/>
    public bool IsEnabled(AutomationSession session, string elementId)
    {
        var element = GetElement(session, elementId);
        return element.IsEnabled;
    }

    /// <inheritdoc/>
    public bool IsDisplayed(AutomationSession session, string elementId)
    {
        var element = GetElement(session, elementId);
        return !element.IsOffscreen;
    }

    /// <inheritdoc/>
    public string GetControlType(AutomationSession session, string elementId)
    {
        var element = GetElement(session, elementId);
        return element.ControlType.ToString();
    }

    /// <inheritdoc/>
    public string TakeScreenshot(AutomationSession session)
    {
        var mainWindow = session.Application.GetMainWindow(session.Automation)
            ?? throw new InvalidOperationException("Could not find the main window of the application.");
        var rect = mainWindow.BoundingRectangle;

        using var bitmap = new Bitmap(rect.Width, rect.Height, PixelFormat.Format32bppArgb);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.CopyFromScreen(rect.Left, rect.Top, 0, 0, new Size(rect.Width, rect.Height));

        using var stream = new MemoryStream();
        bitmap.Save(stream, ImageFormat.Png);
        return Convert.ToBase64String(stream.ToArray());
    }

    // ------------------------------------------------------------------ helpers

    private AutomationElement GetSearchRoot(AutomationSession session, string? parentElementId)
    {
        if (!string.IsNullOrEmpty(parentElementId))
        {
            return session.GetCachedElement(parentElementId)
                ?? throw new InvalidOperationException(
                    $"Parent element '{parentElementId}' not found in session cache.");
        }

        return session.Application.GetMainWindow(session.Automation)
            ?? throw new InvalidOperationException("Could not find the main window of the application.");
    }

    private AutomationElement GetElement(AutomationSession session, string elementId) =>
        session.GetCachedElement(elementId)
            ?? throw new InvalidOperationException(
                $"Element '{elementId}' not found in session cache. Find the element first.");

    private ConditionBase BuildCondition(AutomationSession session, FindElementRequest request)
    {
        var cf = session.Automation.ConditionFactory;

        return request.Using.ToLowerInvariant() switch
        {
            "automation id" or "id" =>
                cf.ByAutomationId(request.Value),

            "name" or "link text" =>
                cf.ByName(request.Value),

            "partial link text" =>
                new PropertyCondition(
                    session.Automation.PropertyLibrary.Element.Name,
                    request.Value,
                    PropertyConditionFlags.MatchSubstring),

            "class name" =>
                cf.ByClassName(request.Value),

            "tag name" =>
                cf.ByControlType(ParseControlType(request.Value)),

            "xpath" =>
                BuildXPathCondition(session, request.Value),

            _ => throw new ArgumentException(
                $"Unsupported element location strategy: '{request.Using}'. " +
                "Supported: automation id, id, name, link text, partial link text, class name, tag name, xpath.")
        };
    }

    private static ControlType ParseControlType(string value)
    {
        if (Enum.TryParse<ControlType>(value, ignoreCase: true, out var ct))
            return ct;
        throw new ArgumentException($"Unknown control type: '{value}'. " +
            "Valid values: Button, CheckBox, ComboBox, DataGrid, DataItem, Document, Edit, " +
            "Group, Header, HeaderItem, Hyperlink, Image, List, ListItem, Menu, MenuBar, " +
            "MenuItem, Pane, ProgressBar, RadioButton, ScrollBar, Separator, Slider, Spinner, " +
            "SplitButton, StatusBar, Tab, TabItem, Table, Text, Thumb, TitleBar, ToolBar, " +
            "ToolTip, Tree, TreeItem, Window.");
    }

    /// <summary>
    /// Parses a simplified XPath-like expression.
    /// Supports patterns like:
    ///   //Button[@AutomationId='okButton']
    ///   //Edit[@Name='Search']
    ///   //Button[1]
    /// </summary>
    private ConditionBase BuildXPathCondition(AutomationSession session, string xpath)
    {
        var cf = session.Automation.ConditionFactory;

        // Match //ControlType[@Attr='Value']
        var attrMatch = Regex.Match(xpath,
            @"//(\w+)\[@(\w+)='([^']+)'\]", RegexOptions.IgnoreCase);
        if (attrMatch.Success)
        {
            var controlType = attrMatch.Groups[1].Value;
            var attrName = attrMatch.Groups[2].Value.ToLowerInvariant();
            var attrValue = attrMatch.Groups[3].Value;

            var typeCond = cf.ByControlType(ParseControlType(controlType));
            ConditionBase attrCond = attrName switch
            {
                "automationid" => cf.ByAutomationId(attrValue),
                "name"         => cf.ByName(attrValue),
                "classname"    => cf.ByClassName(attrValue),
                _ => throw new ArgumentException(
                    $"Unsupported XPath attribute '@{attrMatch.Groups[2].Value}'. " +
                    "Supported: @AutomationId, @Name, @ClassName.")
            };

            return new AndCondition(typeCond, attrCond);
        }

        // Match //ControlType  (any element of that type)
        var typeOnlyMatch = Regex.Match(xpath, @"//(\w+)$", RegexOptions.IgnoreCase);
        if (typeOnlyMatch.Success)
            return cf.ByControlType(ParseControlType(typeOnlyMatch.Groups[1].Value));

        throw new ArgumentException(
            $"XPath expression '{xpath}' is not supported. " +
            "Supported patterns: //ControlType, //ControlType[@Attr='Value'].");
    }

    private static string? TryGetValue(AutomationElement element)
    {
        try
        {
            if (element.Patterns.Value.IsSupported)
                return element.Patterns.Value.Pattern.Value;
        }
        catch { /* best effort */ }
        return null;
    }
}
