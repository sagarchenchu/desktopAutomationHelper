using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Models.Response;

namespace DesktopAutomationDriver.Services;

/// <summary>
/// Provides element-level automation operations for a given session.
/// </summary>
public interface IAutomationService
{
    /// <summary>
    /// Finds a single element in the application window using the given strategy.
    /// </summary>
    /// <param name="session">Active session context.</param>
    /// <param name="request">Find strategy and value.</param>
    /// <param name="parentElementId">Optional parent element ID to scope the search within.</param>
    /// <returns>An element reference ID that can be used in subsequent commands.</returns>
    string FindElement(AutomationSession session, FindElementRequest request, string? parentElementId = null);

    /// <summary>
    /// Finds all elements matching the given strategy.
    /// </summary>
    /// <param name="session">Active session context.</param>
    /// <param name="request">Find strategy and value.</param>
    /// <param name="parentElementId">Optional parent element ID to scope the search within.</param>
    /// <returns>A list of element reference IDs.</returns>
    IList<string> FindElements(AutomationSession session, FindElementRequest request, string? parentElementId = null);

    /// <summary>
    /// Performs a mouse click on the specified element.
    /// </summary>
    void Click(AutomationSession session, string elementId);

    /// <summary>
    /// Performs a mouse double-click on the specified element.
    /// </summary>
    void DoubleClick(AutomationSession session, string elementId);

    /// <summary>
    /// Performs a right-click on the specified element.
    /// </summary>
    void RightClick(AutomationSession session, string elementId);

    /// <summary>
    /// Types the given text into the specified element.
    /// Supports special WebDriver key codes (e.g. "\uE007" for Enter).
    /// </summary>
    void SendKeys(AutomationSession session, string elementId, string[] keys);

    /// <summary>
    /// Clears the value of a text element.
    /// </summary>
    void Clear(AutomationSession session, string elementId);

    /// <summary>
    /// Gets the visible text content of an element.
    /// </summary>
    string GetText(AutomationSession session, string elementId);

    /// <summary>
    /// Gets the named attribute / property of an element.
    /// Common attributes: "Name", "AutomationId", "ClassName", "IsEnabled",
    /// "IsOffscreen", "Value.Value", "Toggle.ToggleState", "BoundingRectangle".
    /// </summary>
    string? GetAttribute(AutomationSession session, string elementId, string attributeName);

    /// <summary>
    /// Returns whether the element is enabled (interactable).
    /// </summary>
    bool IsEnabled(AutomationSession session, string elementId);

    /// <summary>
    /// Returns whether the element is currently visible on screen.
    /// </summary>
    bool IsDisplayed(AutomationSession session, string elementId);

    /// <summary>
    /// Returns the control type name of the element (e.g. "Button", "Edit").
    /// </summary>
    string GetControlType(AutomationSession session, string elementId);

    /// <summary>
    /// Takes a screenshot of the application window and returns it as a Base64-encoded PNG.
    /// </summary>
    string TakeScreenshot(AutomationSession session);
}
