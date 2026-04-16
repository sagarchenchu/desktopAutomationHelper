using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Services;
using Microsoft.AspNetCore.Mvc;

namespace DesktopAutomationDriver.Controllers;

/// <summary>
/// Simple, action-oriented endpoints that wrap the underlying <see cref="IUiService"/>
/// with purpose-built request bodies and a consistent <c>{"success": true/false}</c>
/// (or <c>{"found": true/false}</c> for list operations) response shape.
///
/// All routes require: Authorization: Bearer &lt;token&gt;
///
/// POST /launch
/// POST /listallwindows
/// POST /switchwindow
/// POST /maximize
/// POST /click/name
/// POST /click/aid
/// POST /click/advanced
/// POST /doubleclick/name
/// POST /doubleclick/aid
/// POST /select/combobox/name
/// POST /select/combobox/aid
/// </summary>
[ApiController]
public class SimpleAutomationController : ControllerBase
{
    private readonly IUiService _uiService;
    private readonly ILogger<SimpleAutomationController> _logger;

    public SimpleAutomationController(IUiService uiService, ILogger<SimpleAutomationController> logger)
    {
        _uiService = uiService;
        _logger = logger;
    }

    // ── Window / Session management ──────────────────────────────────────────

    /// <summary>
    /// POST /launch
    /// Launches the application at the given path and creates an active session.
    /// Request:  { "exePath": "C:\\path\\to\\app.exe" }
    /// Response: { "success": true } or { "success": false, "error": "..." }
    /// </summary>
    [HttpPost("/launch")]
    public IActionResult Launch([FromBody] LaunchSimpleRequest? req)
    {
        if (string.IsNullOrWhiteSpace(req?.ExePath))
            return BadRequest(new { success = false, error = "'exePath' is required." });

        return RunOperation(new UiRequest { Operation = "launch", Value = req.ExePath });
    }

    /// <summary>
    /// POST /listallwindows
    /// Checks whether any top-level window whose title contains the given name is open.
    /// Request:  { "window": "Notepad" }
    /// Response: { "found": true } or { "found": false }
    /// </summary>
    [HttpPost("/listallwindows")]
    public IActionResult ListAllWindows([FromBody] WindowNameRequest? req)
    {
        if (string.IsNullOrWhiteSpace(req?.Window))
            return BadRequest(new { found = false, error = "'window' is required." });

        try
        {
            var result = _uiService.Execute(
                new UiRequest { Operation = "listwindows", Value = req.Window });

            var list = result as System.Collections.IList;
            bool found = list is { Count: > 0 };
            return Ok(new { found });
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "listallwindows failed for '{Window}'", SanitizeLog(req.Window));
            return Ok(new { found = false });
        }
    }

    /// <summary>
    /// POST /switchwindow
    /// Brings the window whose title contains <c>windowTitle</c> to the foreground
    /// and makes it the active window for subsequent operations.
    /// Request:  { "windowTitle": "Notepad" }
    /// Response: { "success": true } or { "success": false, "error": "..." }
    /// </summary>
    [HttpPost("/switchwindow")]
    public IActionResult SwitchWindow([FromBody] SwitchWindowSimpleRequest? req)
    {
        if (string.IsNullOrWhiteSpace(req?.WindowTitle))
            return BadRequest(new { success = false, error = "'windowTitle' is required." });

        return RunOperation(new UiRequest { Operation = "switchwindow", Value = req.WindowTitle });
    }

    /// <summary>
    /// POST /maximize
    /// Switches to the window whose title contains <c>window</c> and maximizes it.
    /// Request:  { "window": "Notepad" }
    /// Response: { "success": true } or { "success": false, "error": "..." }
    /// </summary>
    [HttpPost("/maximize")]
    public IActionResult Maximize([FromBody] WindowNameRequest? req)
    {
        if (string.IsNullOrWhiteSpace(req?.Window))
            return BadRequest(new { success = false, error = "'window' is required." });

        try
        {
            // Switch to the named window first so that Maximize acts on it.
            _uiService.Execute(new UiRequest { Operation = "switchwindow", Value = req.Window });
            _uiService.Execute(new UiRequest { Operation = "maximize" });
            return Ok(new { success = true });
        }
        catch (ArgumentException ex)
        {
            _logger.LogWarning(ex, "maximize – invalid argument");
            return BadRequest(new { success = false, error = ex.Message });
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "maximize failed for '{Window}'", SanitizeLog(req.Window));
            return Ok(new { success = false, error = ex.Message });
        }
    }

    // ── Click ────────────────────────────────────────────────────────────────

    /// <summary>
    /// POST /click/name
    /// Clicks the first element whose UIA Name matches the given value.
    /// Request:  { "name": "OK" }
    /// Response: { "success": true } or { "success": false, "error": "..." }
    /// </summary>
    [HttpPost("/click/name")]
    public IActionResult ClickByName([FromBody] ClickByNameRequest? req)
    {
        if (string.IsNullOrWhiteSpace(req?.Name))
            return BadRequest(new { success = false, error = "'name' is required." });

        return RunOperation(new UiRequest
        {
            Operation = "click",
            Locator = new UiLocator { Name = req.Name }
        });
    }

    /// <summary>
    /// POST /click/aid
    /// Clicks the first element whose UIA AutomationId matches the given value.
    /// Request:  { "automationId": "btnOK" }
    /// Response: { "success": true } or { "success": false, "error": "..." }
    /// </summary>
    [HttpPost("/click/aid")]
    public IActionResult ClickByAutomationId([FromBody] ClickByAutomationIdRequest? req)
    {
        if (string.IsNullOrWhiteSpace(req?.AutomationId))
            return BadRequest(new { success = false, error = "'automationId' is required." });

        return RunOperation(new UiRequest
        {
            Operation = "click",
            Locator = new UiLocator { AutomationId = req.AutomationId }
        });
    }

    /// <summary>
    /// POST /click/advanced
    /// Clicks the element matching the supplied Name and/or ControlType.
    /// At least one of <c>name</c> or <c>controlType</c> must be provided;
    /// supplying both narrows the match.
    /// Request:  { "name": "Save", "controlType": "Button" }
    /// Response: { "success": true } or { "success": false, "error": "..." }
    /// </summary>
    [HttpPost("/click/advanced")]
    public IActionResult ClickAdvanced([FromBody] ClickAdvancedRequest? req)
    {
        if (string.IsNullOrWhiteSpace(req?.Name) && string.IsNullOrWhiteSpace(req?.ControlType))
            return BadRequest(new { success = false, error = "At least one of 'name' or 'controlType' is required." });

        return RunOperation(new UiRequest
        {
            Operation = "click",
            Locator = new UiLocator { Name = req!.Name, ControlType = req.ControlType }
        });
    }

    // ── Double-click ─────────────────────────────────────────────────────────

    /// <summary>
    /// POST /doubleclick/name
    /// Double-clicks the first element whose UIA Name matches the given value.
    /// Request:  { "name": "MyFile.txt" }
    /// Response: { "success": true } or { "success": false, "error": "..." }
    /// </summary>
    [HttpPost("/doubleclick/name")]
    public IActionResult DoubleClickByName([FromBody] ClickByNameRequest? req)
    {
        if (string.IsNullOrWhiteSpace(req?.Name))
            return BadRequest(new { success = false, error = "'name' is required." });

        return RunOperation(new UiRequest
        {
            Operation = "doubleclick",
            Locator = new UiLocator { Name = req.Name }
        });
    }

    /// <summary>
    /// POST /doubleclick/aid
    /// Double-clicks the first element whose UIA AutomationId matches the given value.
    /// Request:  { "automationId": "listItem1" }
    /// Response: { "success": true } or { "success": false, "error": "..." }
    /// </summary>
    [HttpPost("/doubleclick/aid")]
    public IActionResult DoubleClickByAutomationId([FromBody] ClickByAutomationIdRequest? req)
    {
        if (string.IsNullOrWhiteSpace(req?.AutomationId))
            return BadRequest(new { success = false, error = "'automationId' is required." });

        return RunOperation(new UiRequest
        {
            Operation = "doubleclick",
            Locator = new UiLocator { AutomationId = req.AutomationId }
        });
    }

    // ── ComboBox selection ───────────────────────────────────────────────────

    /// <summary>
    /// POST /select/combobox/name
    /// Finds a ComboBox by its Name and selects an item by the item's visible text.
    /// Request:  { "combobox": "CountryCombo", "itemName": "United States" }
    /// Response: { "success": true } or { "success": false, "error": "..." }
    /// </summary>
    [HttpPost("/select/combobox/name")]
    public IActionResult SelectComboBoxByName([FromBody] SelectComboBoxByNameRequest? req)
    {
        if (string.IsNullOrWhiteSpace(req?.Combobox))
            return BadRequest(new { success = false, error = "'combobox' is required." });
        if (string.IsNullOrWhiteSpace(req.ItemName))
            return BadRequest(new { success = false, error = "'itemName' is required." });

        return RunOperation(new UiRequest
        {
            Operation = "select",
            Locator = new UiLocator { Name = req.Combobox, ControlType = "ComboBox" },
            Value = req.ItemName
        });
    }

    /// <summary>
    /// POST /select/combobox/aid
    /// Finds a ComboBox by its UIA Name property and selects an item by the item's AutomationId.
    /// Request:  { "combobox": "CountryCombo", "automationId": "item_us" }
    /// Response: { "success": true } or { "success": false, "error": "..." }
    /// </summary>
    [HttpPost("/select/combobox/aid")]
    public IActionResult SelectComboBoxByAid([FromBody] SelectComboBoxByAidRequest? req)
    {
        if (string.IsNullOrWhiteSpace(req?.Combobox))
            return BadRequest(new { success = false, error = "'combobox' is required." });
        if (string.IsNullOrWhiteSpace(req.AutomationId))
            return BadRequest(new { success = false, error = "'automationId' is required." });

        return RunOperation(new UiRequest
        {
            Operation = "selectaid",
            Locator = new UiLocator { Name = req.Combobox, ControlType = "ComboBox" },
            Value = req.AutomationId
        });
    }

    // ── Helper ───────────────────────────────────────────────────────────────

    /// <summary>
    /// Executes a <see cref="UiRequest"/> and wraps the result as
    /// <c>{"success": true}</c> or <c>{"success": false, "error": "..."}</c>.
    /// </summary>
    private IActionResult RunOperation(UiRequest req)
    {
        try
        {
            _uiService.Execute(req);
            return Ok(new { success = true });
        }
        catch (ArgumentException ex)
        {
            _logger.LogWarning(ex, "Invalid argument for simple operation '{Op}'", req.Operation);
            return BadRequest(new { success = false, error = ex.Message });
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Simple operation failed: '{Op}'", req.Operation);
            return Ok(new { success = false, error = ex.Message });
        }
    }

    /// <summary>
    /// Removes control characters from a user-supplied string before it is written
    /// to a log entry, preventing log-injection attacks.
    /// </summary>
    private static string SanitizeLog(string? value) =>
        System.Text.RegularExpressions.Regex.Replace(
            value ?? string.Empty, @"[\r\n\t\f]", "_");
}
