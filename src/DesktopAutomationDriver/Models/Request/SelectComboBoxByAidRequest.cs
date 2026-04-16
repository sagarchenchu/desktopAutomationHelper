namespace DesktopAutomationDriver.Models.Request;

/// <summary>Request body for POST /select/combobox/aid.</summary>
public class SelectComboBoxByAidRequest
{
    /// <summary>UIA Name of the ComboBox element.</summary>
    public string? Combobox { get; set; }

    /// <summary>UIA AutomationId of the item to select within the ComboBox.</summary>
    public string? AutomationId { get; set; }
}
