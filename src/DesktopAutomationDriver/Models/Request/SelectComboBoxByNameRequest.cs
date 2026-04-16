namespace DesktopAutomationDriver.Models.Request;

/// <summary>Request body for POST /select/combobox/name.</summary>
public class SelectComboBoxByNameRequest
{
    /// <summary>UIA Name of the ComboBox element.</summary>
    public string? Combobox { get; set; }

    /// <summary>Name of the item to select within the ComboBox.</summary>
    public string? ItemName { get; set; }
}
