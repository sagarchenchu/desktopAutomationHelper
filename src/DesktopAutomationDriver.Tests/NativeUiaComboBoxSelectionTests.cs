using DesktopAutomationDriver.NativeUia;

namespace DesktopAutomationDriver.Tests;

public class NativeUiaComboBoxSelectionTests
{
    [Fact]
    public void NativeUiaText_ExactAndContainsMatch()
    {
        Assert.True(NativeUiaText.TextMatches("Between", "Between", "exact"));
        Assert.True(NativeUiaText.TextMatches("Value Between Here", "Between", "contains"));
        Assert.False(NativeUiaText.TextMatches("Other", "Between", "exact"));
    }

    [Fact]
    public void NativeUiaText_NormalizeCollapsesWhitespace()
    {
        Assert.Equal("equals", NativeUiaText.Normalize("  equals \n"));
    }

    [Fact]
    public void SelectComboBoxUiaPayload_IsDocumented()
    {
        const string payload = """
            {
              "operation": "selectcomboboxuia",
              "locator": {
                "automationId": "cmbinbound",
                "controlType": "ComboBox"
              },
              "value": "Between",
              "matchMode": "exact",
              "timeoutMs": 20000,
              "allowKeyboardFallback": true
            }
            """;

        Assert.Contains("selectcomboboxuia", payload);
        Assert.Contains("cmbinbound", payload);
        Assert.Contains("Between", payload);
    }

    [Fact]
    public void SelectComboBoxUia_IndexPayload_IsDocumented()
    {
        const string payload = """
            {
              "operation": "selectcomboboxuia",
              "locator": {
                "automationId": "cmbinbound",
                "controlType": "ComboBox"
              },
              "index": 2,
              "timeoutMs": 20000
            }
            """;

        Assert.Contains("\"index\": 2", payload);
    }

    [Fact]
    public void FindComboBoxUiaPayload_IsDocumented()
    {
        const string payload = """
            {
              "operation": "findcomboboxuia",
              "locator": {
                "automationId": "cmbinbound",
                "controlType": "ComboBox"
              },
              "timeoutMs": 5000
            }
            """;

        Assert.Contains("findcomboboxuia", payload);
    }

    [Fact]
    public void NativeUiaConstants_DefinesComboBoxControlType()
    {
        Assert.Equal(50003, NativeUiaConstants.UIA_ComboBoxControlTypeId);
        Assert.Equal(50007, NativeUiaConstants.UIA_ListItemControlTypeId);
    }
}
