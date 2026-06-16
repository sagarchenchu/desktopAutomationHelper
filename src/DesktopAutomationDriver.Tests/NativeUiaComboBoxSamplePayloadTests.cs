namespace DesktopAutomationDriver.Tests;

/// <summary>
/// Documents native UIA ComboBox request payloads covered by the first migration phase.
/// </summary>
public class NativeUiaComboBoxSamplePayloadTests
{
    public static IEnumerable<object[]> SamplePayloads =>
    [
        [
            "normal-small-combobox",
            """
            {
              "operation": "select",
              "locator": {
                "automationId": "cmbinbound",
                "controlType": "ComboBox"
              },
              "value": "Inbound",
              "timeoutMs": 10000
            }
            """
        ],
        [
            "automationid-only",
            """
            {
              "operation": "select",
              "locator": {
                "automationId": "cmbinbound"
              },
              "value": "Inbound",
              "timeoutMs": 10000
            }
            """
        ],
        [
            "huge-combobox-keyboard-fallback",
            """
            {
              "operation": "select",
              "locator": {
                "automationId": "cmbLarge",
                "controlType": "ComboBox"
              },
              "value": "Some Very Far Item",
              "allowKeyboardFallback": true,
              "timeoutMs": 30000
            }
            """
        ],
        [
            "selectcomboboxitem-alias",
            """
            {
              "operation": "selectcomboboxitem",
              "locator": {
                "automationId": "cmbinbound",
                "controlType": "ComboBox"
              },
              "value": "Inbound",
              "timeoutMs": 20000
            }
            """
        ]
    ];

    [Theory]
    [MemberData(nameof(SamplePayloads))]
    public void SamplePayloads_AreDocumented(string scenario, string payload)
    {
        Assert.False(string.IsNullOrWhiteSpace(scenario));
        Assert.Contains("operation", payload);
        Assert.Contains("locator", payload);
    }
}
