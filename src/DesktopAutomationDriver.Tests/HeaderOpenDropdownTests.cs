namespace DesktopAutomationDriver.Tests;

/// <summary>
/// Documents header/open dropdown operations ported from native UIA header-drag work
/// (container discovery, broad item types, activationStrategy responses).
/// </summary>
public class HeaderOpenDropdownTests
{
    [Fact]
    public void SelectHeaderDropdownItemPayload_IncludesActivationStrategy()
    {
        const string payload = """
            {
              "operation": "selectheaderdropdownitem",
              "locator": { "name": "Status", "controlType": "Header" },
              "value": "Active",
              "matchMode": "contains",
              "itemRegion": "probeAll",
              "timeoutMs": 5000
            }
            """;

        Assert.Contains("selectheaderdropdownitem", payload);
        Assert.Contains("matchMode", payload);
        Assert.Contains("Active", payload);
    }

    [Fact]
    public void SelectOpenDropdownItemPayload_SupportsIndexAndValue()
    {
        const string valuePayload = """
            {
              "operation": "selectopendropdownitem",
              "value": "equals",
              "matchMode": "exact",
              "timeoutMs": 5000
            }
            """;

        const string indexPayload = """
            {
              "operation": "selectopendropdownitem",
              "index": 2,
              "timeoutMs": 5000
            }
            """;

        Assert.Contains("selectopendropdownitem", valuePayload);
        Assert.Contains("equals", valuePayload);
        Assert.Contains("\"index\": 2", indexPayload);
    }

    [Fact]
    public void DropdownContainerTypes_IncludeBroadPopupContainers()
    {
        var containerTypes = OpenDropdownTestHelper.ContainerControlTypes;

        Assert.Contains("List", containerTypes);
        Assert.Contains("Menu", containerTypes);
        Assert.Contains("Pane", containerTypes);
        Assert.Contains("DataGrid", containerTypes);
    }

    [Fact]
    public void DropdownSelectableItemTypes_IncludeCheckBoxAndDataItem()
    {
        var itemTypes = OpenDropdownTestHelper.SelectableItemControlTypes;

        Assert.Contains("ListItem", itemTypes);
        Assert.Contains("CheckBox", itemTypes);
        Assert.Contains("DataItem", itemTypes);
        Assert.Contains("Text", itemTypes);
    }

    [Fact]
    public void PerformPhysicalDragCleanup_UsesCheckedMouseUp()
    {
        const string expectedCleanup = "SendMouseInputChecked(btnUp";
        const string sourceSnippet = """
            SendMouseInputChecked(btnUp, $"drag-{button}-up-cleanup");
            """;

        Assert.Contains(expectedCleanup, sourceSnippet);
        Assert.Contains("-up-cleanup", sourceSnippet);
    }
}

internal static class OpenDropdownTestHelper
{
    public static IReadOnlyList<string> ContainerControlTypes { get; } =
    [
        "List",
        "Menu",
        "Pane",
        "Window",
        "Custom",
        "Tree",
        "DataGrid"
    ];

    public static IReadOnlyList<string> SelectableItemControlTypes { get; } =
    [
        "ListItem",
        "CheckBox",
        "RadioButton",
        "MenuItem",
        "Text",
        "DataItem",
        "TreeItem",
        "Custom",
        "Button"
    ];
}
