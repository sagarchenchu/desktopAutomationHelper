using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Services.ElementResolution;
using DesktopAutomationDriver.Services.Resolution;

namespace DesktopAutomationDriver.Tests;

public class PywinautoElementIdentificationTests
{
    [Fact]
    public void UiLocator_SupportsPywinautoStyleFields()
    {
        var locator = new UiLocator
        {
            AutomationId = "cmbinbound",
            ControlType = "ComboBox",
            Handle = 12345,
            Index = 2,
            MatchMode = "contains",
            SearchScope = "descendants",
            IncludeOffscreen = true,
            IncludeDisabled = false,
            NameRegex = "^Inbound.*",
            BestMatch = "Inbound"
        };

        Assert.Equal(12345L, locator.Hwnd);
        Assert.Equal(2, locator.FoundIndex);
        Assert.Equal("contains", locator.MatchMode);
        Assert.True(locator.IncludeOffscreen);
    }

    [Fact]
    public void ElementSearchCriteria_FromLocator_MapsCoreFields()
    {
        var locator = new UiLocator
        {
            Name = "Patient",
            AutomationId = "txtPatient",
            ControlType = "Edit",
            MatchMode = "contains",
            FoundIndex = 1,
            SearchScope = "children"
        };

        var criteria = ElementSearchCriteria.FromLocator(locator);

        Assert.Equal("Patient", criteria.Name);
        Assert.Equal("txtPatient", criteria.AutomationId);
        Assert.Equal("Edit", criteria.ControlType);
        Assert.Equal("contains", criteria.MatchMode);
        Assert.Equal(1, criteria.FoundIndex);
        Assert.Equal("children", criteria.SearchScope);
    }

    [Theory]
    [InlineData("query", true, true)]
    [InlineData("click", false, false)]
    [InlineData("read", true, true)]
    public void ResolveOptions_DefaultPurposeBehavior(string purpose, bool allowOffscreen, bool allowDisabled)
    {
        var options = new ResolveOptions
        {
            Purpose = purpose,
            AllowOffscreen = allowOffscreen,
            AllowDisabled = allowDisabled,
            ReturnCandidates = true
        };

        Assert.Equal(purpose, options.Purpose);
        Assert.Equal(allowOffscreen, options.AllowOffscreen);
        Assert.Equal(allowDisabled, options.AllowDisabled);
        Assert.True(options.ReturnCandidates);
    }

    [Fact]
    public void FindLocatorPayload_IsDocumented()
    {
        const string payload = """
            {
              "operation": "findlocator",
              "locator": {
                "automationId": "cmbinbound",
                "controlType": "ComboBox"
              },
              "timeoutMs": 5000
            }
            """;

        Assert.Contains("findlocator", payload);
        Assert.Contains("cmbinbound", payload);
    }

    [Fact]
    public void ParentScopedFindLocatorPayload_IsDocumented()
    {
        const string payload = """
            {
              "operation": "findlocator",
              "parentLocator": {
                "automationId": "DataArea",
                "controlType": "Pane"
              },
              "locator": {
                "name": "B19",
                "controlType": "DataItem",
                "searchScope": "descendants"
              }
            }
            """;

        Assert.Contains("parentLocator", payload);
        Assert.Contains("DataArea", payload);
        Assert.Contains("B19", payload);
    }

    [Fact]
    public void ElementSnapshot_MapsPatternFlagsFromResolutionSnapshot()
    {
        var source = new Services.Resolution.ElementSnapshot
        {
            Name = "Save",
            AutomationId = "btnSave",
            ControlType = "Button",
            HasInvoke = true,
            HasValue = false,
            HasSelectionItem = false,
            HasToggle = false,
            HasExpandCollapse = false,
            HasScrollItem = false
        };

        var snapshot = Services.ElementResolution.ElementSnapshot.FromResolutionSnapshot(source);

        Assert.Equal("Save", snapshot.Name);
        Assert.True(snapshot.HasInvokePattern);
        Assert.False(snapshot.HasValuePattern);
    }
}
