using System;
using System.Collections.Generic;
using DesktopAutomationDriver.Models;
using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Services;
using Xunit;

namespace DesktopAutomationDriver.Tests;

public class ElementResolverTests
{
    [Fact]
    public void ResolvedElement_Properties_CanBeSet()
    {
        var resolved = new ResolvedElement
        {
            Strategy = "test-strategy",
            RootStrategy = "test-root",
            Diagnostics = new ElementSearchDiagnostics
            {
                Status = "ElementAmbiguous",
                Message = "Multiple elements found."
            }
        };

        Assert.Equal("test-strategy", resolved.Strategy);
        Assert.Equal("test-root", resolved.RootStrategy);
        Assert.NotNull(resolved.Diagnostics);
        Assert.Equal("ElementAmbiguous", resolved.Diagnostics.Status);
        Assert.Equal("Multiple elements found.", resolved.Diagnostics.Message);
    }

    [Fact]
    public void ElementCandidate_Properties_CanBeSet()
    {
        var candidate = new ElementCandidate
        {
            Index = 3,
            Name = "Save Button",
            AutomationId = "btn_save",
            ClassName = "Button",
            ControlType = "Button",
            FrameworkId = "WPF",
            RuntimeId = "42.123",
            ProcessId = 456,
            Hwnd = 789L,
            Value = "Save",
            Text = "Save",
            IsEnabled = true,
            IsOffscreen = false,
            Score = 100,
            Reason = "Perfect match"
        };

        Assert.Equal(3, candidate.Index);
        Assert.Equal("Save Button", candidate.Name);
        Assert.Equal("btn_save", candidate.AutomationId);
        Assert.Equal("Button", candidate.ClassName);
        Assert.Equal("Button", candidate.ControlType);
        Assert.Equal("WPF", candidate.FrameworkId);
        Assert.Equal("42.123", candidate.RuntimeId);
        Assert.Equal(456, candidate.ProcessId);
        Assert.Equal(789L, candidate.Hwnd);
        Assert.Equal("Save", candidate.Value);
        Assert.Equal("Save", candidate.Text);
        Assert.True(candidate.IsEnabled);
        Assert.False(candidate.IsOffscreen);
        Assert.Equal(100, candidate.Score);
        Assert.Equal("Perfect match", candidate.Reason);
    }

    [Fact]
    public void UiRequest_Criteria_CanBeSet()
    {
        var request = new UiRequest
        {
            Criteria = new List<UiLocator>
            {
                new UiLocator { Name = "Child1" },
                new UiLocator { Name = "Child2" }
            }
        };

        Assert.NotNull(request.Criteria);
        Assert.Equal(2, request.Criteria.Count);
        Assert.Equal("Child1", request.Criteria[0].Name);
        Assert.Equal("Child2", request.Criteria[1].Name);
    }
}
