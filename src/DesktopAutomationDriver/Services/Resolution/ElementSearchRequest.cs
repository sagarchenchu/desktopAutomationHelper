using System;
using DesktopAutomationDriver.Models.Request;
using FlaUI.Core.AutomationElements;

namespace DesktopAutomationDriver.Services.Resolution;

public sealed class ElementSearchRequest
{
    public UiLocator? Locator { get; init; }
    public UiLocator? ParentLocator { get; init; }

    public AutomationElement? Root { get; init; }

    public bool SearchDesktop { get; init; }
    public bool TopLevelOnly { get; init; }
    public bool IncludeDescendants { get; init; } = true;
    public bool IncludeOffscreen { get; init; }

    public int? Depth { get; init; }
    public int? TimeoutMs { get; init; }
    public int? PollIntervalMs { get; init; }

    public bool RequireSingle { get; init; } = true;
    public bool ReturnCandidates { get; init; } = true;

    public string Operation { get; init; } = "";
}
