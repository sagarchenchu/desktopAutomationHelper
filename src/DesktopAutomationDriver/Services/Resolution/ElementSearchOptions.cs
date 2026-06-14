using System;
using FlaUI.Core.AutomationElements;

namespace DesktopAutomationDriver.Services.Resolution;

public sealed class ElementSearchOptions
{
    public AutomationElement? Parent { get; init; }
    public AutomationElement? Root { get; init; }

    public string SearchRoot { get; init; } = "current";

    public bool TopLevelOnly { get; init; }
    public int? Depth { get; init; }

    public int? FoundIndex { get; init; }
    public int? CtrlIndex { get; init; }

    public bool ActiveOnly { get; init; }
    public bool IncludeOffscreen { get; init; }
    public bool IncludeHidden { get; init; }
    public bool ContentOnly { get; init; }

    public bool PreferAttributes { get; init; } = true;
    public bool PreferXPath { get; init; }
    public bool XPathOnly { get; init; }

    public string? MatchMode { get; init; }
    public string TreeView { get; init; } = "control";

    public bool ThrowIfAmbiguous { get; init; }
    public bool IncludeDiagnostics { get; init; }

    public TimeSpan Timeout { get; init; } = TimeSpan.FromSeconds(5);
    public TimeSpan PollInterval { get; init; } = TimeSpan.FromMilliseconds(250);
}
