using System;
using System.Collections.Generic;
using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Models.Resolver;

namespace DesktopAutomationDriver.Models.Resolver;

public sealed class UiResolutionException : Exception
{
    public string Reason { get; }
    public UiLocator? Locator { get; }
    public IReadOnlyList<ResolvedElement> Candidates { get; }
    public IReadOnlyList<string> Suggestions { get; }

    public UiResolutionException(
        string reason,
        string message,
        UiLocator? locator,
        IReadOnlyList<ResolvedElement> candidates,
        IReadOnlyList<string> suggestions)
        : base(message)
    {
        Reason = reason;
        Locator = locator;
        Candidates = candidates;
        Suggestions = suggestions;
    }
}
