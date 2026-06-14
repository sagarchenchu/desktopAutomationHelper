using System;
using System.Collections.Generic;
using DesktopAutomationDriver.Models.Request;

namespace DesktopAutomationDriver.Services.Resolution;

public sealed class ElementResolutionException : Exception
{
    public UiLocator? Locator { get; }
    public UiLocator? ParentLocator { get; }
    public string SearchRoot { get; }
    public string Operation { get; }
    public int CandidatesScanned { get; }
    public IReadOnlyList<ElementCandidate> TopRejectedCandidates { get; }

    public ElementResolutionException(
        string message,
        UiLocator? locator,
        UiLocator? parentLocator,
        string searchRoot,
        string operation,
        int candidatesScanned,
        IReadOnlyList<ElementCandidate> topRejectedCandidates) : base(message)
    {
        Locator = locator;
        ParentLocator = parentLocator;
        SearchRoot = searchRoot;
        Operation = operation;
        CandidatesScanned = candidatesScanned;
        TopRejectedCandidates = topRejectedCandidates;
    }
}
