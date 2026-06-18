using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Services;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using Microsoft.Extensions.Logging;

namespace DesktopAutomationDriver.Services.Resolution;

public interface IElementResolver
{
    ElementResolveResult ResolveOne(UiRequest request, string operation);
    ElementResolveResult ResolveOne(UiLocator locator, UiRequest request, string operation);
    IReadOnlyList<ElementCandidate> ResolveAll(UiRequest request, string operation);
    IReadOnlyList<ElementCandidate> ResolveAll(UiLocator locator, UiRequest request, string operation);
    AutomationElement ResolveSearchRoot(UiRequest request);
}

public sealed class ElementResolver : IElementResolver
{
    private readonly IUiSessionContext _ctx;
    private readonly ILogger _logger;
    private readonly Func<AutomationSession, bool, AutomationElement>? _getWindowRoot;

    public AutomationElement ResolveSearchRoot(UiRequest request)
    {
        var session = RequireSession();
        var locator = request.Locator ?? new UiLocator();
        return DetermineSearchRoot(locator, request, session);
    }

    public ElementResolver(IUiSessionContext ctx, ILogger logger)
    {
        _ctx = ctx;
        _logger = logger;
    }

    public ElementResolver(IUiSessionContext ctx, ILogger logger, Func<AutomationSession, bool, AutomationElement> getWindowRoot)
    {
        _ctx = ctx;
        _logger = logger;
        _getWindowRoot = getWindowRoot;
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern int GetDlgCtrlID(IntPtr hwnd);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    public static int? SafeControlId(AutomationElement element)
    {
        try
        {
            var handle = element.Properties.NativeWindowHandle.ValueOrDefault;
            if (handle != IntPtr.Zero)
            {
                return GetDlgCtrlID(handle);
            }
        }
        catch { }
        return null;
    }

    private AutomationSession RequireSession()
    {
        return _ctx.ActiveSession ?? throw new InvalidOperationException("No active automation session.");
    }

    private AutomationElement GetWindowRoot(AutomationSession session, bool allowDesktopPopupScan = true)
    {
        if (_getWindowRoot != null)
        {
            return _getWindowRoot(session, allowDesktopPopupScan);
        }
        return session.ActiveWindow ?? session.Application.GetMainWindow(session.Automation);
    }

    private AutomationElement GetForegroundWindowElement(FlaUI.Core.AutomationBase automation)
    {
        var hwnd = GetForegroundWindow();
        if (hwnd != IntPtr.Zero)
        {
            try
            {
                var el = automation.FromHandle(hwnd);
                if (el != null) return el;
            }
            catch { }
        }
        return automation.GetDesktop();
    }

    private AutomationElement GetActivePopupRoot(AutomationSession session)
    {
        var desktop = session.Automation.GetDesktop();
        try
        {
            var children = desktop.FindAllChildren();
            foreach (var child in children)
            {
                var ct = child.ControlType;
                if ((ct == ControlType.Window || ct == ControlType.Menu || ct == ControlType.List) &&
                    UiService.SafeIsOffscreen(child) == false)
                {
                    var cn = UiService.SafeElementClassName(child);
                    var aid = UiService.SafeElementAutomationId(child);
                    if (cn == "#32768" || cn == "ComboLBox" || cn.Contains("Popup", StringComparison.OrdinalIgnoreCase) || aid.Contains("Popup", StringComparison.OrdinalIgnoreCase))
                    {
                        return child;
                    }
                }
            }
        }
        catch { }

        return GetForegroundWindowElement(session.Automation);
    }

    private AutomationElement DetermineSearchRoot(
        UiLocator locator,
        UiRequest request,
        AutomationSession session)
    {
        // 1. SearchScope priority
        var scope = request.SearchScope ?? locator.SearchScope;
        if (!string.IsNullOrWhiteSpace(scope))
        {
            switch (scope.ToLowerInvariant())
            {
                case "desktop":
                    return session.Automation.GetDesktop();
                case "activewindow":
                case "foreground":
                    return GetForegroundWindowElement(session.Automation);
                case "desktoppopup":
                    return GetActivePopupRoot(session);
                case "currentroot":
                    return GetWindowRoot(session, false);
                case "app":
                    return session.Application.GetMainWindow(session.Automation);
                case "parent":
                    if (request.ParentLocator != null)
                    {
                        var resolvedParent = ResolveOne(request.ParentLocator, request, request.Operation);
                        if (resolvedParent.Element != null) return resolvedParent.Element;
                    }
                    return GetWindowRoot(session, false);
            }
        }

        // 2. ActiveOnly / Active window
        if (locator.ActiveOnly == true || request.ActiveOnly == true)
        {
            return GetForegroundWindowElement(session.Automation);
        }

        // 3. DesktopSearch priority
        if (request.DesktopSearch == true)
        {
            return session.Automation.GetDesktop();
        }

        // 4. Default to window root
        return GetWindowRoot(session, false);
    }

    public ElementResolveResult ResolveOne(UiRequest request, string operation)
    {
        var locator = request.Locator ?? new UiLocator();
        return ResolveOne(locator, request, operation);
    }

    public ElementResolveResult ResolveOne(UiLocator locator, UiRequest request, string operation)
    {
        var session = RequireSession();
        int timeoutMs = request.TimeoutMs ?? 5000;
        int pollIntervalMs = request.PollIntervalMs ?? 250;

        var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);

        while (true)
        {
            var candidates = ExecutePipeline(locator, request, operation, session, out int scannedCount, out var allCandidates);

            if (candidates.Count == 1)
            {
                return new ElementResolveResult
                {
                    Element = candidates[0].Element,
                    Candidates = candidates,
                    Strategy = candidates[0].IsAccepted ? "unified-resolver-unique" : "unified-resolver-failed",
                    Diagnostics = new ElementDiagnostics
                    {
                        CandidatesScanned = scannedCount,
                        Candidates = allCandidates,
                        Message = "Found unique element.",
                        Status = "Success"
                    }
                };
            }

            if (candidates.Count > 1)
            {
                // Ambiguity check
                if (request.ThrowIfAmbiguous == true)
                {
                    throw new ElementResolutionException(
                        "Multiple elements matched. Provide foundIndex, ctrlIndex, parentLocator, automationId, or stronger locator.",
                        locator,
                        request.ParentLocator,
                        request.SearchRoot ?? "",
                        operation,
                        scannedCount,
                        candidates);
                }

                // If not throw, pick best scored candidate
                var bestCandidate = candidates.OrderByDescending(c => c.Score).First();
                return new ElementResolveResult
                {
                    Element = bestCandidate.Element,
                    Candidates = candidates,
                    Strategy = "unified-resolver-scored-best",
                    Diagnostics = new ElementDiagnostics
                    {
                        CandidatesScanned = scannedCount,
                        Candidates = allCandidates,
                        Message = "Multiple elements matched; selected highest score.",
                        Status = "Ambiguous"
                    }
                };
            }

            if (DateTime.UtcNow >= deadline)
            {
                var topRejected = allCandidates.OrderByDescending(c => c.Score).Take(5).ToList();
                throw new ElementResolutionException(
                    $"Element not found for locator={UiService.DescribeLocator(locator)} under operation={operation}.",
                    locator,
                    request.ParentLocator,
                    request.SearchRoot ?? "",
                    operation,
                    scannedCount,
                    topRejected);
            }

            Thread.Sleep(pollIntervalMs);
        }
    }

    public IReadOnlyList<ElementCandidate> ResolveAll(UiRequest request, string operation)
    {
        var locator = request.Locator ?? new UiLocator();
        return ResolveAll(locator, request, operation);
    }

    public IReadOnlyList<ElementCandidate> ResolveAll(UiLocator locator, UiRequest request, string operation)
    {
        var session = RequireSession();
        ExecutePipeline(locator, request, operation, session, out _, out var allCandidates);
        return allCandidates.Where(c => c.IsAccepted).ToList();
    }

    private List<ElementCandidate> ExecutePipeline(
        UiLocator locator,
        UiRequest request,
        string operation,
        AutomationSession session,
        out int scannedCount,
        out List<ElementCandidate> allCandidates)
    {
        allCandidates = new List<ElementCandidate>();
        scannedCount = 0;

        // 1. Resolve parent first if present
        AutomationElement? parentEl = null;
        bool isParentScoped = false;
        if (request.ParentLocator != null)
        {
            try
            {
                var parentRequest = new UiRequest { Locator = request.ParentLocator };
                var resolvedParent = ResolveOne(request.ParentLocator, parentRequest, "resolve-parent");
                parentEl = resolvedParent.Element;
                isParentScoped = parentEl != null;
            }
            catch
            {
                if (request.FallbackToWindowRootIfParentChildNotFound != true)
                {
                    return new List<ElementCandidate>();
                }
            }
        }

        // 2. Build search root
        AutomationElement rootEl = parentEl ?? DetermineSearchRoot(locator, request, session);

        // 3. HWND direct
        if (locator.Hwnd.HasValue || locator.Handle.HasValue)
        {
            var hwndValue = locator.Hwnd ?? locator.Handle!.Value;
            try
            {
                var element = session.Automation.FromHandle(new IntPtr(hwndValue));
                if (element != null)
                {
                    var candidate = CreateCandidate(element);
                    candidate.Score = 100;
                    candidate.MatchReasons.Add("HWND exact match (+100)");
                    allCandidates.Add(candidate);
                    scannedCount = 1;
                    return new List<ElementCandidate> { candidate };
                }
            }
            catch { }
        }

        // 4. Collect raw candidates
        var rawElements = CollectCandidates(rootEl, locator, request, session);
        scannedCount = rawElements.Count;

        // 5. Apply ctrlIndex if present
        var ctrlIndex = locator.CtrlIndex ?? request.CtrlIndex;
        if (ctrlIndex.HasValue)
        {
            var idx = ctrlIndex.Value;
            if (idx >= 0 && idx < rawElements.Count)
            {
                rawElements = new List<AutomationElement> { rawElements[idx] };
            }
            else
            {
                rawElements = new List<AutomationElement>();
            }
        }

        // 6. Snapshot candidates
        foreach (var el in rawElements)
        {
            allCandidates.Add(CreateCandidate(el));
        }

        // 7. Match & Score candidates
        var globalMatchMode = request.MatchMode ?? locator.MatchMode ?? "exact";
        var acceptedCandidates = new List<ElementCandidate>();

        foreach (var candidate in allCandidates)
        {
            ElementMatcher.Match(candidate, locator, globalMatchMode);
            if (candidate.IsAccepted)
            {
                ElementScorer.Score(candidate, locator, new ElementSearchRequest { Locator = locator }, session.Application.ProcessId, isParentScoped);
                acceptedCandidates.Add(candidate);
            }
        }

        // 8. Apply foundIndex
        var foundIndex = locator.FoundIndex ?? request.FoundIndex;
        if (foundIndex.HasValue && acceptedCandidates.Count > 0)
        {
            var idx = foundIndex.Value;
            if (idx >= 0 && idx < acceptedCandidates.Count)
            {
                return new List<ElementCandidate> { acceptedCandidates[idx] };
            }
            return new List<ElementCandidate>();
        }

        return acceptedCandidates;
    }

    internal List<AutomationElement> CollectCandidates(
        AutomationElement root,
        UiLocator locator,
        UiRequest request,
        AutomationSession session)
    {
        var depth = locator.Depth ?? request.Depth ?? 20;
        var topLevelOnly = locator.TopLevelOnly == true || request.TopLevelOnly == true;

        // XPath handling
        if (!string.IsNullOrWhiteSpace(locator.XPath))
        {
            var xpathResult = UiService.FindByXPath(root, session, locator.XPath);
            return xpathResult == null ? new List<AutomationElement>() : new List<AutomationElement> { xpathResult };
        }

        List<AutomationElement> rawCandidates;

        var treeView = request.TreeView?.ToLowerInvariant() ?? "control";
        if (string.Equals(treeView, "raw", StringComparison.OrdinalIgnoreCase))
        {
            var walker = session.Automation.TreeWalkerFactory.GetRawViewWalker();
            if (topLevelOnly)
            {
                var children = new List<AutomationElement>();
                var child = walker.GetFirstChild(root);
                while (child != null)
                {
                    children.Add(child);
                    child = walker.GetNextSibling(child);
                }
                rawCandidates = children;
            }
            else
            {
                rawCandidates = UiService.FindDescendantsWithWalker(root, depth, walker);
            }
        }
        else if (string.Equals(treeView, "content", StringComparison.OrdinalIgnoreCase))
        {
            var walker = session.Automation.TreeWalkerFactory.GetContentViewWalker();
            if (topLevelOnly)
            {
                var children = new List<AutomationElement>();
                var child = walker.GetFirstChild(root);
                while (child != null)
                {
                    children.Add(child);
                    child = walker.GetNextSibling(child);
                }
                rawCandidates = children;
            }
            else
            {
                rawCandidates = UiService.FindDescendantsWithWalker(root, depth, walker);
            }
        }
        else // default to control tree
        {
            var searchScope = locator.SearchScope?.ToLowerInvariant();
            if (string.Equals(searchScope, "children", StringComparison.OrdinalIgnoreCase))
            {
                rawCandidates = root.FindAllChildren().ToList();
            }
            else if (topLevelOnly)
            {
                rawCandidates = root.FindAllChildren().ToList();
            }
            else if (!string.IsNullOrWhiteSpace(locator.ControlType))
            {
                try
                {
                    var normCt = locator.ControlType;
                    var ct = UiService.ParseControlType(normCt);
                    rawCandidates = root.FindAllDescendants(session.Automation.ConditionFactory.ByControlType(ct)).ToList();
                }
                catch
                {
                    rawCandidates = UiService.FindDescendantsUpToDepth(root, depth);
                }
            }
            else
            {
                rawCandidates = UiService.FindDescendantsUpToDepth(root, depth);
            }
        }

        // Active only filter at collection time
        if (locator.ActiveOnly == true || request.ActiveOnly == true)
        {
            var fgHwnd = GetForegroundWindow();
            rawCandidates = rawCandidates.Where(c =>
            {
                try
                {
                    var handle = c.Properties.NativeWindowHandle.ValueOrDefault;
                    return handle != IntPtr.Zero && handle == fgHwnd;
                }
                catch { return false; }
            }).ToList();
        }

        // Include offscreen
        bool includeOffscreen = locator.IncludeOffscreen ?? request.IncludeOffscreen ?? true;
        if (!includeOffscreen)
        {
            rawCandidates = rawCandidates.Where(c => UiService.SafeIsOffscreen(c) != true).ToList();
        }

        return rawCandidates;
    }

    private ElementCandidate CreateCandidate(AutomationElement element)
    {
        return new ElementCandidate
        {
            Element = element,
            Snapshot = CreateSnapshot(element)
        };
    }

    public static ElementSnapshot CreateSnapshot(AutomationElement element)
    {
        var name = UiService.SafeElementName(element);
        var automationId = UiService.SafeElementAutomationId(element);
        var className = UiService.SafeElementClassName(element);
        var controlType = UiService.SafeElementControlType(element);
        var frameworkId = UiService.SafeFrameworkId(element);
        var value = UiService.SafeElementValue(element);
        var text = UiService.SafeElementText(element);
        var legacyName = ElementTextExtractor.GetLegacyName(element);
        var legacyValue = ElementTextExtractor.GetLegacyValue(element);

        int? hwnd = null;
        try
        {
            var h = element.Properties.NativeWindowHandle.ValueOrDefault;
            if (h != IntPtr.Zero) hwnd = (int)h.ToInt64();
        }
        catch { }

        int? pid = null;
        try { pid = UiService.SafeProcessId(element); } catch { }

        int? cid = null;
        try { cid = SafeControlId(element); } catch { }

        bool? isEnabled = null;
        try { isEnabled = UiService.SafeIsEnabled(element); } catch { }

        bool? isOffscreen = null;
        try { isOffscreen = UiService.SafeIsOffscreen(element); } catch { }

        bool? isVisible = null;
        try
        {
            var offscreen = isOffscreen ?? false;
            var rect = element.BoundingRectangle;
            isVisible = !offscreen && !rect.IsEmpty && rect.Width > 0 && rect.Height > 0;
        }
        catch { }

        object? rectObj = null;
        try { rectObj = UiService.SafeBoundingRectangleObject(element); } catch { }

        bool hasInvoke = false;
        try { hasInvoke = element.Patterns.Invoke.IsSupported; } catch { }
        bool hasValue = false;
        try { hasValue = element.Patterns.Value.IsSupported; } catch { }
        bool hasSelectionItem = false;
        try { hasSelectionItem = element.Patterns.SelectionItem.IsSupported; } catch { }
        bool hasToggle = false;
        try { hasToggle = element.Patterns.Toggle.IsSupported; } catch { }
        bool hasScrollItem = false;
        try { hasScrollItem = element.Patterns.ScrollItem.IsSupported; } catch { }
        bool hasExpandCollapse = false;
        try { hasExpandCollapse = element.Patterns.ExpandCollapse.IsSupported; } catch { }

        return new ElementSnapshot
        {
            Name = name,
            AutomationId = automationId,
            ClassName = className,
            ControlType = controlType,
            FrameworkId = frameworkId,
            Value = value,
            Text = text,
            LegacyName = legacyName,
            LegacyValue = legacyValue,
            Hwnd = hwnd,
            ProcessId = pid,
            ControlId = cid,
            IsEnabled = isEnabled,
            IsOffscreen = isOffscreen,
            IsVisible = isVisible,
            Rectangle = rectObj,
            HasInvoke = hasInvoke,
            HasValue = hasValue,
            HasSelectionItem = hasSelectionItem,
            HasToggle = hasToggle,
            HasScrollItem = hasScrollItem,
            HasExpandCollapse = hasExpandCollapse
        };
    }
}
