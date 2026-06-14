using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using DesktopAutomationDriver.Models;
using DesktopAutomationDriver.Models.Request;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using Microsoft.Extensions.Logging;

namespace DesktopAutomationDriver.Services;

public class ElementResolver
{
    private readonly IUiSessionContext _sessionCtx;
    private readonly ILogger<ElementResolver> _logger;
    private readonly Func<AutomationSession, bool, AutomationElement> _getWindowRoot;

    public ElementResolver(
        IUiSessionContext sessionCtx,
        ILogger<ElementResolver> logger,
        Func<AutomationSession, bool, AutomationElement> getWindowRoot)
    {
        _sessionCtx = sessionCtx;
        _logger = logger;
        _getWindowRoot = getWindowRoot;
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern int GetDlgCtrlID(IntPtr hwnd);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    internal static int? SafeControlId(AutomationElement element)
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
        return _sessionCtx.ActiveSession ?? throw new InvalidOperationException("No active automation session.");
    }

    /// <summary>
    /// Requirement 1: Resolves one element from a UiRequest.
    /// </summary>
    public ResolvedElement ResolveOne(UiRequest request)
    {
        var session = RequireSession();
        var locator = request.Locator ?? new UiLocator();

        // Determine Search Root
        var root = DetermineRoot(request, session);
        var rootStrategy = DetermineRootStrategy(request, session);

        // Optional criteria chain (Requirement 5)
        var path = request.Criteria ?? request.LocatorPath;
        List<AutomationElement> candidates;
        if (path != null && path.Count > 0)
        {
            candidates = ResolveLocatorPath(request, path, root);
        }
        else
        {
            var options = BuildOptionsFromRequest(request, locator);
            candidates = ResolveElements(request, locator, root, options);
        }

        var result = new ResolvedElement
        {
            RootStrategy = rootStrategy,
            Strategy = candidates.Count > 0 ? "resolver-unique" : "resolver-not-found"
        };

        if (candidates.Count == 0)
        {
            var targetLoc = locator;
            if (path != null && path.Count > 0) targetLoc = path[^1];

            var rawCandidates = BuildRawCandidatesList(root, targetLoc, session, request);
            var diagnosticsList = BuildCandidateDiagnostics(rawCandidates, targetLoc);

            result.Element = null;
            result.Strategy = "resolver-not-found";
            result.Diagnostics = new ElementSearchDiagnostics
            {
                Status = "ElementNotFound",
                Message = $"ElementNotFound: No matching elements found for locator={UiService.DescribeLocator(targetLoc)}",
                Candidates = diagnosticsList,
                Errors = new List<string> { "No elements matched the filters." }
            };
            return result;
        }

        if (candidates.Count > 1 && !locator.FoundIndex.HasValue)
        {
            var targetLoc = locator;
            if (path != null && path.Count > 0) targetLoc = path[^1];

            var diagnosticsList = BuildCandidateDiagnostics(candidates, targetLoc);

            result.Element = null;
            result.Strategy = "resolver-ambiguous";
            result.Diagnostics = new ElementSearchDiagnostics
            {
                Status = "ElementAmbiguous",
                Message = $"ElementAmbiguous: Multiple matching elements found ({candidates.Count}).",
                Candidates = diagnosticsList,
                Errors = new List<string> { $"Ambiguity detected: found {candidates.Count} matches, but no foundIndex was specified." }
            };
            return result;
        }

        result.Element = candidates[0];
        result.Strategy = candidates.Count > 1 ? "resolver-scored-ambiguous" : "resolver-scored-unique";
        return result;
    }

    /// <summary>
    /// Requirement 2: Resolves all elements from a UiRequest.
    /// </summary>
    public List<ResolvedElement> ResolveAll(UiRequest request)
    {
        var session = RequireSession();
        var locator = request.Locator ?? new UiLocator();

        var root = DetermineRoot(request, session);
        var rootStrategy = DetermineRootStrategy(request, session);

        var path = request.Criteria ?? request.LocatorPath;
        List<AutomationElement> candidates;
        if (path != null && path.Count > 0)
        {
            candidates = ResolveLocatorPath(request, path, root);
        }
        else
        {
            var options = BuildOptionsFromRequest(request, locator);
            candidates = ResolveElements(request, locator, root, options);
        }

        return candidates.Select(c => new ResolvedElement
        {
            Element = c,
            RootStrategy = rootStrategy,
            Strategy = "resolver-all-match"
        }).ToList();
    }

    /// <summary>
    /// Requirement 3: Resolves one element from locator, root, and options.
    /// </summary>
    public ResolvedElement ResolveOne(UiLocator locator, AutomationElement root, ResolveOptions options)
    {
        var session = RequireSession();
        var req = new UiRequest { Locator = locator };
        var candidates = ResolveElements(req, locator, root, options);

        var result = new ResolvedElement
        {
            RootStrategy = "explicit-root",
            Strategy = candidates.Count > 0 ? "resolver-options-unique" : "resolver-options-not-found"
        };

        if (candidates.Count == 0)
        {
            var rawCandidates = BuildRawCandidatesList(root, locator, session, req);
            var diagnosticsList = BuildCandidateDiagnostics(rawCandidates, locator);

            result.Element = null;
            result.Strategy = "resolver-options-not-found";
            result.Diagnostics = new ElementSearchDiagnostics
            {
                Status = "ElementNotFound",
                Message = $"ElementNotFound: No matching elements found for locator={UiService.DescribeLocator(locator)}",
                Candidates = diagnosticsList,
                Errors = new List<string> { "No elements matched the filters." }
            };
            return result;
        }

        if (candidates.Count > 1 && !locator.FoundIndex.HasValue)
        {
            var diagnosticsList = BuildCandidateDiagnostics(candidates, locator);

            result.Element = null;
            result.Strategy = "resolver-options-ambiguous";
            result.Diagnostics = new ElementSearchDiagnostics
            {
                Status = "ElementAmbiguous",
                Message = $"ElementAmbiguous: Multiple matching elements found ({candidates.Count}).",
                Candidates = diagnosticsList,
                Errors = new List<string> { $"Ambiguity detected: found {candidates.Count} matches, but no foundIndex was specified." }
            };
            return result;
        }

        result.Element = candidates[0];
        result.Strategy = candidates.Count > 1 ? "resolver-options-scored-ambiguous" : "resolver-options-scored-unique";
        return result;
    }

    private ResolveOptions BuildOptionsFromRequest(UiRequest request, UiLocator locator)
    {
        return new ResolveOptions
        {
            MatchMode = locator.MatchMode ?? request.MatchMode,
            TreeView = request.TreeView,
            Backend = request.Backend,
            ReturnAllMatches = request.ReturnAllMatches,
            MaxMatches = request.MaxMatches,
            IncludeDiagnostics = request.IncludeDiagnostics,
            AllowBestMatch = request.AllowBestMatch,
            BestMatch = locator.BestMatch ?? request.BestMatch,
            UseDesktopRoot = request.UseDesktopRoot,
            UseActiveWindowRoot = request.UseActiveWindowRoot,
            ReturnCandidates = request.ReturnCandidates,
            Debug = request.Debug,
            Ambiguity = request.Ambiguity
        };
    }

    private AutomationElement DetermineRoot(UiRequest request, AutomationSession session)
    {
        var searchRootName = request.SearchRoot;
        if (!string.IsNullOrWhiteSpace(searchRootName))
        {
            if (string.Equals(searchRootName, "desktop", StringComparison.OrdinalIgnoreCase))
            {
                return session.Automation.GetDesktop();
            }
            if (string.Equals(searchRootName, "foreground", StringComparison.OrdinalIgnoreCase))
            {
                return GetForegroundWindowElement(session.Automation);
            }
            if (string.Equals(searchRootName, "activepopup", StringComparison.OrdinalIgnoreCase))
            {
                return GetActivePopupRoot(session);
            }
            if (string.Equals(searchRootName, "parent", StringComparison.OrdinalIgnoreCase))
            {
                if (request.ParentLocator != null)
                {
                    var windowRoot = _getWindowRoot(session, false);
                    var parentResult = UiService.TryFindElementBySmartStrategy(windowRoot, session, request.ParentLocator, true, false);
                    return parentResult.Element ?? windowRoot;
                }
                return _getWindowRoot(session, false);
            }
        }

        if (request.ParentLocator != null)
        {
            var windowRoot = _getWindowRoot(session, false);
            var parentResult = UiService.TryFindElementBySmartStrategy(windowRoot, session, request.ParentLocator, true, false);
            if (parentResult.Element != null)
            {
                return parentResult.Element;
            }
        }

        if (request.UseDesktopRoot == true)
        {
            return session.Automation.GetDesktop();
        }

        return _getWindowRoot(session, false);
    }

    private string DetermineRootStrategy(UiRequest request, AutomationSession session)
    {
        var searchRootName = request.SearchRoot;
        if (!string.IsNullOrWhiteSpace(searchRootName))
        {
            return searchRootName.ToLowerInvariant();
        }
        if (request.ParentLocator != null)
        {
            return "parent-locator";
        }
        if (request.UseDesktopRoot == true)
        {
            return "desktop";
        }
        if (request.UseActiveWindowRoot == true)
        {
            return "active-window";
        }
        return session.ActiveWindow != null ? "active-window" : "app-main-window";
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

    private List<AutomationElement> ResolveLocatorPath(
        UiRequest request,
        List<UiLocator> path,
        AutomationElement startRoot)
    {
        var currentRoots = new List<AutomationElement> { startRoot };

        for (int i = 0; i < path.Count; i++)
        {
            var locator = path[i];
            var nextRoots = new List<AutomationElement>();

            foreach (var root in currentRoots)
            {
                var options = BuildOptionsFromRequest(request, locator);
                var resolved = ResolveElements(request, locator, root, options);
                nextRoots.AddRange(resolved);
            }

            currentRoots = nextRoots.Distinct().ToList();

            if (currentRoots.Count == 0)
                break;
        }

        return currentRoots;
    }

    private List<AutomationElement> ResolveElements(
        UiRequest request,
        UiLocator locator,
        AutomationElement startRoot,
        ResolveOptions options)
    {
        var session = RequireSession();

        // 1. BuildCandidateList
        var candidates = BuildRawCandidatesList(startRoot, locator, session, request);

        // 2. CtrlIndex raw pre-filter
        if (locator.CtrlIndex.HasValue)
        {
            if (locator.CtrlIndex.Value >= 0 && locator.CtrlIndex.Value < candidates.Count)
            {
                return new List<AutomationElement> { candidates[locator.CtrlIndex.Value] };
            }
            return new List<AutomationElement>();
        }

        // 3. Apply filters (Requirement 6)
        candidates = ApplyFilters(candidates, locator, options);

        // 4. BestMatch scoring & picking highest above threshold
        var bestMatchPattern = locator.BestMatch ?? options.BestMatch;
        if (!string.IsNullOrWhiteSpace(bestMatchPattern) && candidates.Count > 0)
        {
            var bestScored = candidates
                .Select(c => new { Element = c, Score = UiService.CalculateBestMatchScore(UiService.SafeElementName(c), bestMatchPattern) })
                .Where(x => x.Score >= 50)
                .OrderByDescending(x => x.Score)
                .ToList();

            if (bestScored.Count > 0)
            {
                candidates = bestScored.Select(x => x.Element).ToList();
            }
        }

        // 5. FoundIndex post-filter
        if (locator.FoundIndex.HasValue)
        {
            if (locator.FoundIndex.Value >= 0 && locator.FoundIndex.Value < candidates.Count)
            {
                return new List<AutomationElement> { candidates[locator.FoundIndex.Value] };
            }
            return new List<AutomationElement>();
        }

        return candidates;
    }

    private List<AutomationElement> BuildRawCandidatesList(
        AutomationElement root,
        UiLocator locator,
        AutomationSession session,
        UiRequest request)
    {
        var depth = locator.Depth ?? request.Locator?.Depth ?? 20;
        var topLevelOnly = locator.TopLevelOnly == true;

        // HWND direct lookup
        if (locator.Hwnd.HasValue)
        {
            try
            {
                var hwndPtr = new IntPtr(locator.Hwnd.Value);
                var element = session.Automation.FromHandle(hwndPtr);
                return element == null ? new List<AutomationElement>() : new List<AutomationElement> { element };
            }
            catch
            {
                return new List<AutomationElement>();
            }
        }

        // XPath
        if (!string.IsNullOrWhiteSpace(locator.XPath))
        {
            var byXPath = UiService.FindByXPath(root, session, locator.XPath);
            return byXPath == null ? new List<AutomationElement>() : new List<AutomationElement> { byXPath };
        }

        // TreeView walker
        if (!string.IsNullOrWhiteSpace(request.TreeView) && 
            !string.Equals(request.TreeView, "control", StringComparison.OrdinalIgnoreCase))
        {
            var walker = UiService.GetTreeWalker(session.Automation, request.TreeView);
            if (topLevelOnly)
            {
                var children = new List<AutomationElement>();
                try
                {
                    var child = walker.GetFirstChild(root);
                    while (child != null)
                    {
                        children.Add(child);
                        child = walker.GetNextSibling(child);
                    }
                }
                catch { }
                return children;
            }
            return UiService.FindDescendantsWithWalker(root, depth, walker);
        }

        // TopLevelOnly
        if (topLevelOnly)
        {
            return root.FindAllChildren().ToList();
        }

        // ControlType optimization
        if (!string.IsNullOrWhiteSpace(locator.ControlType))
        {
            try
            {
                var ct = UiService.ParseControlType(locator.ControlType);
                return root.FindAllDescendants(
                    session.Automation.ConditionFactory.ByControlType(ct)).ToList();
            }
            catch
            {
                return new List<AutomationElement>();
            }
        }

        // All descendants walk
        return UiService.FindDescendantsUpToDepth(root, depth);
    }

    private List<AutomationElement> ApplyFilters(
        List<AutomationElement> candidates,
        UiLocator locator,
        ResolveOptions options)
    {
        var filtered = new List<AutomationElement>();
        var globalMode = options.MatchMode;

        foreach (var c in candidates)
        {
            // ControlType
            if (!string.IsNullOrWhiteSpace(locator.ControlType))
            {
                var ct = UiService.SafeElementControlType(c);
                if (!string.Equals(ct, locator.ControlType, StringComparison.OrdinalIgnoreCase))
                    continue;
            }

            // ClassName
            if (!string.IsNullOrWhiteSpace(locator.ClassName))
            {
                var cn = UiService.SafeElementClassName(c);
                var mode = locator.ClassNameMatchMode ?? globalMode;
                if (!UiService.StringMatchesByMode(cn, locator.ClassName, mode, caseSensitive: false))
                    continue;
            }

            // ClassNameRegex
            if (!string.IsNullOrWhiteSpace(locator.ClassNameRegex))
            {
                var cn = UiService.SafeElementClassName(c);
                if (!UiService.SafeRegexIsMatch(cn, locator.ClassNameRegex))
                    continue;
            }

            // ProcessId
            if (locator.ProcessId.HasValue)
            {
                var pid = UiService.SafeProcessId(c);
                if (pid != locator.ProcessId.Value)
                    continue;
            }

            // ControlId (Requirement 4 & 6)
            if (locator.ControlId.HasValue)
            {
                var ctrlId = SafeControlId(c);
                if (ctrlId != locator.ControlId.Value)
                    continue;
            }

            // FrameworkId
            if (!string.IsNullOrWhiteSpace(locator.FrameworkId))
            {
                var fwId = UiService.SafeFrameworkId(c);
                if (!string.Equals(fwId, locator.FrameworkId, StringComparison.OrdinalIgnoreCase))
                    continue;
            }

            // RuntimeId
            if (!string.IsNullOrWhiteSpace(locator.RuntimeId))
            {
                var rtId = UiService.SafeRuntimeIdString(c);
                if (!string.Equals(rtId, locator.RuntimeId, StringComparison.Ordinal))
                    continue;
            }

            // Name
            if (!string.IsNullOrWhiteSpace(locator.Name))
            {
                var name = UiService.SafeElementName(c);
                var mode = locator.NameMatchMode ?? globalMode;
                if (!UiService.StringMatchesByMode(name, locator.Name, mode, caseSensitive: true))
                    continue;
            }

            // NameRegex
            if (!string.IsNullOrWhiteSpace(locator.NameRegex))
            {
                var name = UiService.SafeElementName(c);
                if (!UiService.SafeRegexIsMatch(name, locator.NameRegex))
                    continue;
            }

            // AutomationId
            if (!string.IsNullOrWhiteSpace(locator.AutomationId))
            {
                var aid = UiService.SafeElementAutomationId(c);
                var mode = locator.AutomationIdMatchMode ?? globalMode;
                if (!UiService.StringMatchesByMode(aid, locator.AutomationId, mode, caseSensitive: true))
                    continue;
            }

            // AutomationIdRegex
            if (!string.IsNullOrWhiteSpace(locator.AutomationIdRegex))
            {
                var aid = UiService.SafeElementAutomationId(c);
                if (!UiService.SafeRegexIsMatch(aid, locator.AutomationIdRegex))
                    continue;
            }

            // Value
            if (!string.IsNullOrWhiteSpace(locator.Value))
            {
                var val = UiService.SafeElementValue(c);
                var mode = locator.ValueMatchMode ?? globalMode;
                if (!UiService.StringMatchesByMode(val, locator.Value, mode, caseSensitive: false))
                    continue;
            }

            // ValueRegex
            if (!string.IsNullOrWhiteSpace(locator.ValueRegex))
            {
                var val = UiService.SafeElementValue(c);
                if (!UiService.SafeRegexIsMatch(val, locator.ValueRegex))
                    continue;
            }

            // Text
            if (!string.IsNullOrWhiteSpace(locator.Text))
            {
                var text = UiService.SafeElementText(c);
                var mode = locator.TextMatchMode ?? globalMode;
                if (!UiService.StringMatchesByMode(text, locator.Text, mode, caseSensitive: false))
                    continue;
            }

            // Visible
            if (locator.Visible.HasValue)
            {
                var offscreen = UiService.SafeIsOffscreen(c) ?? false;
                var visible = !offscreen;
                if (visible != locator.Visible.Value)
                    continue;
            }

            // Enabled
            if (locator.Enabled.HasValue)
            {
                var enabled = UiService.SafeIsEnabled(c) ?? false;
                if (enabled != locator.Enabled.Value)
                    continue;
            }

            // Offscreen
            if (locator.Offscreen.HasValue)
            {
                var offscreen = UiService.SafeIsOffscreen(c) ?? false;
                if (offscreen != locator.Offscreen.Value)
                    continue;
            }

            // IncludeOffscreen (as a filter)
            if (locator.IncludeOffscreen.HasValue && locator.IncludeOffscreen.Value == false)
            {
                var offscreen = UiService.SafeIsOffscreen(c) ?? false;
                if (offscreen)
                    continue;
            }

            filtered.Add(c);
        }

        return filtered;
    }

    private List<ElementCandidate> BuildCandidateDiagnostics(List<AutomationElement> candidates, UiLocator locator)
    {
        var list = new List<ElementCandidate>();
        for (int i = 0; i < candidates.Count; i++)
        {
            var c = candidates[i];
            var rectangleObj = UiService.SafeBoundingRectangleObject(c);

            list.Add(new ElementCandidate
            {
                Index = i,
                Name = UiService.SafeElementName(c),
                AutomationId = UiService.SafeElementAutomationId(c),
                ClassName = UiService.SafeElementClassName(c),
                ControlType = UiService.SafeElementControlType(c),
                FrameworkId = UiService.SafeFrameworkId(c),
                RuntimeId = UiService.SafeRuntimeIdString(c),
                ProcessId = UiService.SafeProcessId(c),
                Hwnd = c.Properties.NativeWindowHandle.ValueOrDefault != IntPtr.Zero ? c.Properties.NativeWindowHandle.ValueOrDefault.ToInt64() : (long?)null,
                Rectangle = rectangleObj,
                Value = UiService.SafeElementValue(c),
                Text = UiService.SafeElementText(c),
                IsEnabled = UiService.SafeIsEnabled(c),
                IsOffscreen = UiService.SafeIsOffscreen(c),
                Score = 0,
                Reason = "Candidate checked during resolution."
            });
        }
        return list;
    }
}
