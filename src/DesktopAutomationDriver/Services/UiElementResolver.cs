using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using System.Threading;
using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Models.Resolver;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using Microsoft.Extensions.Logging;

namespace DesktopAutomationDriver.Services;

public sealed class UiElementResolver
{
    private readonly IUiSessionContext _sessionCtx;
    private readonly ILogger<UiElementResolver> _logger;
    private readonly Func<AutomationSession, bool, AutomationElement> _getWindowRoot;

    public UiElementResolver(
        IUiSessionContext sessionCtx,
        ILogger<UiElementResolver> logger,
        Func<AutomationSession, bool, AutomationElement> getWindowRoot)
    {
        _sessionCtx = sessionCtx;
        _logger = logger;
        _getWindowRoot = getWindowRoot;
    }

    private static string GetSearchRootName(UiRequest request)
    {
        return request.SearchRoot ?? (request.UseDesktopRoot == true ? "desktop" : (request.UseActiveWindowRoot == true ? "foreground" : "currentWindow"));
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

    public ResolvedElement ResolveOne(
        UiRequest request,
        UiLocator? locator = null,
        AutomationElement? explicitRoot = null,
        string? purpose = null)
    {
        var candidates = ResolveMany(request, locator, explicitRoot, purpose);

        if (candidates.Count == 0)
        {
            var session = RequireSession();
            var targetLocator = locator ?? request.Locator ?? new UiLocator();
            var root = explicitRoot ?? ResolveSearchRoot(request, session);
            var diag = BuildNotFoundDiagnostics(request, locator, explicitRoot, purpose);
            var rawCandidates = CollectCandidates(root, targetLocator, request, diag);
            diag.CandidateCount = rawCandidates.Count;

            var returnCandidates = request.ReturnCandidates == true || request.Debug == true || request.IncludeDiagnostics == true;
            if (returnCandidates)
            {
                for (int i = 0; i < Math.Min(20, rawCandidates.Count); i++)
                {
                    var c = rawCandidates[i];
                    diag.Candidates.Add(new ResolvedElement
                    {
                        Element = c,
                        Strategy = "candidate",
                        Score = 0,
                        Index = i
                    });
                }
            }

            var errMessage = FormatNotFoundError(diag, request, locator);
            throw new InvalidOperationException(errMessage);
        }

        var ambiguity = request.Ambiguity ?? "error";

        if (candidates.Count > 1)
        {
            if (string.Equals(ambiguity, "error", StringComparison.OrdinalIgnoreCase))
            {
                var diag = BuildAmbiguityDiagnostics(request, locator, candidates);
                var errMessage = FormatAmbiguityError(diag, request, locator);
                throw new InvalidOperationException(errMessage);
            }
            else if (string.Equals(ambiguity, "first", StringComparison.OrdinalIgnoreCase))
            {
                return candidates[0];
            }
            else if (string.Equals(ambiguity, "all", StringComparison.OrdinalIgnoreCase))
            {
                var diag = BuildResolveDiagnostics(request, locator, candidates, "all");
                var first = candidates[0];
                return new ResolvedElement
                {
                    Element = first.Element,
                    Strategy = first.Strategy,
                    Score = first.Score,
                    Index = first.Index,
                    Snapshot = diag
                };
            }
        }

        return candidates[0];
    }

    public List<ResolvedElement> ResolveMany(
        UiRequest request,
        UiLocator? locator = null,
        AutomationElement? explicitRoot = null,
        string? purpose = null)
    {
        var session = RequireSession();
        var targetLocator = locator ?? request.Locator ?? new UiLocator();

        // Determine Search Root
        var root = explicitRoot ?? ResolveSearchRoot(request, session);

        var diagnostics = new ResolveDiagnostics
        {
            SearchRoot = GetSearchRootName(request),
            TreeView = request.TreeView ?? "control",
            Backend = request.Backend ?? "uia",
            Strategy = purpose ?? "resolve-many"
        };

        // Collect Candidates
        var rawCandidates = CollectCandidates(root, targetLocator, request, diagnostics);

        // Apply ctrlIndex BEFORE applying filters
        if (targetLocator.CtrlIndex.HasValue)
        {
            var idx = targetLocator.CtrlIndex.Value;
            if (idx >= 0 && idx < rawCandidates.Count)
            {
                rawCandidates = new List<AutomationElement> { rawCandidates[idx] };
            }
            else
            {
                rawCandidates = new List<AutomationElement>();
            }
        }

        // Apply filters in sequence
        var filteredCandidates = ApplyFilters(rawCandidates, targetLocator, request);

        // Scoring/Best-match
        var useBestMatch = request.AllowBestMatch == true || !string.IsNullOrWhiteSpace(request.BestMatch) || !string.IsNullOrWhiteSpace(targetLocator.BestMatch);
        var targetPattern = targetLocator.BestMatch ?? request.BestMatch ?? string.Empty;

        var resolvedList = new List<ResolvedElement>();
        for (int i = 0; i < filteredCandidates.Count; i++)
        {
            var c = filteredCandidates[i];
            var score = useBestMatch ? CalculateScore(c, targetPattern, targetLocator) : 100;

            resolvedList.Add(new ResolvedElement
            {
                Element = c,
                Strategy = useBestMatch ? "best-match" : "exact-match",
                Score = score,
                Index = i
            });
        }

        if (useBestMatch)
        {
            resolvedList = resolvedList.OrderByDescending(r => r.Score).ToList();
        }

        // Apply foundIndex after filtering/scoring
        if (targetLocator.FoundIndex.HasValue)
        {
            var idx = targetLocator.FoundIndex.Value;
            if (idx >= 0 && idx < resolvedList.Count)
            {
                resolvedList = new List<ResolvedElement> { resolvedList[idx] };
            }
            else
            {
                resolvedList = new List<ResolvedElement>();
            }
        }

        // Limit results to maxMatches
        var maxMatches = request.MaxMatches ?? request.Limit ?? 50;
        if (resolvedList.Count > maxMatches)
        {
            resolvedList = resolvedList.Take(maxMatches).ToList();
        }

        return resolvedList;
    }

    public List<ResolvedElement> ResolvePath(UiRequest request)
    {
        var path = request.LocatorPath ?? request.Criteria;
        if (path == null || path.Count == 0)
        {
            return ResolveMany(request, request.Locator);
        }

        var session = RequireSession();
        AutomationElement root = ResolveSearchRoot(request, session);
        var currentRoots = new List<AutomationElement> { root };

        for (var i = 0; i < path.Count; i++)
        {
            var stepLocator = path[i];
            var nextRoots = new List<AutomationElement>();

            foreach (var stepRoot in currentRoots)
            {
                var stepRequest = CloneRequestForStep(request, stepLocator);
                var resolved = ResolveMany(
                    stepRequest,
                    stepLocator,
                    explicitRoot: stepRoot,
                    purpose: $"locatorPath[{i}]");

                foreach (var r in resolved)
                {
                    if (r.Element != null)
                    {
                        nextRoots.Add(r.Element);
                    }
                }
            }

            currentRoots = nextRoots.Distinct().ToList();
            if (currentRoots.Count == 0)
                break;
        }

        return currentRoots.Select((el, idx) => new ResolvedElement
        {
            Element = el,
            Strategy = "locator-path",
            Score = 100,
            Index = idx
        }).ToList();
    }

    public ResolvedElement ResolveOne(UiRequest request)
    {
        var candidates = ResolveMany(request);

        if (candidates.Count == 0)
            throw BuildNotFoundException(request, candidates);

        if (candidates.Count > 1)
        {
            var ambiguity = request.Ambiguity ?? "error";

            if (ambiguity.Equals("first", StringComparison.OrdinalIgnoreCase))
                return candidates[0];

            throw BuildAmbiguousException(request, candidates);
        }

        return candidates[0];
    }

    public IReadOnlyList<ResolvedElement> ResolveMany(UiRequest request)
    {
        return ResolveMany(request, request.Locator, null, null);
    }

    public ResolvedElement ResolveLocatorPath(UiRequest request)
    {
        var path = request.LocatorPath ?? request.Criteria;

        if (path == null || path.Count == 0)
        {
            return ResolveOne(request);
        }

        var root = ResolveSearchRoot(request);

        foreach (var step in path)
        {
            var stepCandidates = ResolveManyWithinRoot(root, step);

            if (stepCandidates.Count == 0)
                throw BuildPathNotFoundException(request, step);

            if (stepCandidates.Count > 1 && step.FoundIndex == null && step.CtrlIndex == null)
                throw BuildPathAmbiguousException(request, step, stepCandidates);

            root = stepCandidates.First().Element;
        }

        return new ResolvedElement
        {
            Element = root,
            Strategy = "locatorPath"
        };
    }

    internal AutomationElement ResolveSearchRoot(UiRequest request, AutomationSession session)
    {
        var searchRoot = request.SearchRoot;
        if (request.UseDesktopRoot == true)
        {
            searchRoot = "desktop";
        }
        else if (request.UseActiveWindowRoot == true)
        {
            searchRoot = "foreground";
        }

        if (string.IsNullOrWhiteSpace(searchRoot))
        {
            searchRoot = "currentWindow";
        }

        switch (searchRoot.ToLowerInvariant())
        {
            case "desktop":
                return session.Automation.GetDesktop();

            case "foreground":
                return GetForegroundWindowElement(session.Automation);

            case "activepopup":
                return GetActivePopupRoot(session);

            case "parent":
                if (request.ParentLocator != null)
                {
                    var parentRequest = new UiRequest
                    {
                        Operation = request.Operation,
                        Locator = request.ParentLocator,
                        SearchRoot = "currentWindow",
                        TreeView = request.TreeView,
                        Backend = request.Backend,
                        Ambiguity = "first"
                    };
                    try
                    {
                        var resolvedParent = ResolveOne(parentRequest, request.ParentLocator);
                        if (resolvedParent?.Element != null)
                        {
                            return resolvedParent.Element;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "Failed to resolve parent locator. Falling back to window root.");
                    }
                }
                return _getWindowRoot(session, false);

            case "currentwindow":
            default:
                return _getWindowRoot(session, false);
        }
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

    internal List<AutomationElement> CollectCandidates(
        AutomationElement root,
        UiLocator locator,
        UiRequest request,
        ResolveDiagnostics diagnostics)
    {
        var session = RequireSession();

        // Direct HWND
        if (locator.Hwnd.HasValue)
        {
            try
            {
                var hwndPtr = new IntPtr(locator.Hwnd.Value);
                var element = session.Automation.FromHandle(hwndPtr);
                if (element != null)
                {
                    diagnostics.Strategy = "direct-hwnd";
                    return new List<AutomationElement> { element };
                }
            }
            catch (Exception ex)
            {
                diagnostics.Errors.Add($"HWND resolution failed: {ex.Message}");
            }
            return new List<AutomationElement>();
        }

        var depth = locator.Depth ?? request.Locator?.Depth ?? 20;
        var topLevelOnly = locator.TopLevelOnly == true;

        List<AutomationElement> rawCandidates = new List<AutomationElement>();

        var treeView = request.TreeView?.ToLowerInvariant() ?? "control";
        diagnostics.TreeView = treeView;

        if (string.Equals(treeView, "raw", StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                var walker = session.Automation.TreeWalkerFactory.GetRawViewWalker();
                if (walker != null)
                {
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
                    diagnostics.Strategy = "raw-walker";
                }
                else
                {
                    throw new InvalidOperationException("Raw walker not returned by factory.");
                }
            }
            catch (Exception ex)
            {
                diagnostics.Errors.Add($"raw tree requested but raw walker unavailable; used control descendants fallback. Detail: {ex.Message}");
                treeView = "control";
            }
        }

        if (string.Equals(treeView, "content", StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                var dummyProperty = session.Automation.PropertyLibrary.Element.IsContentElement;
                if (dummyProperty != null)
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
                    diagnostics.Strategy = "content-walker";
                }
                else
                {
                    throw new InvalidOperationException("IsContentElement property not in library.");
                }
            }
            catch (Exception ex)
            {
                diagnostics.Errors.Add($"content tree requested but IsContentElement not available; used control descendants fallback. Detail: {ex.Message}");
                treeView = "control";
            }
        }

        if (string.Equals(treeView, "control", StringComparison.OrdinalIgnoreCase))
        {
            if (topLevelOnly)
            {
                rawCandidates = root.FindAllChildren().ToList();
                diagnostics.Strategy = "control-top-level";
            }
            else
            {
                if (!string.IsNullOrWhiteSpace(locator.ControlType))
                {
                    try
                    {
                        var ct = UiService.ParseControlType(locator.ControlType);
                        rawCandidates = root.FindAllDescendants(session.Automation.ConditionFactory.ByControlType(ct)).ToList();
                        diagnostics.Strategy = "control-type-descendants";
                    }
                    catch
                    {
                        rawCandidates = UiService.FindDescendantsUpToDepth(root, depth);
                        diagnostics.Strategy = "control-descendants-fallback";
                    }
                }
                else
                {
                    rawCandidates = UiService.FindDescendantsUpToDepth(root, depth);
                    diagnostics.Strategy = "control-descendants";
                }
            }
        }

        // Active only filter
        if (locator.ActiveOnly == true)
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
        bool allowedOffscreen = true;
        if (locator.IncludeOffscreen.HasValue)
        {
            allowedOffscreen = locator.IncludeOffscreen.Value;
        }
        else
        {
            var op = request.Operation?.ToLowerInvariant();
            var isAction = op is "click" or "doubleclick" or "rightclick" or "hover" or "type" or "clear" or "select" or "selectopendropdownitem" or "clickopendropdownitem";
            if (isAction)
            {
                allowedOffscreen = false;
            }
        }

        if (!allowedOffscreen)
        {
            rawCandidates = rawCandidates.Where(c => UiService.SafeIsOffscreen(c) != true).ToList();
        }

        return rawCandidates;
    }

    private List<AutomationElement> ApplyFilters(
        List<AutomationElement> candidates,
        UiLocator locator,
        UiRequest request)
    {
        var filtered = new List<AutomationElement>(candidates);
        var globalMode = locator.MatchMode ?? request.MatchMode;

        // 1. hwnd
        if (locator.Hwnd.HasValue)
        {
            filtered = filtered.Where(c =>
            {
                try
                {
                    var handle = c.Properties.NativeWindowHandle.ValueOrDefault;
                    return handle != IntPtr.Zero && handle.ToInt64() == locator.Hwnd.Value;
                }
                catch { return false; }
            }).ToList();
        }

        // 2. controlType
        if (!string.IsNullOrWhiteSpace(locator.ControlType))
        {
            filtered = filtered.Where(c =>
            {
                var ct = UiService.SafeElementControlType(c);
                return string.Equals(ct, locator.ControlType, StringComparison.OrdinalIgnoreCase);
            }).ToList();
        }

        // 3. className / classNameRegex
        if (!string.IsNullOrWhiteSpace(locator.ClassName))
        {
            var mode = locator.ClassNameMatchMode ?? globalMode;
            filtered = filtered.Where(c =>
            {
                var cn = UiService.SafeElementClassName(c);
                return MatchText(cn, locator.ClassName, null, mode);
            }).ToList();
        }
        if (!string.IsNullOrWhiteSpace(locator.ClassNameRegex))
        {
            filtered = filtered.Where(c =>
            {
                var cn = UiService.SafeElementClassName(c);
                return MatchText(cn, null, locator.ClassNameRegex, null);
            }).ToList();
        }

        // 4. processId
        if (locator.ProcessId.HasValue)
        {
            filtered = filtered.Where(c =>
            {
                var pid = UiService.SafeProcessId(c);
                return pid == locator.ProcessId.Value;
            }).ToList();
        }

        // 5. controlId
        if (locator.ControlId.HasValue)
        {
            filtered = filtered.Where(c =>
            {
                var ctrlId = SafeControlId(c);
                return ctrlId == locator.ControlId.Value;
            }).ToList();
        }

        // 6. visible / offscreen
        if (locator.Visible.HasValue)
        {
            filtered = filtered.Where(c =>
            {
                var offscreen = UiService.SafeIsOffscreen(c) ?? false;
                var visible = !offscreen;
                return visible == locator.Visible.Value;
            }).ToList();
        }
        if (locator.Offscreen.HasValue)
        {
            filtered = filtered.Where(c =>
            {
                var offscreen = UiService.SafeIsOffscreen(c) ?? false;
                return offscreen == locator.Offscreen.Value;
            }).ToList();
        }

        // 7. enabled
        if (locator.Enabled.HasValue)
        {
            filtered = filtered.Where(c =>
            {
                var enabled = UiService.SafeIsEnabled(c) ?? false;
                return enabled == locator.Enabled.Value;
            }).ToList();
        }

        // 8. name / nameRegex / nameMatchMode
        if (!string.IsNullOrWhiteSpace(locator.Name))
        {
            var mode = locator.NameMatchMode ?? globalMode;
            filtered = filtered.Where(c =>
            {
                var name = UiService.SafeElementName(c);
                return MatchText(name, locator.Name, null, mode);
            }).ToList();
        }
        if (!string.IsNullOrWhiteSpace(locator.NameRegex))
        {
            filtered = filtered.Where(c =>
            {
                var name = UiService.SafeElementName(c);
                return MatchText(name, null, locator.NameRegex, null);
            }).ToList();
        }

        // 9. automationId / automationIdRegex / automationIdMatchMode
        if (!string.IsNullOrWhiteSpace(locator.AutomationId))
        {
            var mode = locator.AutomationIdMatchMode ?? globalMode;
            filtered = filtered.Where(c =>
            {
                var aid = UiService.SafeElementAutomationId(c);
                return MatchText(aid, locator.AutomationId, null, mode);
            }).ToList();
        }
        if (!string.IsNullOrWhiteSpace(locator.AutomationIdRegex))
        {
            filtered = filtered.Where(c =>
            {
                var aid = UiService.SafeElementAutomationId(c);
                return MatchText(aid, null, locator.AutomationIdRegex, null);
            }).ToList();
        }

        // 10. frameworkId
        if (!string.IsNullOrWhiteSpace(locator.FrameworkId))
        {
            filtered = filtered.Where(c =>
            {
                var fw = UiService.SafeFrameworkId(c);
                return string.Equals(fw, locator.FrameworkId, StringComparison.OrdinalIgnoreCase);
            }).ToList();
        }

        // 11. runtimeId
        if (!string.IsNullOrWhiteSpace(locator.RuntimeId))
        {
            filtered = filtered.Where(c =>
            {
                var rt = UiService.SafeRuntimeIdString(c);
                return string.Equals(rt, locator.RuntimeId, StringComparison.Ordinal);
            }).ToList();
        }

        // 12. value / valueRegex / valueMatchMode
        if (!string.IsNullOrWhiteSpace(locator.Value))
        {
            var mode = locator.ValueMatchMode ?? globalMode;
            filtered = filtered.Where(c =>
            {
                var val = UiService.SafeElementValue(c);
                return MatchText(val, locator.Value, null, mode);
            }).ToList();
        }
        if (!string.IsNullOrWhiteSpace(locator.ValueRegex))
        {
            filtered = filtered.Where(c =>
            {
                var val = UiService.SafeElementValue(c);
                return MatchText(val, null, locator.ValueRegex, null);
            }).ToList();
        }

        // 13. text / textMatchMode
        if (!string.IsNullOrWhiteSpace(locator.Text))
        {
            var mode = locator.TextMatchMode ?? globalMode;
            filtered = filtered.Where(c =>
            {
                var txt = UiService.SafeElementText(c);
                return MatchText(txt, locator.Text, null, mode);
            }).ToList();
        }

        // 14. Rectangle filters
        filtered = filtered.Where(c => MatchesRectangleFilters(c, locator)).ToList();

        return filtered;
    }

    private static bool MatchText(
        string? actual,
        string? expected,
        string? regex,
        string? matchMode)
    {
        if (!string.IsNullOrWhiteSpace(regex))
        {
            if (actual == null) return false;
            try
            {
                return Regex.IsMatch(actual, regex, RegexOptions.IgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        if (expected == null) return true;
        if (actual == null) return false;

        var mode = matchMode?.ToLowerInvariant() ?? "exact";
        switch (mode)
        {
            case "exact":
            case "ignorecase":
                return string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase);
            case "contains":
                return actual.Contains(expected, StringComparison.OrdinalIgnoreCase);
            case "startswith":
                return actual.StartsWith(expected, StringComparison.OrdinalIgnoreCase);
            case "endswith":
                return actual.EndsWith(expected, StringComparison.OrdinalIgnoreCase);
            case "regex":
                try
                {
                    return Regex.IsMatch(actual, expected, RegexOptions.IgnoreCase);
                }
                catch { return false; }
            default:
                return string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase);
        }
    }

    private static int CalculateScore(AutomationElement element, string targetPattern, UiLocator locator)
    {
        int score = 0;
        if (!string.IsNullOrWhiteSpace(targetPattern))
        {
            var name = UiService.SafeElementName(element) ?? string.Empty;
            var aid = UiService.SafeElementAutomationId(element) ?? string.Empty;
            var className = UiService.SafeElementClassName(element) ?? string.Empty;
            var controlType = UiService.SafeElementControlType(element) ?? string.Empty;
            var value = UiService.SafeElementValue(element) ?? string.Empty;
            var text = UiService.SafeElementText(element) ?? string.Empty;

            // 1. Exact matches with specific weights (ordered descending):
            // AutomationId exact: 100
            int exactScore = 0;
            if (string.Equals(aid, targetPattern, StringComparison.OrdinalIgnoreCase))
            {
                exactScore = 100;
            }
            // name exact: 95
            else if (string.Equals(name, targetPattern, StringComparison.OrdinalIgnoreCase))
            {
                exactScore = 95;
            }
            // value exact: 90
            else if (string.Equals(value, targetPattern, StringComparison.OrdinalIgnoreCase))
            {
                exactScore = 90;
            }
            // text exact: 85
            else if (string.Equals(text, targetPattern, StringComparison.OrdinalIgnoreCase))
            {
                exactScore = 85;
            }

            // 2. Contains match fallback scoring (ordered descending):
            int containsScore = 0;
            if (aid.Contains(targetPattern, StringComparison.OrdinalIgnoreCase))
            {
                containsScore = 50;
            }
            else if (name.Contains(targetPattern, StringComparison.OrdinalIgnoreCase))
            {
                containsScore = 45;
            }
            else if (value.Contains(targetPattern, StringComparison.OrdinalIgnoreCase))
            {
                containsScore = 40;
            }
            else if (text.Contains(targetPattern, StringComparison.OrdinalIgnoreCase))
            {
                containsScore = 35;
            }

            // 3. Token overlap fallback scoring
            int tokenScore = 0;
            var targetTokens = targetPattern.Split(new[] { ' ', '_', '-', '.', '/' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var field in new[] { aid, name, value, text })
            {
                if (string.IsNullOrWhiteSpace(field)) continue;
                var fieldTokens = field.Split(new[] { ' ', '_', '-', '.', '/' }, StringSplitOptions.RemoveEmptyEntries);
                int overlapCount = targetTokens.Count(tToken => fieldTokens.Any(fToken => string.Equals(fToken, tToken, StringComparison.OrdinalIgnoreCase)));
                if (overlapCount > 0)
                {
                    tokenScore = Math.Max(tokenScore, 10 + overlapCount * 5);
                }
            }

            // Select the highest score from the active matching tier: exact match has priority over contains, which has priority over token overlap.
            if (exactScore > 0)
            {
                score = exactScore;
            }
            else if (containsScore > 0)
            {
                score = containsScore;
            }
            else
            {
                score = tokenScore;
            }
        }
        else
        {
            score = 50;
        }

        // Modifiers:
        // controlType match: +20
        if (!string.IsNullOrWhiteSpace(locator.ControlType))
        {
            var elCt = UiService.SafeElementControlType(element);
            if (string.Equals(elCt, locator.ControlType, StringComparison.OrdinalIgnoreCase))
            {
                score += 20;
            }
        }

        // className match: +15
        if (!string.IsNullOrWhiteSpace(locator.ClassName))
        {
            var elCn = UiService.SafeElementClassName(element);
            if (string.Equals(elCn, locator.ClassName, StringComparison.OrdinalIgnoreCase))
            {
                score += 15;
            }
        }

        // same row / near point match: +10
        if (locator.NearX.HasValue || locator.NearY.HasValue || locator.Left.HasValue || locator.Top.HasValue)
        {
            score += 10;
        }

        return score;
    }

    private UiRequest CloneRequestForStep(UiRequest request, UiLocator stepLocator)
    {
        return new UiRequest
        {
            Operation = request.Operation,
            Locator = stepLocator,
            TreeView = request.TreeView,
            Backend = request.Backend,
            Ambiguity = request.Ambiguity ?? "error",
            MaxMatches = request.MaxMatches,
            ReturnCandidates = request.ReturnCandidates,
            IncludeDiagnostics = request.IncludeDiagnostics,
            AllowBestMatch = request.AllowBestMatch,
            BestMatch = stepLocator.BestMatch ?? request.BestMatch,
            TimeoutMs = request.TimeoutMs
        };
    }

    private ResolveDiagnostics BuildNotFoundDiagnostics(
        UiRequest request,
        UiLocator? locator,
        AutomationElement? explicitRoot,
        string? purpose)
    {
        return new ResolveDiagnostics
        {
            SearchRoot = GetSearchRootName(request),
            TreeView = request.TreeView ?? "control",
            Backend = request.Backend ?? "uia",
            Strategy = purpose ?? "element-not-found"
        };
    }

    private string FormatNotFoundError(ResolveDiagnostics diag, UiRequest request, UiLocator? locator)
    {
        var targetLoc = locator ?? request.Locator;
        var summary = UiService.DescribeLocator(targetLoc ?? new UiLocator());
        var msg = $"ElementNotFound: No matching elements found.\n" +
                  $"Operation: {request.Operation}\n" +
                  $"Locator: {summary}\n" +
                  $"SearchRoot: {diag.SearchRoot}\n" +
                  $"TreeView: {diag.TreeView}\n" +
                  $"Backend: {diag.Backend}\n" +
                  $"CandidateCount: {diag.CandidateCount}\n";

        if (diag.Candidates != null && diag.Candidates.Count > 0)
        {
            msg += "Candidates:\n";
            foreach (var c in diag.Candidates)
            {
                msg += $"  - Name: {UiService.SafeElementName(c.Element)}, ControlType: {UiService.SafeElementControlType(c.Element)}, AutomationId: {UiService.SafeElementAutomationId(c.Element)}, ClassName: {UiService.SafeElementClassName(c.Element)}\n";
            }
        }
        else
        {
            msg += "FiltersApplied: hwnd, controlType, className, processId, controlId, visible, enabled, name, automationId, frameworkId, runtimeId, value, text, bestMatch, foundIndex";
        }

        return msg;
    }

    private ResolveDiagnostics BuildAmbiguityDiagnostics(
        UiRequest request,
        UiLocator? locator,
        List<ResolvedElement> candidates)
    {
        var diag = new ResolveDiagnostics
        {
            SearchRoot = GetSearchRootName(request),
            TreeView = request.TreeView ?? "control",
            Backend = request.Backend ?? "uia",
            Strategy = "element-ambiguous",
            CandidateCount = candidates.Count
        };
        foreach (var c in candidates)
        {
            diag.Candidates.Add(c);
        }
        return diag;
    }

    private ResolveDiagnostics BuildResolveDiagnostics(
        UiRequest request,
        UiLocator? locator,
        List<ResolvedElement> candidates,
        string strategy)
    {
        var diag = new ResolveDiagnostics
        {
            SearchRoot = GetSearchRootName(request),
            TreeView = request.TreeView ?? "control",
            Backend = request.Backend ?? "uia",
            Strategy = strategy,
            CandidateCount = candidates.Count
        };
        foreach (var c in candidates)
        {
            diag.Candidates.Add(c);
        }
        return diag;
    }

    private string FormatAmbiguityError(ResolveDiagnostics diag, UiRequest request, UiLocator? locator)
    {
        var suggestions = new List<string>();
        var targetLoc = locator ?? request.Locator;
        if (targetLoc != null)
        {
            if (!targetLoc.FoundIndex.HasValue)
                suggestions.Add("add foundIndex (e.g. 0 to get the first match)");
            if (request.ParentLocator == null)
                suggestions.Add("add parentLocator to narrow the search scope");
            if (request.LocatorPath == null && request.Criteria == null)
                suggestions.Add("add locatorPath/criteria for hierarchical matching");
            if (string.IsNullOrEmpty(targetLoc.ClassName))
                suggestions.Add("add className filter");
            if (string.IsNullOrEmpty(targetLoc.ControlType))
                suggestions.Add("add controlType filter");
        }

        var msg = $"ElementAmbiguous: Multiple matching elements found ({diag.CandidateCount}).\n" +
                  $"Operation: {request.Operation}\n" +
                  $"Locator: {UiService.DescribeLocator(targetLoc ?? new UiLocator())}\n" +
                  $"MatchCount: {diag.CandidateCount}\n" +
                  $"Ambiguity: {request.Ambiguity ?? "error"}\n" +
                  $"Suggestions:\n  " + string.Join("\n  ", suggestions);

        if (diag.Candidates != null && diag.Candidates.Count > 0)
        {
            msg += "\nMatching Candidates:\n";
            foreach (var c in diag.Candidates)
            {
                msg += $"  - Name: {UiService.SafeElementName(c.Element)}, ControlType: {UiService.SafeElementControlType(c.Element)}, AutomationId: {UiService.SafeElementAutomationId(c.Element)}, ClassName: {UiService.SafeElementClassName(c.Element)}\n";
            }
        }

        return msg;
    }

    internal AutomationElement ResolveSearchRoot(UiRequest request)
    {
        return ResolveSearchRoot(request, RequireSession());
    }

    private IReadOnlyList<ResolvedElement> ResolveManyWithinRoot(AutomationElement root, UiLocator step)
    {
        var stepRequest = new UiRequest { Locator = step };
        return ResolveMany(stepRequest, step, root);
    }

    private UiResolutionException BuildNotFoundException(UiRequest request, IReadOnlyList<ResolvedElement> candidates)
    {
        var locator = request.Locator;
        var message = $"ElementNotFound: No matching elements found for locator={UiService.DescribeLocator(locator ?? new UiLocator())}";
        var suggestions = new List<string>
        {
            "try automationId only",
            "try matchMode contains",
            "try useActiveWindowRoot true",
            "try dumptree"
        };
        return new UiResolutionException("not-found", message, locator, candidates, suggestions);
    }

    private UiResolutionException BuildAmbiguousException(UiRequest request, IReadOnlyList<ResolvedElement> candidates)
    {
        var locator = request.Locator;
        var message = "ElementAmbiguous: Multiple matching elements found.";
        var suggestions = new List<string>
        {
            "add foundIndex",
            "add parentLocator",
            "add locatorPath",
            "add rectangle filters",
            "use dumptree to inspect stable identifiers"
        };
        return new UiResolutionException("ambiguous", message, locator, candidates, suggestions);
    }

    private UiResolutionException BuildPathNotFoundException(UiRequest request, UiLocator step)
    {
        var message = $"ElementNotFound: No matching elements found for step={UiService.DescribeLocator(step)} in locatorPath.";
        var suggestions = new List<string>
        {
            "check step locator attributes",
            "try matchMode contains on step"
        };
        return new UiResolutionException("not-found", message, step, new List<ResolvedElement>(), suggestions);
    }

    private UiResolutionException BuildPathAmbiguousException(UiRequest request, UiLocator step, IReadOnlyList<ResolvedElement> candidates)
    {
        var message = $"ElementAmbiguous: Multiple matching elements found for step={UiService.DescribeLocator(step)} in locatorPath.";
        var suggestions = new List<string>
        {
            "add foundIndex to step locator",
            "add parentLocator",
            "use dumptree"
        };
        return new UiResolutionException("ambiguous", message, step, candidates, suggestions);
    }

    private bool MatchesRectangleFilters(AutomationElement element, UiLocator locator)
    {
        var r = element.BoundingRectangle;
        var tolerance = locator.Tolerance ?? 0;

        if (locator.Left.HasValue && Math.Abs(r.Left - locator.Left.Value) > tolerance)
            return false;

        if (locator.Top.HasValue && Math.Abs(r.Top - locator.Top.Value) > tolerance)
            return false;

        if (locator.Right.HasValue && Math.Abs(r.Right - locator.Right.Value) > tolerance)
            return false;

        if (locator.Bottom.HasValue && Math.Abs(r.Bottom - locator.Bottom.Value) > tolerance)
            return false;

        if (locator.Width.HasValue && Math.Abs(r.Width - locator.Width.Value) > tolerance)
            return false;

        if (locator.Height.HasValue && Math.Abs(r.Height - locator.Height.Value) > tolerance)
            return false;

        if (locator.NearX.HasValue && locator.NearY.HasValue)
        {
            var nearTolerance = locator.Tolerance ?? 5;

            var containsNearPoint =
                locator.NearX.Value >= r.Left - nearTolerance &&
                locator.NearX.Value <= r.Right + nearTolerance &&
                locator.NearY.Value >= r.Top - nearTolerance &&
                locator.NearY.Value <= r.Bottom + nearTolerance;

            if (!containsNearPoint)
                return false;
        }

        if (locator.ContainsPoint == true && locator.NearX.HasValue && locator.NearY.HasValue)
        {
            var containsPoint =
                locator.NearX.Value >= r.Left &&
                locator.NearX.Value <= r.Right &&
                locator.NearY.Value >= r.Top &&
                locator.NearY.Value <= r.Bottom;

            if (!containsPoint)
                return false;
        }

        if (locator.IntersectsRectangle == true && locator.Left.HasValue && locator.Top.HasValue && locator.Right.HasValue && locator.Bottom.HasValue)
        {
            var intersects =
                !(r.Left > locator.Right.Value ||
                  r.Right < locator.Left.Value ||
                  r.Top > locator.Bottom.Value ||
                  r.Bottom < locator.Top.Value);

            if (!intersects)
                return false;
        }

        return true;
    }
}
