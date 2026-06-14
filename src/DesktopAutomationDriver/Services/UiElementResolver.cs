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
            var diag = BuildNotFoundDiagnostics(request, locator, explicitRoot, purpose);
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
            var score = useBestMatch ? CalculateScore(c, targetPattern) : 100;

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

    public ResolvedElement ResolveLocatorPath(UiRequest request)
    {
        var path = request.LocatorPath ?? request.Criteria;

        if (path == null || path.Count == 0)
        {
            return ResolveOne(request, request.Locator);
        }

        var session = RequireSession();
        AutomationElement? root = ResolveSearchRoot(request, session);
        ResolvedElement? current = null;

        for (var i = 0; i < path.Count; i++)
        {
            var stepLocator = path[i];

            var stepRequest = CloneRequestForStep(request, stepLocator);
            var stepRoot = current?.Element ?? root;

            current = ResolveOne(
                stepRequest,
                stepLocator,
                explicitRoot: stepRoot,
                purpose: $"locatorPath[{i}]");
        }

        if (current == null)
        {
            throw new InvalidOperationException("Locator path resolution failed to yield an element.");
        }
        return current;
    }

    private AutomationElement ResolveSearchRoot(UiRequest request, AutomationSession session)
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

    private static int CalculateScore(AutomationElement element, string targetPattern)
    {
        if (string.IsNullOrWhiteSpace(targetPattern)) return 0;

        var name = UiService.SafeElementName(element);
        var aid = UiService.SafeElementAutomationId(element);
        var className = UiService.SafeElementClassName(element);
        var controlType = UiService.SafeElementControlType(element);
        var value = UiService.SafeElementValue(element);
        var text = UiService.SafeElementText(element);

        var candidateFields = new[] { name, aid, className, controlType, value, text };

        int maxScore = 0;
        foreach (var field in candidateFields)
        {
            if (string.IsNullOrWhiteSpace(field)) continue;

            int fieldScore = 0;
            if (string.Equals(field, targetPattern, StringComparison.Ordinal))
            {
                fieldScore = 100;
            }
            else if (string.Equals(field, targetPattern, StringComparison.OrdinalIgnoreCase))
            {
                fieldScore = 90;
            }
            else if (field.StartsWith(targetPattern, StringComparison.OrdinalIgnoreCase))
            {
                fieldScore = 75;
            }
            else if (field.Contains(targetPattern, StringComparison.OrdinalIgnoreCase))
            {
                fieldScore = 60;
            }
            else
            {
                var fieldTokens = field.Split(new[] { ' ', '_', '-', '.', '/' }, StringSplitOptions.RemoveEmptyEntries);
                var targetTokens = targetPattern.Split(new[] { ' ', '_', '-', '.', '/' }, StringSplitOptions.RemoveEmptyEntries);
                int overlapCount = 0;
                foreach (var tToken in targetTokens)
                {
                    if (fieldTokens.Any(fToken => string.Equals(fToken, tToken, StringComparison.OrdinalIgnoreCase)))
                    {
                        overlapCount++;
                    }
                }
                fieldScore = overlapCount * 10;
            }

            if (fieldScore > maxScore)
            {
                maxScore = fieldScore;
            }
        }

        return maxScore;
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
        return $"ElementNotFound: No matching elements found.\n" +
               $"Operation: {request.Operation}\n" +
               $"Locator: {UiService.DescribeLocator(locator ?? request.Locator)}\n" +
               $"SearchRoot: {diag.SearchRoot}\n" +
               $"TreeView: {diag.TreeView}\n" +
               $"Backend: {diag.Backend}\n" +
               $"CandidateCount: 0\n" +
               $"FiltersApplied: hwnd, controlType, className, processId, controlId, visible, enabled, name, automationId, frameworkId, runtimeId, value, text, bestMatch, foundIndex";
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

        return $"ElementAmbiguous: Multiple matching elements found ({diag.CandidateCount}).\n" +
               $"Operation: {request.Operation}\n" +
               $"Locator: {UiService.DescribeLocator(targetLoc)}\n" +
               $"MatchCount: {diag.CandidateCount}\n" +
               $"Ambiguity: {request.Ambiguity ?? "error"}\n" +
               $"Suggestions:\n  " + string.Join("\n  ", suggestions);
    }
}
