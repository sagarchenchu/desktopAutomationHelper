using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using DesktopAutomationDriver.Models.Request;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using Microsoft.Extensions.Logging;

namespace DesktopAutomationDriver.Services.Resolution;

public sealed class ElementResolver
{
    private readonly IUiSessionContext _ctx;
    private readonly ILogger _logger;
    private readonly Func<AutomationSession, bool, AutomationElement>? _getWindowRoot;

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

    private const int DefaultNearPointTolerancePixels = 5;

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
        return _ctx.ActiveSession ?? throw new InvalidOperationException("No active automation session.");
    }

    public ElementMatchResult ResolveOne(
        UiRequest request,
        UiLocator locator,
        ElementSearchOptions options)
    {
        var matches = ResolveAll(request, locator, options);

        if (matches.Count == 0)
        {
            throw new InvalidOperationException(BuildNotFoundMessage(locator, options));
        }

        if (options.CtrlIndex.HasValue)
        {
            var index = options.CtrlIndex.Value;
            if (index < 0 || index >= matches.Count)
                throw new InvalidOperationException($"ctrlIndex {index} out of range. matches={matches.Count}");

            return matches[index];
        }

        if (options.FoundIndex.HasValue)
        {
            var index = options.FoundIndex.Value;
            if (index < 0 || index >= matches.Count)
                throw new InvalidOperationException($"foundIndex {index} out of range. matches={matches.Count}");

            return matches[index];
        }

        var best = matches.OrderByDescending(x => x.Score).First();

        if (options.ThrowIfAmbiguous)
        {
            var sameScore = matches.Where(x => x.Score == best.Score).Take(2).ToList();
            if (sameScore.Count > 1)
            {
                throw new InvalidOperationException(BuildAmbiguousMessage(locator, matches));
            }
        }

        return best;
    }

    public IReadOnlyList<ElementMatchResult> ResolveAll(
        UiRequest request,
        UiLocator locator,
        ElementSearchOptions options)
    {
        var session = RequireSession();
        var deadline = DateTime.UtcNow + options.Timeout;

        while (true)
        {
            var rawResults = ExecuteResolveAll(request, locator, options, session);
            if (rawResults.Count > 0)
            {
                return rawResults;
            }

            if (DateTime.UtcNow >= deadline)
            {
                return rawResults;
            }

            System.Threading.Thread.Sleep(options.PollInterval);
        }
    }

    private IReadOnlyList<ElementMatchResult> ExecuteResolveAll(
        UiRequest request,
        UiLocator locator,
        ElementSearchOptions options,
        AutomationSession session)
    {
        // 1. Parent-scoped lookup first
        AutomationElement? parentEl = null;
        if (request.ParentLocator != null)
        {
            var parentOptions = new ElementSearchOptions
            {
                SearchRoot = options.SearchRoot,
                Timeout = TimeSpan.Zero,
                PollInterval = TimeSpan.FromMilliseconds(100),
                ThrowIfAmbiguous = true
            };
            try
            {
                var parentMatch = ResolveOne(request, request.ParentLocator, parentOptions);
                parentEl = parentMatch.Element;
            }
            catch (Exception)
            {
                if (request.FallbackToWindowRootIfParentChildNotFound != true)
                {
                    return Array.Empty<ElementMatchResult>();
                }
            }
        }

        // 2. Determine root
        AutomationElement rootEl;
        if (parentEl != null)
        {
            rootEl = parentEl;
        }
        else
        {
            rootEl = DetermineSearchRoot(options.SearchRoot, session);
        }

        // 3. HWND direct
        if (locator.Hwnd.HasValue)
        {
            var element = session.Automation.FromHandle(new IntPtr(locator.Hwnd.Value));
            if (element != null)
            {
                return new List<ElementMatchResult> { CreateMatchResult(element, "direct-hwnd", 100, 0, "hwnd direct", session) };
            }
        }

        // 4. RuntimeId
        if (!string.IsNullOrEmpty(locator.RuntimeId))
        {
            var candidates = CollectCandidates(rootEl, locator, options, session);
            var matched = candidates.Where(c => string.Equals(UiService.SafeRuntimeIdString(c), locator.RuntimeId, StringComparison.Ordinal)).ToList();
            if (matched.Count > 0)
            {
                return matched.Select((el, idx) => CreateMatchResult(el, "runtimeid-search", 90, idx, "runtimeid match", session)).ToList();
            }
        }

        // 5. XPath only
        if (!string.IsNullOrEmpty(locator.XPath) && options.XPathOnly)
        {
            var element = UiService.FindByXPath(rootEl, session, locator.XPath);
            if (element != null)
            {
                return new List<ElementMatchResult> { CreateMatchResult(element, "xpath-only", 100, 0, "xpath only match", session) };
            }
            return Array.Empty<ElementMatchResult>();
        }

        // 6. Prefer XPath
        if (!string.IsNullOrEmpty(locator.XPath) && options.PreferXPath)
        {
            var element = UiService.FindByXPath(rootEl, session, locator.XPath);
            if (element != null)
            {
                return new List<ElementMatchResult> { CreateMatchResult(element, "xpath-preferred", 100, 0, "xpath preferred match", session) };
            }
        }

        // 7. Attribute search
        var rawCandidates = CollectCandidates(rootEl, locator, options, session);

        // Pre-filter with ctrlIndex
        if (locator.CtrlIndex.HasValue || options.CtrlIndex.HasValue)
        {
            var idx = locator.CtrlIndex ?? options.CtrlIndex!.Value;
            if (idx >= 0 && idx < rawCandidates.Count)
            {
                rawCandidates = new List<AutomationElement> { rawCandidates[idx] };
            }
            else
            {
                rawCandidates = new List<AutomationElement>();
            }
        }

        var matchedList = new List<AutomationElement>();
        foreach (var c in rawCandidates)
        {
            if (IsMatch(c, locator, options))
            {
                matchedList.Add(c);
            }
        }

        // Deduplicate
        matchedList = Deduplicate(matchedList);

        // Score them
        var results = new List<ElementMatchResult>();
        var globalMode = locator.MatchMode ?? options.MatchMode ?? "exact";
        bool isBestMode = string.Equals(globalMode, "best", StringComparison.OrdinalIgnoreCase);

        for (int i = 0; i < matchedList.Count; i++)
        {
            var el = matchedList[i];
            var (score, reason) = ElementScoring.Score(el, locator, options, session.Application.ProcessId);
            results.Add(CreateMatchResult(el, isBestMode ? "best-match" : "attribute-search", score, i, reason, session));
        }

        if (isBestMode)
        {
            results = results.OrderByDescending(r => r.Score).ToList();
        }

        // Post-filter with foundIndex
        if (locator.FoundIndex.HasValue || options.FoundIndex.HasValue)
        {
            var idx = locator.FoundIndex ?? options.FoundIndex!.Value;
            if (idx >= 0 && idx < results.Count)
            {
                return new List<ElementMatchResult> { results[idx] };
            }
            return Array.Empty<ElementMatchResult>();
        }

        // Fallback to Window Root if parent child not found
        if (results.Count == 0 && request.FallbackToWindowRootIfParentChildNotFound == true && parentEl != null)
        {
            var windowRoot = GetWindowRoot(session, false);
            var fallbackOptions = new ElementSearchOptions
            {
                SearchRoot = options.SearchRoot,
                Timeout = TimeSpan.Zero,
                PollInterval = options.PollInterval,
                ThrowIfAmbiguous = options.ThrowIfAmbiguous
            };
            var fallbackRequest = new UiRequest
            {
                Locator = locator,
                SearchRoot = options.SearchRoot
            };
            return ExecuteResolveAll(fallbackRequest, locator, fallbackOptions, session);
        }

        // Try XPath fallback if preferAttributes is true
        if (results.Count == 0 && !string.IsNullOrEmpty(locator.XPath) && options.PreferAttributes)
        {
            var element = UiService.FindByXPath(rootEl, session, locator.XPath);
            if (element != null)
            {
                return new List<ElementMatchResult> { CreateMatchResult(element, "xpath-fallback", 100, 0, "xpath fallback match", session) };
            }
        }

        return results;
    }

    private AutomationElement DetermineSearchRoot(string searchRoot, AutomationSession session)
    {
        switch (searchRoot.ToLowerInvariant())
        {
            case "active":
                return session.ActiveWindow ?? GetWindowRoot(session, false);
            case "foreground":
                return GetForegroundWindowElement(session.Automation);
            case "desktop":
                return session.Automation.GetDesktop();
            case "popup":
                return GetActivePopupRoot(session);
            case "current":
            default:
                return GetWindowRoot(session, false);
        }
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

    private List<AutomationElement> CollectCandidates(
        AutomationElement root,
        UiLocator locator,
        ElementSearchOptions options,
        AutomationSession session)
    {
        var depth = locator.Depth ?? options.Depth ?? 20;
        var topLevelOnly = locator.TopLevelOnly == true || options.TopLevelOnly;
        var treeView = options.TreeView;

        List<AutomationElement> candidates = new List<AutomationElement>();

        if (string.Equals(treeView, "raw", StringComparison.OrdinalIgnoreCase) || locator.RawView == true)
        {
            var walker = session.Automation.TreeWalkerFactory.GetRawViewWalker();
            if (topLevelOnly)
            {
                var child = walker.GetFirstChild(root);
                while (child != null)
                {
                    candidates.Add(child);
                    child = walker.GetNextSibling(child);
                }
            }
            else
            {
                candidates = UiService.FindDescendantsWithWalker(root, depth, walker);
            }
        }
        else if (string.Equals(treeView, "content", StringComparison.OrdinalIgnoreCase) || locator.ContentOnly == true || options.ContentOnly)
        {
            var walker = session.Automation.TreeWalkerFactory.GetContentViewWalker();
            if (topLevelOnly)
            {
                var child = walker.GetFirstChild(root);
                while (child != null)
                {
                    candidates.Add(child);
                    child = walker.GetNextSibling(child);
                }
            }
            else
            {
                candidates = UiService.FindDescendantsWithWalker(root, depth, walker);
            }
        }
        else
        {
            if (topLevelOnly)
            {
                candidates = root.FindAllChildren().ToList();
            }
            else if (!string.IsNullOrWhiteSpace(locator.ControlType))
            {
                try
                {
                    var normCt = NormalizeControlTypeAlias(locator.ControlType);
                    var ct = UiService.ParseControlType(normCt);
                    candidates = root.FindAllDescendants(session.Automation.ConditionFactory.ByControlType(ct)).ToList();
                }
                catch
                {
                    candidates = UiService.FindDescendantsUpToDepth(root, depth);
                }
            }
            else
            {
                candidates = UiService.FindDescendantsUpToDepth(root, depth);
            }
        }

        // Active only filter
        if (locator.ActiveOnly == true || options.ActiveOnly)
        {
            var fgHwnd = GetForegroundWindow();
            candidates = candidates.Where(c =>
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
        bool allowedOffscreen = options.IncludeOffscreen || (locator.IncludeOffscreen == true);
        if (!allowedOffscreen)
        {
            candidates = candidates.Where(c => UiService.SafeIsOffscreen(c) != true).ToList();
        }

        return candidates;
    }

    private bool IsMatch(AutomationElement element, UiLocator locator, ElementSearchOptions options)
    {
        var globalMode = locator.MatchMode ?? options.MatchMode ?? "exact";

        // 1. ControlType
        if (!string.IsNullOrWhiteSpace(locator.ControlType))
        {
            var actualCt = UiService.SafeElementControlType(element);
            var normCt = NormalizeControlTypeAlias(locator.ControlType);
            if (!string.Equals(actualCt, normCt, StringComparison.OrdinalIgnoreCase))
            {
                var localizedCt = element.Properties.LocalizedControlType.ValueOrDefault;
                if (string.IsNullOrWhiteSpace(localizedCt) || !string.Equals(localizedCt, locator.ControlType, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }
        }

        // 2. ClassName / ClassNameRegex
        if (!string.IsNullOrWhiteSpace(locator.ClassName))
        {
            var actualCn = UiService.SafeElementClassName(element);
            if (!MatchesValue(actualCn, locator.ClassName, globalMode)) return false;
        }
        if (!string.IsNullOrWhiteSpace(locator.ClassNameRegex))
        {
            var actualCn = UiService.SafeElementClassName(element);
            if (!MatchesRegex(actualCn, locator.ClassNameRegex)) return false;
        }

        // 3. Name / Title / NameRegex
        var targetName = locator.Name ?? locator.Title;
        if (!string.IsNullOrWhiteSpace(targetName))
        {
            var actualName = UiService.SafeElementName(element);
            if (!MatchesValue(actualName, targetName, globalMode)) return false;
        }
        if (!string.IsNullOrWhiteSpace(locator.NameRegex))
        {
            var actualName = UiService.SafeElementName(element);
            if (!MatchesRegex(actualName, locator.NameRegex)) return false;
        }

        // 4. AutomationId / AutoId / AutomationIdRegex / AutoIdRegex
        var targetAid = locator.AutomationId ?? locator.AutoId;
        if (!string.IsNullOrWhiteSpace(targetAid))
        {
            var actualAid = UiService.SafeElementAutomationId(element);
            if (!MatchesValue(actualAid, targetAid, "exact")) return false;
        }
        var targetAidRegex = locator.AutomationIdRegex ?? locator.AutoIdRegex;
        if (!string.IsNullOrWhiteSpace(targetAidRegex))
        {
            var actualAid = UiService.SafeElementAutomationId(element);
            if (!MatchesRegex(actualAid, targetAidRegex)) return false;
        }

        // 5. Value / ValueRegex
        if (!string.IsNullOrWhiteSpace(locator.Value))
        {
            var actualValue = UiService.SafeElementValue(element);
            if (!MatchesValue(actualValue, locator.Value, globalMode)) return false;
        }
        if (!string.IsNullOrWhiteSpace(locator.ValueRegex))
        {
            var actualValue = UiService.SafeElementValue(element);
            if (!MatchesRegex(actualValue, locator.ValueRegex)) return false;
        }

        // 6. Extra UIA fields
        if (!string.IsNullOrWhiteSpace(locator.LocalizedControlType))
        {
            var actualLct = element.Properties.LocalizedControlType.ValueOrDefault;
            if (!MatchesValue(actualLct, locator.LocalizedControlType, globalMode)) return false;
        }
        if (!string.IsNullOrWhiteSpace(locator.HelpText))
        {
            var actualHt = element.Properties.HelpText.ValueOrDefault;
            if (!MatchesValue(actualHt, locator.HelpText, globalMode)) return false;
        }
        if (!string.IsNullOrWhiteSpace(locator.AccessKey))
        {
            var actualAk = element.Properties.AccessKey.ValueOrDefault;
            if (!MatchesValue(actualAk, locator.AccessKey, globalMode)) return false;
        }
        if (!string.IsNullOrWhiteSpace(locator.AcceleratorKey))
        {
            var actualAck = element.Properties.AcceleratorKey.ValueOrDefault;
            if (!MatchesValue(actualAck, locator.AcceleratorKey, globalMode)) return false;
        }

        // 7. Native / Identity fields
        if (locator.ProcessId.HasValue)
        {
            var actualPid = UiService.SafeProcessId(element);
            if (actualPid != locator.ProcessId.Value) return false;
        }
        if (locator.ControlId.HasValue)
        {
            var actualCid = SafeControlId(element);
            if (actualCid != locator.ControlId.Value) return false;
        }

        // 8. State filters
        if (!MatchesStateFilters(element, locator, options)) return false;

        // 9. Rectangle filters
        if (!MatchesRectangleFilters(element, locator)) return false;

        return true;
    }

    private static string NormalizeControlTypeAlias(string controlType)
    {
        if (string.IsNullOrWhiteSpace(controlType)) return string.Empty;
        var norm = controlType.Trim().ToLowerInvariant();
        switch (norm)
        {
            case "dialog":
                return "Window";
            case "button":
                return "Button";
            case "text":
                return "Text";
            case "pane":
                return "Pane";
            case "list item":
            case "listitem":
                return "ListItem";
            default:
                if (norm.Length > 0)
                {
                    return char.ToUpperInvariant(norm[0]) + norm.Substring(1);
                }
                return controlType;
        }
    }

    private static string NormalizeText(string? text)
    {
        if (text == null) return string.Empty;
        var result = text.Replace("&", "");
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\s+", " ").Trim();
        return result;
    }

    private static bool MatchesValue(string? actual, string? expected, string matchMode)
    {
        if (expected == null) return true;
        if (actual == null) return false;

        var normActual = NormalizeText(actual);
        var normExpected = NormalizeText(expected);

        switch (matchMode.ToLowerInvariant())
        {
            case "contains":
                return normActual.Contains(normExpected, StringComparison.OrdinalIgnoreCase);
            case "startswith":
                return normActual.StartsWith(normExpected, StringComparison.OrdinalIgnoreCase);
            case "endswith":
                return normActual.EndsWith(normExpected, StringComparison.OrdinalIgnoreCase);
            case "regex":
                try
                {
                    return System.Text.RegularExpressions.Regex.IsMatch(actual, expected, System.Text.RegularExpressions.RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(250));
                }
                catch { return false; }
            case "exact":
            default:
                return string.Equals(normActual, normExpected, StringComparison.OrdinalIgnoreCase);
        }
    }

    private static bool MatchesRegex(string? actual, string? pattern)
    {
        if (string.IsNullOrEmpty(pattern)) return true;
        if (actual == null) return false;
        try
        {
            return System.Text.RegularExpressions.Regex.IsMatch(actual, pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(250));
        }
        catch
        {
            return false;
        }
    }

    private static bool MatchesStateFilters(AutomationElement element, UiLocator locator, ElementSearchOptions options)
    {
        if (locator.Visible.HasValue)
        {
            var offscreen = UiService.SafeIsOffscreen(element) ?? false;
            var rect = element.BoundingRectangle;
            var visible = !offscreen && !rect.IsEmpty && rect.Width > 0 && rect.Height > 0;
            if (visible != locator.Visible.Value) return false;
        }

        if (locator.Enabled.HasValue)
        {
            var enabled = UiService.SafeIsEnabled(element) ?? false;
            if (enabled != locator.Enabled.Value) return false;
        }

        if (locator.Offscreen.HasValue)
        {
            var offscreen = UiService.SafeIsOffscreen(element) ?? false;
            if (offscreen != locator.Offscreen.Value) return false;
        }

        if (locator.ContentOnly.HasValue || options.ContentOnly)
        {
            var contentOnlyVal = locator.ContentOnly ?? options.ContentOnly;
            try
            {
                var isContent = element.Properties.IsContentElement.ValueOrDefault;
                if (contentOnlyVal && !isContent) return false;
            }
            catch { return false; }
        }

        return true;
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
            var nearTolerance = locator.Tolerance ?? DefaultNearPointTolerancePixels;

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

    private List<AutomationElement> Deduplicate(List<AutomationElement> list)
    {
        var deduped = new List<AutomationElement>();
        var seenRuntimeIds = new HashSet<string>();
        var seenHwnds = new HashSet<IntPtr>();
        var seenKeys = new HashSet<string>();

        foreach (var el in list)
        {
            try
            {
                var rtId = UiService.SafeRuntimeIdString(el);
                if (!string.IsNullOrEmpty(rtId))
                {
                    if (seenRuntimeIds.Contains(rtId)) continue;
                    seenRuntimeIds.Add(rtId);
                }

                var hwnd = el.Properties.NativeWindowHandle.ValueOrDefault;
                if (hwnd != IntPtr.Zero)
                {
                    if (seenHwnds.Contains(hwnd)) continue;
                    seenHwnds.Add(hwnd);
                }

                var aid = UiService.SafeElementAutomationId(el) ?? string.Empty;
                var ct = UiService.SafeElementControlType(el) ?? string.Empty;
                var rect = el.BoundingRectangle;
                var rectKey = rect.IsEmpty ? "empty" : $"{rect.Left},{rect.Top},{rect.Width},{rect.Height}";
                var key = $"{aid}|{ct}|{rectKey}";
                if (seenKeys.Contains(key)) continue;
                seenKeys.Add(key);

                deduped.Add(el);
            }
            catch
            {
                deduped.Add(el);
            }
        }
        return deduped;
    }

    private ElementSnapshot CreateSnapshot(AutomationElement element)
    {
        if (element == null) return new ElementSnapshot();

        bool supportsInvoke = false;
        bool supportsValue = false;
        bool supportsSelectionItem = false;
        bool supportsToggle = false;
        bool supportsExpandCollapse = false;
        bool supportsScrollItem = false;

        try { supportsInvoke = element.Patterns.Invoke.IsSupported; } catch {}
        try { supportsValue = element.Patterns.Value.IsSupported; } catch {}
        try { supportsSelectionItem = element.Patterns.SelectionItem.IsSupported; } catch {}
        try { supportsToggle = element.Patterns.Toggle.IsSupported; } catch {}
        try { supportsExpandCollapse = element.Patterns.ExpandCollapse.IsSupported; } catch {}
        try { supportsScrollItem = element.Patterns.ScrollItem.IsSupported; } catch {}

        long? hwnd = null;
        try
        {
            var h = element.Properties.NativeWindowHandle.ValueOrDefault;
            if (h != IntPtr.Zero) hwnd = h.ToInt64();
        }
        catch {}

        int? pid = null;
        try { pid = UiService.SafeProcessId(element); } catch {}

        object? rectObj = null;
        try { rectObj = UiService.SafeBoundingRectangleObject(element); } catch {}

        return new ElementSnapshot
        {
            Name = UiService.SafeElementName(element) ?? string.Empty,
            AutomationId = UiService.SafeElementAutomationId(element) ?? string.Empty,
            ClassName = UiService.SafeElementClassName(element) ?? string.Empty,
            ControlType = UiService.SafeElementControlType(element) ?? string.Empty,
            LocalizedControlType = element.Properties.LocalizedControlType.ValueOrDefault ?? string.Empty,
            FrameworkId = UiService.SafeFrameworkId(element) ?? string.Empty,
            Value = UiService.SafeElementValue(element) ?? string.Empty,
            Hwnd = hwnd,
            ProcessId = pid,
            RuntimeId = UiService.SafeRuntimeIdString(element) ?? string.Empty,
            IsEnabled = UiService.SafeIsEnabled(element),
            IsOffscreen = UiService.SafeIsOffscreen(element),
            HasKeyboardFocus = element.Properties.HasKeyboardFocus.ValueOrDefault,
            Rectangle = rectObj,
            SupportsInvoke = supportsInvoke,
            SupportsValue = supportsValue,
            SupportsSelectionItem = supportsSelectionItem,
            SupportsToggle = supportsToggle,
            SupportsExpandCollapse = supportsExpandCollapse,
            SupportsScrollItem = supportsScrollItem
        };
    }

    private ElementMatchResult CreateMatchResult(
        AutomationElement element,
        string strategy,
        int score,
        int index,
        string reason,
        AutomationSession session)
    {
        return new ElementMatchResult
        {
            Element = element,
            Strategy = strategy,
            Score = score,
            Index = index,
            Reason = reason,
            Snapshot = CreateSnapshot(element)
        };
    }

    private string BuildNotFoundMessage(UiLocator locator, ElementSearchOptions options)
    {
        return $"ElementNotFound: No matching elements found for locator={UiService.DescribeLocator(locator)} under search root: {options.SearchRoot}.";
    }

    private string BuildAmbiguousMessage(UiLocator locator, IReadOnlyList<ElementMatchResult> matches)
    {
        var msg = $"ElementAmbiguous: Multiple matching elements found ({matches.Count}) for locator={UiService.DescribeLocator(locator)}.\n";
        msg += "Matches:\n";
        foreach (var m in matches)
        {
            msg += $"  - Name: {m.Snapshot.Name}, ControlType: {m.Snapshot.ControlType}, AutomationId: {m.Snapshot.AutomationId}, ClassName: {m.Snapshot.ClassName}, Score: {m.Score}, Reason: {m.Reason}\n";
        }
        return msg;
    }
}
