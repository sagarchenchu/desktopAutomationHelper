using System;
using System.Collections.Generic;
using System.Linq;
using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Services;

namespace DesktopAutomationDriver.Services.Resolution;

public static class ElementScorer
{
    public static void Score(
        ElementCandidate candidate,
        UiLocator locator,
        ElementSearchRequest searchRequest,
        int? appProcessId,
        bool isParentScoped)
    {
        var snapshot = candidate.Snapshot;
        int score = 0;

        // 1. +100 direct HWND match
        if (locator.Hwnd.HasValue && snapshot.Hwnd == locator.Hwnd.Value)
        {
            score += 100;
            candidate.MatchReasons.Add("HWND direct match (+100)");
        }

        // 2. +90 exact automationId
        var targetAid = locator.AutomationId ?? locator.AutoId;
        if (!string.IsNullOrWhiteSpace(targetAid) && string.Equals(snapshot.AutomationId, targetAid, StringComparison.OrdinalIgnoreCase))
        {
            score += 90;
            candidate.MatchReasons.Add("AutomationId exact match (+90)");
        }

        // 3. +80 exact runtimeId
        var targetRtId = locator.RuntimeId;
        var actualRtId = UiService.SafeRuntimeIdString(candidate.Element);
        if (!string.IsNullOrWhiteSpace(targetRtId) && string.Equals(actualRtId, targetRtId, StringComparison.Ordinal))
        {
            score += 80;
            candidate.MatchReasons.Add("RuntimeId exact match (+80)");
        }

        // 4. +70 exact controlType
        if (!string.IsNullOrWhiteSpace(locator.ControlType) && string.Equals(snapshot.ControlType, locator.ControlType, StringComparison.OrdinalIgnoreCase))
        {
            score += 70;
            candidate.MatchReasons.Add("ControlType exact match (+70)");
        }

        // 5. +60 exact name
        var targetName = locator.Name ?? locator.Title;
        if (!string.IsNullOrWhiteSpace(targetName) && string.Equals(snapshot.Name, targetName, StringComparison.OrdinalIgnoreCase))
        {
            score += 60;
            candidate.MatchReasons.Add("Name exact match (+60)");
        }

        // 6. +50 exact className
        if (!string.IsNullOrWhiteSpace(locator.ClassName) && string.Equals(snapshot.ClassName, locator.ClassName, StringComparison.OrdinalIgnoreCase))
        {
            score += 50;
            candidate.MatchReasons.Add("ClassName exact match (+50)");
        }

        // 7. +40 exact frameworkId
        if (!string.IsNullOrWhiteSpace(locator.FrameworkId) && string.Equals(snapshot.FrameworkId, locator.FrameworkId, StringComparison.OrdinalIgnoreCase))
        {
            score += 40;
            candidate.MatchReasons.Add("FrameworkId exact match (+40)");
        }

        // 8. +30 value/text match
        if (!string.IsNullOrWhiteSpace(locator.Value) && string.Equals(snapshot.Value, locator.Value, StringComparison.OrdinalIgnoreCase))
        {
            score += 30;
            candidate.MatchReasons.Add("Value exact match (+30)");
        }
        else if (!string.IsNullOrWhiteSpace(locator.Text) && string.Equals(snapshot.Text, locator.Text, StringComparison.OrdinalIgnoreCase))
        {
            score += 30;
            candidate.MatchReasons.Add("Text exact match (+30)");
        }

        // 9. +25 visible
        if (snapshot.IsVisible == true)
        {
            score += 25;
            candidate.MatchReasons.Add("IsVisible is true (+25)");
        }

        // 10. +20 enabled
        if (snapshot.IsEnabled == true)
        {
            score += 20;
            candidate.MatchReasons.Add("IsEnabled is true (+20)");
        }

        // 11. +20 active/top-level match
        if (searchRequest.TopLevelOnly)
        {
            score += 20;
            candidate.MatchReasons.Add("TopLevelOnly search root match (+20)");
        }

        // 12. +15 rectangle/near point match
        if (locator.NearX.HasValue || locator.NearY.HasValue || locator.Left.HasValue || locator.Top.HasValue)
        {
            score += 15;
            candidate.MatchReasons.Add("Rectangle or NearPoint filter match (+15)");
        }

        // 13. +10 same process
        if (appProcessId.HasValue && snapshot.ProcessId == appProcessId.Value)
        {
            score += 10;
            candidate.MatchReasons.Add("Same process match (+10)");
        }

        // 14. +10 parent-scoped match
        if (isParentScoped)
        {
            score += 10;
            candidate.MatchReasons.Add("Parent-scoped match (+10)");
        }

        // 15. +bestMatch score depending on fuzzy confidence
        var bestMatch = locator.BestMatch ?? searchRequest.Locator?.BestMatch;
        if (!string.IsNullOrWhiteSpace(bestMatch))
        {
            // Compute maximum fuzzy score across all possible text fields
            int bestFuzzy = new[] { snapshot.Name, snapshot.Value, snapshot.Text, snapshot.LegacyName, snapshot.LegacyValue }
                .Select(text => CalculateFuzzyScore(text, bestMatch))
                .Max();

            if (bestFuzzy > 0)
            {
                score += bestFuzzy;
                candidate.MatchReasons.Add($"BestMatch fuzzy confidence matched with score {bestFuzzy} (+{bestFuzzy})");
            }
        }

        candidate.Score = score;
    }

    public static int CalculateFuzzyScore(string? actual, string? target)
    {
        if (string.IsNullOrWhiteSpace(actual) || string.IsNullOrWhiteSpace(target)) return 0;
        actual = actual.Trim().ToLowerInvariant();
        target = target.Trim().ToLowerInvariant();
        if (actual == target) return 100;
        if (actual.Contains(target)) return 85;

        // Compute Levenshtein distance
        int lenAct = actual.Length;
        int lenTar = target.Length;
        int[,] d = new int[lenAct + 1, lenTar + 1];

        for (int i = 0; i <= lenAct; i++) d[i, 0] = i;
        for (int j = 0; j <= lenTar; j++) d[0, j] = j;

        for (int i = 1; i <= lenAct; i++)
        {
            for (int j = 1; j <= lenTar; j++)
            {
                int cost = (target[j - 1] == actual[i - 1]) ? 0 : 1;
                d[i, j] = Math.Min(
                    Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                    d[i - 1, j - 1] + cost);
            }
        }

        int distance = d[lenAct, lenTar];
        int maxLen = Math.Max(lenAct, lenTar);
        double score = (1.0 - ((double)distance / maxLen)) * 100.0;
        return (int)score;
    }
}
