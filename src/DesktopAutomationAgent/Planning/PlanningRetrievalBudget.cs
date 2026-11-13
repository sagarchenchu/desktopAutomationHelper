using System.Text.Json;

namespace DesktopAutomationAgent.Planning;

/// <summary>
/// Builds a bounded retrieval context that remains valid JSON (never mid-string truncation).
/// </summary>
public static class PlanningRetrievalBudget
{
    public static string Build(RetrievalManifest retrieval, int maxChars)
    {
        ArgumentNullException.ThrowIfNull(retrieval);
        maxChars = Math.Max(256, maxChars);

        var working = Clone(retrieval);
        working.Hits = working.Hits
            .OrderByDescending(h => h.Score)
            .ThenBy(h => h.Corpus, StringComparer.Ordinal)
            .ThenBy(h => h.Id, StringComparer.Ordinal)
            .ToList();

        while (true)
        {
            var json = JsonSerializer.Serialize(working, PlanningJson.Compact);
            using (var _ = JsonDocument.Parse(json))
            {
                if (json.Length <= maxChars)
                    return json;
            }

            if (working.Hits.Count == 0)
            {
                // Drop optional metadata fields until under budget.
                working.Exclusions = [];
                working.KnowledgeContentHash = null;
                working.ObjectRepositoryHash = null;
                working.StepRegistryHash = null;
                working.CatalogFingerprint = null;
                json = JsonSerializer.Serialize(working, PlanningJson.Compact);
                using var doc = JsonDocument.Parse(json);
                if (json.Length <= maxChars)
                    return json;

                // Last resort: keep only goal identity.
                var minimal = new RetrievalManifest
                {
                    RetrievalVersion = working.RetrievalVersion,
                    CanonicalGoal = Truncate(working.CanonicalGoal, Math.Max(32, maxChars / 4)),
                    GoalHash = working.GoalHash,
                    Hits = [],
                    Exclusions = ["prompt-budget-truncated"]
                };
                json = JsonSerializer.Serialize(minimal, PlanningJson.Compact);
                using var minDoc = JsonDocument.Parse(json);
                return json.Length <= maxChars
                    ? json
                    : JsonSerializer.Serialize(
                        new RetrievalManifest
                        {
                            RetrievalVersion = working.RetrievalVersion,
                            CanonicalGoal = "budget",
                            GoalHash = working.GoalHash,
                            Hits = [],
                            Exclusions = ["prompt-budget-truncated"]
                        },
                        PlanningJson.Compact);
            }

            // Prefer shortening the lowest-ranked snippet, else drop the hit.
            var lowest = working.Hits[^1];
            if (!string.IsNullOrEmpty(lowest.Snippet) && lowest.Snippet.Length > 24)
            {
                lowest.Snippet = Truncate(lowest.Snippet, Math.Max(12, lowest.Snippet.Length / 2));
                continue;
            }

            working.Hits.RemoveAt(working.Hits.Count - 1);
        }
    }

    private static RetrievalManifest Clone(RetrievalManifest source) =>
        new()
        {
            RetrievalVersion = source.RetrievalVersion,
            CanonicalGoal = source.CanonicalGoal,
            GoalHash = source.GoalHash,
            ApplicationId = source.ApplicationId,
            BuildFingerprint = source.BuildFingerprint,
            KnowledgeContentHash = source.KnowledgeContentHash,
            ObjectRepositoryHash = source.ObjectRepositoryHash,
            StepRegistryHash = source.StepRegistryHash,
            CatalogFingerprint = source.CatalogFingerprint,
            Hits = source.Hits.Select(h => new RetrievalHitRecord
            {
                Corpus = h.Corpus,
                Kind = h.Kind,
                Id = h.Id,
                Score = h.Score,
                ContentHash = h.ContentHash,
                Snippet = h.Snippet
            }).ToList(),
            Exclusions = source.Exclusions.ToList(),
            EmbeddingsUsed = source.EmbeddingsUsed
        };

    private static string Truncate(string text, int max) =>
        text.Length <= max ? text : text[..max];
}
