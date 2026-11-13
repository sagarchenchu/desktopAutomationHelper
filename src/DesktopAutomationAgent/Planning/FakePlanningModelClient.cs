using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using DesktopAutomationAgent.Driver;

namespace DesktopAutomationAgent.Planning;

/// <summary>
/// Deterministic fake model for CI and local tests. Never contacts a provider or the driver.
/// Scenario selection is derived from the canonical goal / untrusted context keywords.
/// </summary>
public sealed class FakePlanningModelClient : IPlanningModelClient
{
    public string ProviderName => "fake";

    public Task<PlanningModelResponse> ProposeAsync(
        PlanningModelRequest request,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);
        cancellationToken.ThrowIfCancellationRequested();

        var goal = request.CanonicalGoal ?? string.Empty;
        var untrusted = request.UntrustedContext ?? string.Empty;
        var combined = goal + "\n" + untrusted;
        var hits = ParseHitIds(request.RetrievalContextJson);

        string raw;
        if (ContainsAny(combined, "ignore previous instructions", "system override", "jailbreak", "bypass policy"))
        {
            // Model might try to comply with injection; validator/guard must reject.
            raw = BuildProposalJson(
                goal,
                risk: "low",
                unresolved: [],
                steps:
                [
                    Step("injected-shell", "listwindows", """{"includeDesktopDescendants":false}""", mappingId: null)
                ],
                objectRefs: [],
                knowledgeRefs: [],
                extraProperty: """ ,"systemOverride":"enabled" """);
        }
        else if (ContainsAny(goal, "delete all records", "uninstall the application", "wipe the database"))
        {
            raw = BuildProposalJson(
                goal,
                risk: "prohibited",
                unresolved: [],
                steps:
                [
                    Step("destroy", "clickuia", """{"locator":{"automationId":"DeleteAll"}}""", mappingId: null, rationale: "delete everything")
                ],
                objectRefs: [],
                knowledgeRefs: []);
        }
        else if (ContainsAny(goal, "ambiguous", "maybe click", "or perhaps"))
        {
            raw = BuildProposalJson(
                goal,
                risk: "medium",
                unresolved: ["Which window should be targeted?", "Is the control named Save or Submit?"],
                steps:
                [
                    Step("unsure", "listwindows", """{"includeDesktopDescendants":false}""")
                ],
                objectRefs: [],
                knowledgeRefs: []);
        }
        else if (ContainsAny(goal, "missing object", "unknown button"))
        {
            raw = BuildProposalJson(
                goal,
                risk: "low",
                unresolved: ["Required object ref is not in the approved repository."],
                steps:
                [
                    Step("click-missing", "clickuia", """{"$objectRef":"page.doesNotExist"}""")
                ],
                objectRefs: [new { @ref = "page.doesNotExist", role = "target" }],
                knowledgeRefs: []);
        }
        else if (ContainsAny(goal, "missing transition", "no path to"))
        {
            raw = BuildProposalJson(
                goal,
                risk: "medium",
                unresolved: ["No approved transition reaches the requested state."],
                steps:
                [
                    Step("stuck", "listwindows", """{"includeDesktopDescendants":false}""")
                ],
                objectRefs: [],
                knowledgeRefs: [new { kind = "transition", id = "tr_missing", role = "navigation" }]);
        }
        else if (ContainsAny(goal, "stale build", "wrong fingerprint"))
        {
            raw = BuildProposalJson(
                goal,
                risk: "low",
                unresolved: [],
                steps:
                [
                    Step("list", "listwindows", """{"includeDesktopDescendants":false}""", mappingId: "list-visible-windows")
                ],
                objectRefs: [],
                knowledgeRefs: hits.Knowledge.Take(1).Select(id => new { kind = "state", id, role = "navigation" }).ToArray());
        }
        else if (ContainsAny(goal, "incorrect op", "notanoperation", "fakeop"))
        {
            raw = BuildProposalJson(
                goal,
                risk: "low",
                unresolved: [],
                steps:
                [
                    Step("bad-op", "notanoperation", """{}""")
                ],
                objectRefs: [],
                knowledgeRefs: []);
        }
        else if (ContainsAny(goal, "hallucinated ref", "invented knowledge"))
        {
            raw = BuildProposalJson(
                goal,
                risk: "low",
                unresolved: [],
                steps:
                [
                    Step("hallucinated", "listwindows", """{"includeDesktopDescendants":false}""")
                ],
                objectRefs: [new { @ref = "hallucinated.page.element", role = "target" }],
                knowledgeRefs: [new { kind = "state", id = "st_hallucinated_not_real", role = "navigation" }]);
        }
        else if (ContainsAny(goal, "enter password", "type credentials", "submit payment"))
        {
            raw = BuildProposalJson(
                goal,
                risk: "high",
                unresolved: [],
                steps:
                [
                    Step("type-secret", "typeuia", """{"locator":{"automationId":"Password"},"value":"secret"}""", sensitive: true)
                ],
                objectRefs: [],
                knowledgeRefs: []);
        }
        else if (ContainsAny(goal, "multi-window", "second window", "both windows"))
        {
            var knowledge = hits.Knowledge.Count >= 2
                ? hits.Knowledge.Take(2).Select(id => new { kind = "window", id, role = "navigation" }).ToArray()
                : new[]
                {
                    new { kind = "window", id = hits.Knowledge.FirstOrDefault() ?? "win_main", role = "navigation" },
                    new { kind = "window", id = "win_dialog", role = "navigation" }
                };

            raw = BuildProposalJson(
                goal,
                risk: "low",
                unresolved: [],
                steps:
                [
                    Step("list-main", "listwindows", """{"includeDesktopDescendants":false}""", mappingId: "list-visible-windows"),
                    Step("list-again", "listwindows", """{"includeDesktopDescendants":false}""", mappingId: "list-visible-windows", capture: true)
                ],
                objectRefs: [],
                knowledgeRefs: knowledge,
                expectedState: "multi-window-ready",
                postcondition: "both windows observed");
        }
        else
        {
            // Straightforward / default: map to approved listwindows path.
            var mapping = hits.Mappings.FirstOrDefault() ?? "list-visible-windows";
            var knowledgeRefs = hits.Knowledge
                .Take(2)
                .Select(id => new { kind = "state", id, role = "navigation" })
                .ToArray<object>();

            raw = BuildProposalJson(
                goal,
                risk: "low",
                unresolved: [],
                steps:
                [
                    Step("list-windows", "listwindows", """{"includeDesktopDescendants":false}""", mappingId: mapping, capture: true)
                ],
                objectRefs: hits.Objects.Take(1).Select(r => new { @ref = r, role = "target" }).ToArray(),
                knowledgeRefs: knowledgeRefs,
                expectedState: hits.Knowledge.FirstOrDefault(),
                postcondition: "window list captured");
        }

        if (Encoding.UTF8.GetByteCount(raw) > request.MaxOutputBytes)
        {
            raw = raw[..Math.Min(raw.Length, Math.Max(16, request.MaxOutputBytes / 2))];
            return Task.FromResult(new PlanningModelResponse
            {
                RawJson = SecretRedactor.Redact(raw),
                Provider = ProviderName,
                ModelId = "fake-deterministic-v1",
                Truncated = true
            });
        }

        return Task.FromResult(new PlanningModelResponse
        {
            RawJson = raw,
            Provider = ProviderName,
            ModelId = "fake-deterministic-v1",
            Truncated = false
        });
    }

    private static bool ContainsAny(string text, params string[] needles) =>
        needles.Any(n => text.Contains(n, StringComparison.OrdinalIgnoreCase));

    private static object Step(
        string id,
        string operation,
        string argumentsJson,
        string? mappingId = null,
        string? rationale = null,
        bool sensitive = false,
        bool capture = false)
    {
        using var argsDoc = JsonDocument.Parse(argumentsJson);
        return new Dictionary<string, object?>
        {
            ["id"] = id,
            ["operation"] = operation,
            ["arguments"] = argsDoc.RootElement.Clone(),
            ["mappingId"] = mappingId,
            ["rationale"] = rationale,
            ["sensitive"] = sensitive,
            ["captureResponse"] = capture ? true : null
        };
    }

    private static string BuildProposalJson(
        string goal,
        string risk,
        IEnumerable<string> unresolved,
        IEnumerable<object> steps,
        IEnumerable<object> objectRefs,
        IEnumerable<object> knowledgeRefs,
        string? expectedState = null,
        string? postcondition = null,
        string? extraProperty = null)
    {
        var proposalId = "prop_" + Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(goal)))[..12].ToLowerInvariant();
        var payload = new Dictionary<string, object?>
        {
            ["schemaVersion"] = PlanningVersions.ProposalSchemaVersion,
            ["proposalId"] = proposalId,
            ["goal"] = goal,
            ["assumptions"] = Array.Empty<string>(),
            ["unresolvedQuestions"] = unresolved.ToArray(),
            ["steps"] = steps.ToArray(),
            ["objectRefs"] = objectRefs.ToArray(),
            ["knowledgeRefs"] = knowledgeRefs.ToArray(),
            ["expectedState"] = expectedState,
            ["postcondition"] = postcondition,
            ["risk"] = risk,
            ["confidence"] = 0.5
        };

        var json = JsonSerializer.Serialize(payload, PlanningJson.Compact);
        if (!string.IsNullOrWhiteSpace(extraProperty))
        {
            // Inject unknown property for injection/malformed scenarios (before closing brace).
            json = json.TrimEnd();
            if (json.EndsWith('}'))
                json = json[..^1] + extraProperty + "}";
        }

        return json;
    }

    private static HitBag ParseHitIds(string retrievalJson)
    {
        var result = new HitBag();
        if (string.IsNullOrWhiteSpace(retrievalJson))
            return result;

        try
        {
            using var doc = JsonDocument.Parse(retrievalJson);
            if (!doc.RootElement.TryGetProperty("hits", out var hits) || hits.ValueKind != JsonValueKind.Array)
                return result;

            foreach (var hit in hits.EnumerateArray())
            {
                var corpus = hit.TryGetProperty("corpus", out var c) ? c.GetString() ?? "" : "";
                var id = hit.TryGetProperty("id", out var i) ? i.GetString() ?? "" : "";
                if (string.IsNullOrWhiteSpace(id))
                    continue;

                switch (corpus)
                {
                    case "knowledge":
                        result.Knowledge.Add(id);
                        break;
                    case "object-repository":
                        result.Objects.Add(id);
                        break;
                    case "step-registry":
                        result.Mappings.Add(id);
                        break;
                }
            }
        }
        catch (JsonException)
        {
            // Fake client stays deterministic even on bad retrieval JSON.
        }

        return result;
    }

    private sealed class HitBag
    {
        public List<string> Knowledge { get; } = [];
        public List<string> Objects { get; } = [];
        public List<string> Mappings { get; } = [];
    }
}
