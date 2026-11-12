using System.Text;
using Gherkin;
using Gherkin.Ast;

namespace DesktopAutomationAgent.Compilation;

public sealed class BddParseResult
{
    public bool IsSuccess => Errors.Count == 0 && Document is not null;

    public BddIrDocument? Document { get; init; }

    public List<string> Errors { get; init; } = [];
}

/// <summary>
/// Parses Phase 4D BDD v1 using the official Cucumber <c>Gherkin</c> package (MIT),
/// then rejects unsupported constructs with precise errors. Does not claim full Gherkin.
/// </summary>
public sealed class GherkinBddParser
{
    public BddParseResult ParseTestCase(JiraTestCaseDocument testCase)
    {
        ArgumentNullException.ThrowIfNull(testCase);

        if (!string.IsNullOrWhiteSpace(testCase.Feature))
            return ParseFeatureText(testCase.Feature, testCase.Tags);

        return ParseStructured(testCase);
    }

    public BddParseResult ParseFeatureText(string featureText, IReadOnlyList<string>? documentTags = null)
    {
        GherkinDocument gherkinDocument;
        try
        {
            var parser = new Parser();
            gherkinDocument = parser.Parse(new StringReader(NormalizeNewlines(featureText)));
        }
        catch (ParserException ex)
        {
            return new BddParseResult
            {
                Errors = [$"unsupported or invalid Gherkin: {ex.Message}"]
            };
        }
        catch (Exception ex) when (ex is InvalidOperationException or ArgumentException)
        {
            return new BddParseResult
            {
                Errors = [$"Gherkin parse failed: {ex.Message}"]
            };
        }

        if (gherkinDocument.Feature is null)
        {
            return new BddParseResult
            {
                Errors = ["Gherkin document has no Feature."]
            };
        }

        return ToIr(gherkinDocument.Feature, documentTags);
    }

    private BddParseResult ParseStructured(JiraTestCaseDocument testCase)
    {
        if (testCase.Scenarios is null || testCase.Scenarios.Count == 0)
        {
            return new BddParseResult
            {
                Errors = ["structured scenarios are required when feature text is omitted."]
            };
        }

        var sb = new StringBuilder();
        sb.AppendLine("Feature: " + SanitizeFeatureName(testCase.Summary));
        if (testCase.Background is { Steps.Count: > 0 } background)
        {
            sb.AppendLine("  Background:" + (string.IsNullOrWhiteSpace(background.Name) ? string.Empty : " " + background.Name));
            foreach (var step in background.Steps)
                sb.AppendLine($"    {step.Keyword} {step.Text}");
        }

        foreach (var scenario in testCase.Scenarios)
        {
            if (scenario.Tags is { Count: > 0 })
            {
                foreach (var tag in scenario.Tags.OrderBy(t => t, StringComparer.Ordinal))
                    sb.AppendLine("  @" + tag.TrimStart('@'));
            }

            sb.AppendLine("  Scenario: " + scenario.Name);
            foreach (var step in scenario.Steps)
                sb.AppendLine($"    {step.Keyword} {step.Text}");
        }

        return ParseFeatureText(sb.ToString(), testCase.Tags);
    }

    private static BddParseResult ToIr(Feature feature, IReadOnlyList<string>? documentTags)
    {
        var errors = new List<string>();
        var scenarios = new List<BddIrScenario>();
        IReadOnlyList<BddIrStep> backgroundSteps = Array.Empty<BddIrStep>();

        foreach (var child in feature.Children)
        {
            switch (child)
            {
                case Background background:
                    backgroundSteps = NormalizeSteps(background.Steps, "Background", errors);
                    break;
                case Scenario scenario:
                    if (scenario.Examples is { } examples && examples.Any())
                    {
                        errors.Add(
                            $"unsupported-feature: Scenario Outline/Examples is not supported in BDD v1 " +
                            $"(scenario '{scenario.Name}' at line {scenario.Location.Line}).");
                        continue;
                    }

                    var steps = new List<BddIrStep>();
                    steps.AddRange(backgroundSteps);
                    steps.AddRange(NormalizeSteps(scenario.Steps, $"Scenario '{scenario.Name}'", errors));
                    // Re-number after merge for stable ordinals.
                    for (var i = 0; i < steps.Count; i++)
                    {
                        steps[i] = new BddIrStep
                        {
                            RawKeyword = steps[i].RawKeyword,
                            Keyword = steps[i].Keyword,
                            Text = steps[i].Text,
                            Line = steps[i].Line,
                            Ordinal = i
                        };
                    }

                    scenarios.Add(new BddIrScenario
                    {
                        Name = scenario.Name,
                        Tags = scenario.Tags
                            .Select(t => t.Name.TrimStart('@'))
                            .OrderBy(t => t, StringComparer.Ordinal)
                            .ToArray(),
                        Steps = steps
                    });
                    break;
                case Rule rule:
                    errors.Add(
                        $"unsupported-feature: Rule is not supported in BDD v1 " +
                        $"(rule '{rule.Name}' at line {rule.Location.Line}).");
                    break;
                default:
                    errors.Add(
                        $"unsupported-feature: unexpected feature child '{child.GetType().Name}'.");
                    break;
            }
        }

        if (scenarios.Count == 0 && errors.Count == 0)
            errors.Add("Feature must contain at least one Scenario.");

        if (errors.Count > 0)
            return new BddParseResult { Errors = errors };

        var tags = (documentTags ?? Array.Empty<string>())
            .Concat(feature.Tags.Select(t => t.Name.TrimStart('@')))
            .Distinct(StringComparer.Ordinal)
            .OrderBy(t => t, StringComparer.Ordinal)
            .ToArray();

        return new BddParseResult
        {
            Document = new BddIrDocument
            {
                FeatureName = feature.Name,
                Scenarios = scenarios
                    .OrderBy(s => s.Name, StringComparer.Ordinal)
                    .ThenBy(s => s.Steps.Count)
                    .ToArray(),
                Tags = tags
            }
        };
    }

    private static List<BddIrStep> NormalizeSteps(
        IEnumerable<Step> steps,
        string scope,
        List<string> errors)
    {
        var result = new List<BddIrStep>();
        string? currentPrimary = null;
        var ordinal = 0;

        foreach (var step in steps)
        {
            if (step.Argument is DataTable)
            {
                errors.Add(
                    $"unsupported-feature: data tables are not supported in BDD v1 " +
                    $"({scope} step at line {step.Location.Line}).");
                continue;
            }

            if (step.Argument is DocString)
            {
                errors.Add(
                    $"unsupported-feature: doc strings are not supported in BDD v1 " +
                    $"({scope} step at line {step.Location.Line}).");
                continue;
            }

            if (step.Argument is not null)
            {
                errors.Add(
                    $"unsupported-feature: step argument '{step.Argument.GetType().Name}' is not supported " +
                    $"({scope} step at line {step.Location.Line}).");
                continue;
            }

            var raw = step.Keyword.Trim();
            var normalized = NormalizeKeyword(raw, ref currentPrimary, scope, step.Location.Line, errors);
            if (normalized is null)
                continue;

            result.Add(new BddIrStep
            {
                RawKeyword = raw,
                Keyword = normalized,
                Text = step.Text.Trim(),
                Line = step.Location.Line,
                Ordinal = ordinal++
            });
        }

        return result;
    }

    private static string? NormalizeKeyword(
        string raw,
        ref string? currentPrimary,
        string scope,
        int line,
        List<string> errors)
    {
        if (raw.Equals("Given", StringComparison.OrdinalIgnoreCase)
            || raw.Equals("When", StringComparison.OrdinalIgnoreCase)
            || raw.Equals("Then", StringComparison.OrdinalIgnoreCase))
        {
            var primary = char.ToUpperInvariant(raw[0]) + raw[1..].ToLowerInvariant();
            currentPrimary = primary;
            return primary;
        }

        if (raw.Equals("And", StringComparison.OrdinalIgnoreCase)
            || raw.Equals("But", StringComparison.OrdinalIgnoreCase))
        {
            if (currentPrimary is null)
            {
                errors.Add(
                    $"{scope}: '{raw}' at line {line} has no preceding Given/When/Then to normalize against.");
                return null;
            }

            return currentPrimary;
        }

        errors.Add($"{scope}: unsupported step keyword '{raw}' at line {line}.");
        return null;
    }

    private static string NormalizeNewlines(string text) =>
        text.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n');

    private static string SanitizeFeatureName(string summary)
    {
        var cleaned = summary.Replace('\r', ' ').Replace('\n', ' ').Trim();
        return string.IsNullOrWhiteSpace(cleaned) ? "Exported" : cleaned;
    }
}
