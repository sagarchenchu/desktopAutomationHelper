using System.Text.Json;
using System.Text.RegularExpressions;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Plans;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Execution;

public sealed class AssertionEvaluator
{
    private readonly RunnerOptions _runnerOptions;

    public AssertionEvaluator(IOptions<AgentOptions> options)
    {
        _runnerOptions = options.Value.Runner;
    }

    public IReadOnlyList<AssertionResult> Evaluate(
        JsonElement? responseValue,
        IReadOnlyList<PlanAssertion>? assertions)
    {
        if (assertions is null || assertions.Count == 0)
            return Array.Empty<AssertionResult>();

        var root = responseValue ?? default;
        var results = new List<AssertionResult>(assertions.Count);

        foreach (var assertion in assertions)
        {
            results.Add(EvaluateOne(root, assertion));
        }

        return results;
    }

    private AssertionResult EvaluateOne(JsonElement root, PlanAssertion assertion)
    {
        var path = assertion.Path ?? string.Empty;
        var op = assertion.Operator ?? string.Empty;

        if (!TryResolvePointer(root, path, out var actual, out var pointerError))
        {
            return new AssertionResult
            {
                Path = path,
                Operator = op,
                Passed = false,
                Message = pointerError,
                Expected = assertion.Expected,
                Actual = null
            };
        }

        var passed = op switch
        {
            "equals" => CompareEquals(actual, assertion.Expected, assertion.IgnoreCase),
            "notEquals" => !CompareEquals(actual, assertion.Expected, assertion.IgnoreCase),
            "contains" => CompareContains(actual, assertion.Expected, assertion.IgnoreCase),
            "matchesRegex" => CompareRegex(actual, assertion.Expected),
            "isTrue" => actual.ValueKind == JsonValueKind.True,
            "isFalse" => actual.ValueKind == JsonValueKind.False,
            "isNull" => actual.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined,
            "isNotNull" => actual.ValueKind is not (JsonValueKind.Null or JsonValueKind.Undefined),
            _ => false
        };

        return new AssertionResult
        {
            Path = path,
            Operator = op,
            Passed = passed,
            Message = passed ? null : BuildFailureMessage(op, assertion.Expected, actual),
            Expected = assertion.Expected,
            Actual = actual.ValueKind is JsonValueKind.Undefined ? null : actual
        };
    }

    private static string BuildFailureMessage(string op, JsonElement? expected, JsonElement actual) =>
        op switch
        {
            "matchesRegex" => $"Value did not match expected regex.",
            "contains" => "Value does not contain the expected item.",
            "isTrue" => "Expected true.",
            "isFalse" => "Expected false.",
            "isNull" => "Expected null.",
            "isNotNull" => "Expected a non-null value.",
            _ => "Assertion comparison failed."
        };

    private bool CompareEquals(JsonElement actual, JsonElement? expected, bool ignoreCase)
    {
        if (expected is null)
            return false;

        if (actual.ValueKind == JsonValueKind.Number && expected.Value.ValueKind == JsonValueKind.Number)
        {
            if (actual.TryGetDecimal(out var left) && expected.Value.TryGetDecimal(out var right))
                return left == right;
        }

        if (actual.ValueKind == JsonValueKind.String && expected.Value.ValueKind == JsonValueKind.String)
        {
            var comparison = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            return string.Equals(actual.GetString(), expected.Value.GetString(), comparison);
        }

        return JsonElementDeepEquals(actual, expected.Value);
    }

    private static bool CompareContains(JsonElement actual, JsonElement? expected, bool ignoreCase)
    {
        if (expected is null)
            return false;

        return actual.ValueKind switch
        {
            JsonValueKind.String when expected.Value.ValueKind == JsonValueKind.String =>
                (actual.GetString() ?? string.Empty).Contains(
                    expected.Value.GetString() ?? string.Empty,
                    ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal),
            JsonValueKind.Array =>
                actual.EnumerateArray().Any(item => JsonElementDeepEquals(item, expected.Value)),
            _ => false
        };
    }

    private bool CompareRegex(JsonElement actual, JsonElement? expected)
    {
        if (expected is null || expected.Value.ValueKind != JsonValueKind.String)
            return false;

        if (actual.ValueKind != JsonValueKind.String)
            return false;

        var pattern = expected.Value.GetString();
        if (string.IsNullOrEmpty(pattern))
            return false;

        try
        {
            var regex = new Regex(
                pattern,
                RegexOptions.CultureInvariant,
                TimeSpan.FromMilliseconds(_runnerOptions.RegexTimeoutMilliseconds));
            return regex.IsMatch(actual.GetString() ?? string.Empty);
        }
        catch (RegexMatchTimeoutException)
        {
            return false;
        }
        catch (ArgumentException)
        {
            // Invalid patterns should be rejected offline; fail closed at runtime.
            return false;
        }
    }

    internal static bool TryResolvePointer(JsonElement root, string pointer, out JsonElement value, out string? error)
    {
        error = null;
        value = default;

        if (pointer.Length == 0)
        {
            value = root;
            return true;
        }

        if (pointer[0] != '/')
        {
            error = $"Invalid JSON pointer '{pointer}'.";
            return false;
        }

        var current = root;
        var tokens = pointer[1..].Split('/');
        foreach (var rawToken in tokens)
        {
            var token = rawToken.Replace("~1", "/", StringComparison.Ordinal)
                .Replace("~0", "~", StringComparison.Ordinal);

            switch (current.ValueKind)
            {
                case JsonValueKind.Object:
                {
                    var found = false;
                    foreach (var property in current.EnumerateObject())
                    {
                        if (!string.Equals(property.Name, token, StringComparison.Ordinal))
                            continue;

                        current = property.Value;
                        found = true;
                        break;
                    }

                    if (!found)
                    {
                        error = $"JSON pointer '{pointer}' did not resolve: property '{token}' was not found.";
                        return false;
                    }

                    break;
                }
                case JsonValueKind.Array when int.TryParse(token, out var index):
                {
                    var array = current.EnumerateArray().ToArray();
                    if (index < 0 || index >= array.Length)
                    {
                        error = $"JSON pointer '{pointer}' index {index} is out of range.";
                        return false;
                    }

                    current = array[index];
                    break;
                }
                default:
                    error = $"JSON pointer '{pointer}' cannot traverse token '{token}'.";
                    return false;
            }
        }

        value = current;
        return true;
    }

    internal static bool JsonElementDeepEquals(JsonElement left, JsonElement right)
    {
        if (left.ValueKind != right.ValueKind)
            return false;

        return left.ValueKind switch
        {
            JsonValueKind.Object => ObjectEquals(left, right),
            JsonValueKind.Array => ArrayEquals(left, right),
            JsonValueKind.String => string.Equals(left.GetString(), right.GetString(), StringComparison.Ordinal),
            JsonValueKind.Number => left.GetRawText() == right.GetRawText(),
            JsonValueKind.True or JsonValueKind.False or JsonValueKind.Null => true,
            _ => left.GetRawText() == right.GetRawText()
        };
    }

    private static bool ObjectEquals(JsonElement left, JsonElement right)
    {
        var leftProps = left.EnumerateObject()
            .ToDictionary(p => p.Name, p => p.Value, StringComparer.Ordinal);
        var rightProps = right.EnumerateObject()
            .ToDictionary(p => p.Name, p => p.Value, StringComparer.Ordinal);

        if (leftProps.Count != rightProps.Count)
            return false;

        foreach (var pair in leftProps)
        {
            if (!rightProps.TryGetValue(pair.Key, out var other))
                return false;

            if (!JsonElementDeepEquals(pair.Value, other))
                return false;
        }

        return true;
    }

    private static bool ArrayEquals(JsonElement left, JsonElement right)
    {
        var leftItems = left.EnumerateArray().ToArray();
        var rightItems = right.EnumerateArray().ToArray();
        if (leftItems.Length != rightItems.Length)
            return false;

        for (var i = 0; i < leftItems.Length; i++)
        {
            if (!JsonElementDeepEquals(leftItems[i], rightItems[i]))
                return false;
        }

        return true;
    }
}
