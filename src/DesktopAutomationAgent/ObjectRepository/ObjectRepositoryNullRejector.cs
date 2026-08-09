using System.Text.Json;

namespace DesktopAutomationAgent.ObjectRepository;

/// <summary>
/// Rejects explicit JSON null values where the checked-in schemas use non-nullable types.
/// Null is permitted inside open schema bags: values under <c>source.metadata</c> and
/// properties of <c>unresolved[*]</c> objects (both allow arbitrary JSON).
/// Parsing matches <see cref="ObjectRepositoryReader"/> (comments + trailing commas).
/// </summary>
internal static class ObjectRepositoryNullRejector
{
    private static readonly JsonDocumentOptions DocumentOptions = new()
    {
        CommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true
    };

    public static IReadOnlyList<string> Detect(ReadOnlySpan<byte> utf8Json, string location)
    {
        try
        {
            using var document = JsonDocument.Parse(utf8Json.ToArray(), DocumentOptions);
            var errors = new List<string>();
            Walk(document.RootElement, location, NullPolicy.Strict, errors);
            return errors;
        }
        catch (JsonException ex)
        {
            return [$"{location}: invalid JSON ({ex.Message})."];
        }
    }

    private static void Walk(
        JsonElement element,
        string location,
        NullPolicy policy,
        List<string> errors)
    {
        switch (element.ValueKind)
        {
            case JsonValueKind.Object:
                foreach (var property in element.EnumerateObject())
                {
                    var childLocation = $"{location}.{property.Name}";

                    // Null permission is inherited from the parent bag, so containers such as
                    // source.metadata themselves remain non-nullable (schema type: object).
                    if (property.Value.ValueKind == JsonValueKind.Null)
                    {
                        if (policy != NullPolicy.AllowNestedNulls)
                            errors.Add($"{childLocation}: value must not be null.");
                        continue;
                    }

                    Walk(property.Value, childLocation, NextPolicy(policy, property.Name), errors);
                }

                break;

            case JsonValueKind.Array:
                var index = 0;
                foreach (var item in element.EnumerateArray())
                {
                    var childLocation = $"{location}[{index}]";
                    if (item.ValueKind == JsonValueKind.Null)
                    {
                        if (policy != NullPolicy.AllowNestedNulls)
                            errors.Add($"{childLocation}: value must not be null.");
                    }
                    else
                    {
                        var itemPolicy = policy == NullPolicy.UnresolvedArray
                            ? NullPolicy.AllowNestedNulls
                            : policy;
                        Walk(item, childLocation, itemPolicy, errors);
                    }

                    index++;
                }

                break;
        }
    }

    private static NullPolicy NextPolicy(NullPolicy current, string propertyName)
    {
        if (current == NullPolicy.AllowNestedNulls)
            return NullPolicy.AllowNestedNulls;

        if (current == NullPolicy.Strict
            && propertyName.Equals("unresolved", StringComparison.Ordinal))
        {
            return NullPolicy.UnresolvedArray;
        }

        if (current == NullPolicy.InSource
            && propertyName.Equals("metadata", StringComparison.Ordinal))
        {
            return NullPolicy.AllowNestedNulls;
        }

        if (propertyName.Equals("source", StringComparison.Ordinal))
            return NullPolicy.InSource;

        return current;
    }

    private enum NullPolicy
    {
        Strict,
        InSource,
        UnresolvedArray,
        AllowNestedNulls
    }
}
