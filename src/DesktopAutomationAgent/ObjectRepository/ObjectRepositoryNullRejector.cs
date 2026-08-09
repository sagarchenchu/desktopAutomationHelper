using System.Text.Json;

namespace DesktopAutomationAgent.ObjectRepository;

/// <summary>
/// Rejects explicit JSON null values. Checked-in schemas use non-nullable types and
/// <c>additionalProperties: false</c>, so present-but-null properties are invalid.
/// </summary>
internal static class ObjectRepositoryNullRejector
{
    public static IReadOnlyList<string> Detect(ReadOnlySpan<byte> utf8Json, string location)
    {
        try
        {
            using var document = JsonDocument.Parse(utf8Json.ToArray());
            var errors = new List<string>();
            Walk(document.RootElement, location, errors);
            return errors;
        }
        catch (JsonException ex)
        {
            return [$"{location}: invalid JSON ({ex.Message})."];
        }
    }

    private static void Walk(JsonElement element, string location, List<string> errors)
    {
        switch (element.ValueKind)
        {
            case JsonValueKind.Object:
                foreach (var property in element.EnumerateObject())
                {
                    var childLocation = $"{location}.{property.Name}";
                    if (property.Value.ValueKind == JsonValueKind.Null)
                    {
                        errors.Add($"{childLocation}: value must not be null.");
                        continue;
                    }

                    Walk(property.Value, childLocation, errors);
                }

                break;

            case JsonValueKind.Array:
                var index = 0;
                foreach (var item in element.EnumerateArray())
                {
                    var childLocation = $"{location}[{index}]";
                    if (item.ValueKind == JsonValueKind.Null)
                    {
                        errors.Add($"{childLocation}: value must not be null.");
                    }
                    else
                    {
                        Walk(item, childLocation, errors);
                    }

                    index++;
                }

                break;
        }
    }
}
