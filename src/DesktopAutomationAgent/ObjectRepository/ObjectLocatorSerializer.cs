using System.Text.Json;
using System.Text.Json.Serialization;

namespace DesktopAutomationAgent.ObjectRepository;

public static class ObjectLocatorSerializer
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    public static JsonElement ToJsonElement(ObjectLocator locator) =>
        JsonSerializer.SerializeToElement(locator, JsonOptions);

    public static Dictionary<string, JsonElement> ToArgumentDictionary(ObjectLocator locator)
    {
        var element = ToJsonElement(locator);
        if (element.ValueKind != JsonValueKind.Object)
            return new Dictionary<string, JsonElement>(StringComparer.Ordinal);

        var result = new Dictionary<string, JsonElement>(StringComparer.Ordinal);
        foreach (var property in element.EnumerateObject())
            result[property.Name] = property.Value.Clone();

        return result;
    }
}
