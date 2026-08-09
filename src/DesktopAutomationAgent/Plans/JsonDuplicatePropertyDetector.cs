using System.Text;
using System.Text.Json;

namespace DesktopAutomationAgent.Plans;

public static class JsonDuplicatePropertyDetector
{
    public static IReadOnlyList<string> DetectDuplicates(ReadOnlySpan<byte> utf8Json)
    {
        var errors = new List<string>();
        var reader = new Utf8JsonReader(utf8Json, new JsonReaderOptions
        {
            CommentHandling = JsonCommentHandling.Skip,
            AllowTrailingCommas = true
        });

        if (!reader.Read())
            return errors;

        if (reader.TokenType is JsonTokenType.StartObject or JsonTokenType.StartArray)
            VisitValue(ref reader, "$", errors);

        return errors;
    }

    public static IReadOnlyList<string> DetectDuplicates(string json) =>
        DetectDuplicates(Encoding.UTF8.GetBytes(json));

    private static void VisitValue(ref Utf8JsonReader reader, string path, List<string> errors)
    {
        switch (reader.TokenType)
        {
            case JsonTokenType.StartObject:
                VisitObject(ref reader, path, errors);
                break;
            case JsonTokenType.StartArray:
                VisitArray(ref reader, path, errors);
                break;
        }
    }

    private static void VisitObject(ref Utf8JsonReader reader, string path, List<string> errors)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var index = 0;

        while (reader.Read())
        {
            if (reader.TokenType == JsonTokenType.EndObject)
                return;

            if (reader.TokenType != JsonTokenType.PropertyName)
                continue;

            var propertyName = reader.GetString() ?? string.Empty;
            if (!seen.Add(propertyName))
            {
                errors.Add($"{path}: duplicate property '{propertyName}'.");
            }

            var childPath = $"{path}.{propertyName}";
            if (!reader.Read())
                return;

            VisitValue(ref reader, childPath, errors);
            index++;
        }
    }

    private static void VisitArray(ref Utf8JsonReader reader, string path, List<string> errors)
    {
        var index = 0;

        while (reader.Read())
        {
            if (reader.TokenType == JsonTokenType.EndArray)
                return;

            VisitValue(ref reader, $"{path}[{index}]", errors);
            index++;
        }
    }
}
