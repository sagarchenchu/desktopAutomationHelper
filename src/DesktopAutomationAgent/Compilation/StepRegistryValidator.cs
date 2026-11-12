using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Json.Schema;

namespace DesktopAutomationAgent.Compilation;

public sealed class StepRegistryValidationResult
{
    public bool IsValid => Errors.Count == 0;

    public string Path { get; init; } = string.Empty;

    public string? Sha256 { get; init; }

    public StepRegistryDocument? Document { get; init; }

    public List<string> Errors { get; init; } = [];
}

public sealed class StepRegistryValidator
{
    private static readonly Lazy<JsonSchema> Schema = new(LoadSchema);

    private static readonly Regex MappingIdPattern = new(
        @"^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled,
        BddCompilerVersions.RegexMatchTimeout);

    public StepRegistryValidationResult ValidateFile(string fullPath, string displayPath)
    {
        if (!File.Exists(fullPath))
        {
            return new StepRegistryValidationResult
            {
                Path = displayPath,
                Errors = [$"{displayPath}: step registry file not found."]
            };
        }

        byte[] bytes;
        try
        {
            bytes = File.ReadAllBytes(fullPath);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return new StepRegistryValidationResult
            {
                Path = displayPath,
                Errors = [$"{displayPath}: failed to read ({ex.Message})."]
            };
        }

        return ValidateBytes(bytes, displayPath);
    }

    public StepRegistryValidationResult ValidateBytes(byte[] bytes, string displayPath)
    {
        var sha = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(bytes)).ToLowerInvariant();
        JsonNode? node;
        try
        {
            node = JsonNode.Parse(bytes);
        }
        catch (JsonException ex)
        {
            return new StepRegistryValidationResult
            {
                Path = displayPath,
                Sha256 = sha,
                Errors = [$"{displayPath}: invalid JSON ({ex.Message})."]
            };
        }

        if (node is null)
        {
            return new StepRegistryValidationResult
            {
                Path = displayPath,
                Sha256 = sha,
                Errors = [$"{displayPath}: empty JSON document."]
            };
        }

        var evaluation = Schema.Value.Evaluate(node, new EvaluationOptions
        {
            OutputFormat = OutputFormat.List,
            RequireFormatValidation = true
        });

        var errors = new List<string>();
        if (!evaluation.IsValid)
        {
            foreach (var detail in evaluation.Details.Where(d => d.HasErrors))
            {
                if (detail.Errors is null)
                    continue;
                foreach (var error in detail.Errors)
                    errors.Add($"{displayPath}: schema {detail.InstanceLocation}: {error.Value}");
            }

            if (errors.Count == 0)
                errors.Add($"{displayPath}: failed JSON Schema validation.");
        }

        StepRegistryDocument? document = null;
        try
        {
            document = JsonSerializer.Deserialize<StepRegistryDocument>(bytes, CompilationJson.DocumentRead);
        }
        catch (Exception ex) when (ex is JsonException or NotSupportedException)
        {
            errors.Add($"{displayPath}: failed to deserialize ({ex.Message}).");
        }

        if (document is not null)
        {
            if (document.ExtensionData is { Count: > 0 })
            {
                foreach (var key in document.ExtensionData.Keys.OrderBy(k => k, StringComparer.Ordinal))
                    errors.Add($"{displayPath}: unknown property '{key}'.");
            }

            var seenIds = new HashSet<string>(StringComparer.Ordinal);
            for (var i = 0; i < document.Mappings.Count; i++)
            {
                var mapping = document.Mappings[i];
                var loc = $"{displayPath}: mappings[{i}]";
                if (!MappingIdPattern.IsMatch(mapping.MappingId))
                    errors.Add($"{loc}: invalid mappingId '{mapping.MappingId}'.");
                else if (!seenIds.Add(mapping.MappingId))
                    errors.Add($"{loc}: duplicate mappingId '{mapping.MappingId}'.");

                if (mapping.Keywords.Count == 0)
                    errors.Add($"{loc}: keywords must not be empty.");

                CompilePattern(mapping, loc, errors);
            }
        }

        return new StepRegistryValidationResult
        {
            Path = displayPath,
            Sha256 = sha,
            Document = errors.Count == 0 ? document : null,
            Errors = errors
        };
    }

    internal static void CompilePattern(StepMapping mapping, string location, List<string> errors)
    {
        try
        {
            string pattern;
            if (mapping.Pattern.Kind.Equals("exact", StringComparison.OrdinalIgnoreCase))
            {
                pattern = ExactPatternCompiler.ToRegex(mapping.Pattern.Value, mapping.Parameters);
            }
            else
            {
                pattern = mapping.Pattern.Value;
                // Regex mappings must match the entire step text (no partial hits).
                if (!pattern.StartsWith('^') || !pattern.EndsWith('$'))
                {
                    errors.Add($"{location}: regex patterns must be anchored with '^' and '$'.");
                    return;
                }
            }

            // Reject catastrophic constructs with a length / nesting budget.
            if (pattern.Length > 2048)
            {
                errors.Add($"{location}: pattern exceeds 2048 characters.");
                return;
            }

            if (CountOpenGroups(pattern) > 32)
            {
                errors.Add($"{location}: pattern has too many capturing groups (max 32).");
                return;
            }

            mapping.CompiledRegex = new Regex(
                pattern,
                RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.ExplicitCapture,
                BddCompilerVersions.RegexMatchTimeout);

            // Force compile + one dry match to surface RegexParseException early.
            _ = mapping.CompiledRegex.IsMatch(string.Empty);
        }
        catch (RegexMatchTimeoutException)
        {
            errors.Add($"{location}: pattern match timed out during validation (possible ReDoS).");
        }
        catch (ArgumentException ex)
        {
            errors.Add($"{location}: invalid regex ({ex.Message}).");
        }
    }

    private static int CountOpenGroups(string pattern)
    {
        var count = 0;
        for (var i = 0; i < pattern.Length; i++)
        {
            if (pattern[i] == '(' && (i == 0 || pattern[i - 1] != '\\'))
                count++;
        }

        return count;
    }

    private static JsonSchema LoadSchema()
    {
        var resource = "DesktopAutomationAgent.Compilation.step-registry.schema.json";
        using var stream = typeof(StepRegistryValidator).Assembly.GetManifestResourceStream(resource)
            ?? throw new InvalidOperationException($"Missing embedded schema '{resource}'.");
        using var reader = new StreamReader(stream, Encoding.UTF8);
        return JsonSchema.FromText(reader.ReadToEnd());
    }
}

internal static class ExactPatternCompiler
{
    private static readonly Regex Placeholder = new(
        @"\{([A-Za-z][A-Za-z0-9_]*)\}",
        RegexOptions.CultureInvariant | RegexOptions.Compiled,
        BddCompilerVersions.RegexMatchTimeout);

    public static string ToRegex(string exact, IReadOnlyList<StepParameter> parameters)
    {
        var paramTypes = parameters.ToDictionary(p => p.Name, p => p.Type, StringComparer.Ordinal);
        var sb = new StringBuilder("^");
        var index = 0;
        foreach (Match match in Placeholder.Matches(exact))
        {
            sb.Append(Regex.Escape(exact[index..match.Index]));
            var name = match.Groups[1].Value;
            var type = paramTypes.TryGetValue(name, out var t) ? t : "string";
            sb.Append(type switch
            {
                "int" => $"(?<{name}>[+-]?[0-9]+)",
                "bool" => $"(?<{name}>true|false)",
                "objectRef" => $"(?<{name}>[a-z][a-z0-9-]{{0,63}}\\.[a-z][a-z0-9-]{{0,63}})",
                _ => $"(?<{name}>[^\\\"]+)"
            });
            index = match.Index + match.Length;
        }

        sb.Append(Regex.Escape(exact[index..]));
        sb.Append('$');
        return sb.ToString();
    }
}
