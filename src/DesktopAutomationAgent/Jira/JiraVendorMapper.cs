using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using DesktopAutomationAgent.Compilation;
using DesktopAutomationAgent.Configuration;

namespace DesktopAutomationAgent.Jira;

/// <summary>
/// Maps vendor/plugin JSON into the Phase 4D normalized model using configurable
/// field paths. Profiles only seed defaults — callers may override paths.
/// </summary>
internal static class JiraVendorMapper
{
    public static JiraFieldMappingOptions ResolveMapping(JiraOptions options)
    {
        var profile = (options.MappingProfile ?? "plain-fields").Trim().ToLowerInvariant();
        var defaults = ProfileDefaults(profile);
        var configured = options.FieldMapping ?? new JiraFieldMappingOptions();

        return new JiraFieldMappingOptions
        {
            SummaryPath = Prefer(configured.SummaryPath, defaults.SummaryPath)!,
            DescriptionPath = Prefer(configured.DescriptionPath, defaults.DescriptionPath),
            FeaturePath = Prefer(configured.FeaturePath, defaults.FeaturePath),
            NormalizedDocumentPath = Prefer(configured.NormalizedDocumentPath, defaults.NormalizedDocumentPath),
            RevisionPath = Prefer(configured.RevisionPath, defaults.RevisionPath),
            UpdatedAtPath = Prefer(configured.UpdatedAtPath, defaults.UpdatedAtPath),
            TagsPath = Prefer(configured.TagsPath, defaults.TagsPath),
            KeyPath = Prefer(configured.KeyPath, defaults.KeyPath)!,
            ExtractAdfText = configured.ExtractAdfText
        };
    }

    public static IReadOnlyList<string> ResolveIssueFields(JiraOptions options, JiraFieldMappingOptions mapping)
    {
        if (!string.IsNullOrWhiteSpace(options.IssueFields))
        {
            return options.IssueFields
                .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Distinct(StringComparer.Ordinal)
                .ToArray();
        }

        var fields = new HashSet<string>(StringComparer.Ordinal)
        {
            "summary",
            "updated",
            "description",
            "labels"
        };

        foreach (var path in new[]
                 {
                     mapping.SummaryPath,
                     mapping.DescriptionPath,
                     mapping.FeaturePath,
                     mapping.NormalizedDocumentPath,
                     mapping.RevisionPath,
                     mapping.UpdatedAtPath,
                     mapping.TagsPath
                 })
        {
            if (string.IsNullOrWhiteSpace(path))
                continue;
            if (path.StartsWith("fields.", StringComparison.OrdinalIgnoreCase))
                fields.Add(path["fields.".Length..]);
        }

        return fields.OrderBy(f => f, StringComparer.Ordinal).ToArray();
    }

    public static bool TryMap(
        JsonElement root,
        JiraFieldMappingOptions mapping,
        string expectedKey,
        out JiraTestCaseDocument document,
        out JiraSourceMetadata metadata,
        out List<string> errors,
        out List<string> redactedPaths)
    {
        document = new JiraTestCaseDocument();
        metadata = new JiraSourceMetadata();
        errors = [];
        redactedPaths = [];

        var key = ReadString(root, mapping.KeyPath)?.Trim();
        if (string.IsNullOrWhiteSpace(key))
        {
            errors.Add("mapped issue is missing key.");
            return false;
        }

        if (!string.Equals(key, expectedKey, StringComparison.Ordinal))
        {
            errors.Add($"mapped issue key '{key}' does not match requested '{expectedKey}'.");
            return false;
        }

        // Prefer an already-normalized Phase 4D document when configured.
        if (!string.IsNullOrWhiteSpace(mapping.NormalizedDocumentPath)
            && TryGetElement(root, mapping.NormalizedDocumentPath!, out var normalizedNode))
        {
            if (!TryParseNormalized(normalizedNode, expectedKey, out document, out var normErrors))
            {
                errors.AddRange(normErrors);
                return false;
            }
        }
        else
        {
            var summary = ReadPlainText(root, mapping.SummaryPath, mapping.ExtractAdfText);
            if (string.IsNullOrWhiteSpace(summary))
            {
                errors.Add("mapped issue is missing summary.");
                return false;
            }

            var feature = !string.IsNullOrWhiteSpace(mapping.FeaturePath)
                ? ReadPlainText(root, mapping.FeaturePath!, mapping.ExtractAdfText)
                : null;
            if (string.IsNullOrWhiteSpace(feature) && !string.IsNullOrWhiteSpace(mapping.DescriptionPath))
                feature = ReadPlainText(root, mapping.DescriptionPath!, mapping.ExtractAdfText);

            if (string.IsNullOrWhiteSpace(feature))
            {
                errors.Add(
                    "mapped issue is missing BDD/feature text. Configure Jira:FieldMapping:FeaturePath " +
                    "(or NormalizedDocumentPath) for the active mapping profile.");
                return false;
            }

            var revision = !string.IsNullOrWhiteSpace(mapping.RevisionPath)
                ? ReadString(root, mapping.RevisionPath!)
                : null;
            DateTimeOffset? updatedAt = null;
            if (!string.IsNullOrWhiteSpace(mapping.UpdatedAtPath))
            {
                var updatedRaw = ReadString(root, mapping.UpdatedAtPath!);
                if (!string.IsNullOrWhiteSpace(updatedRaw)
                    && DateTimeOffset.TryParse(
                        updatedRaw,
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.RoundtripKind,
                        out var parsed))
                {
                    updatedAt = parsed.ToUniversalTime();
                }
            }

            if (string.IsNullOrWhiteSpace(revision) && updatedAt is null)
            {
                errors.Add("mapped issue is missing revision and updatedAt.");
                return false;
            }

            var tags = ReadStringArray(root, mapping.TagsPath);

            document = new JiraTestCaseDocument
            {
                SchemaVersion = 1,
                JiraKey = key,
                Revision = string.IsNullOrWhiteSpace(revision) ? null : revision.Trim(),
                UpdatedAt = updatedAt,
                Summary = summary.Trim(),
                Description = null, // never cache unrestricted description dump
                Feature = feature,
                Tags = tags
            };
        }

        redactedPaths.AddRange(CollectRedactedPaths(root, mapping));
        var contentSha = Sha256Hex(CompilationJson.CanonicalSerializeUtf8(document));
        metadata = new JiraSourceMetadata
        {
            JiraKey = key,
            MappingProfile = "resolved",
            RevisionToken = document.ResolveRevisionToken(),
            UpdatedAt = document.UpdatedAt?.UtcDateTime.ToString("O"),
            ContentSha256 = contentSha,
            FetchedAtUtc = DateTimeOffset.UtcNow.ToString("O"),
            SourceKind = "jira-rest",
            Redacted = true,
            RedactedFieldPaths = redactedPaths
        };
        return true;
    }

    private static JiraFieldMappingOptions ProfileDefaults(string profile) =>
        profile switch
        {
            // Plain Jira issue fields — BDD in description or a configured custom field.
            "plain-fields" => new JiraFieldMappingOptions
            {
                SummaryPath = "fields.summary",
                DescriptionPath = "fields.description",
                FeaturePath = "fields.description",
                UpdatedAtPath = "fields.updated",
                TagsPath = "fields.labels",
                KeyPath = "key",
                ExtractAdfText = true
            },
            // Xray-style: cucumber/scenario custom field path must be configured per instance.
            // Defaults only set common labels; FeaturePath remains override-required when empty
            // description is insufficient. Callers should set FieldMapping:FeaturePath to the
            // Xray Cucumber Tests custom field id for their site.
            "xray" => new JiraFieldMappingOptions
            {
                SummaryPath = "fields.summary",
                DescriptionPath = "fields.description",
                FeaturePath = "fields.customfield_xray_cucumber",
                UpdatedAtPath = "fields.updated",
                TagsPath = "fields.labels",
                KeyPath = "key",
                ExtractAdfText = true
            },
            // Zephyr Scale-style: BDD/steps custom field path is instance-specific.
            "zephyr" => new JiraFieldMappingOptions
            {
                SummaryPath = "fields.summary",
                DescriptionPath = "fields.description",
                FeaturePath = "fields.customfield_zephyr_bdd",
                UpdatedAtPath = "fields.updated",
                TagsPath = "fields.labels",
                KeyPath = "key",
                ExtractAdfText = true
            },
            "normalized" => new JiraFieldMappingOptions
            {
                SummaryPath = "fields.summary",
                NormalizedDocumentPath = "fields.customfield_normalized_testcase",
                UpdatedAtPath = "fields.updated",
                KeyPath = "key",
                ExtractAdfText = false
            },
            _ => new JiraFieldMappingOptions()
        };

    private static string? Prefer(string? configured, string? fallback) =>
        string.IsNullOrWhiteSpace(configured) ? fallback : configured;

    private static bool TryParseNormalized(
        JsonElement node,
        string expectedKey,
        out JiraTestCaseDocument document,
        out List<string> errors)
    {
        document = new JiraTestCaseDocument();
        errors = [];
        try
        {
            string json = node.ValueKind switch
            {
                JsonValueKind.String => node.GetString() ?? string.Empty,
                JsonValueKind.Object => node.GetRawText(),
                _ => string.Empty
            };
            if (string.IsNullOrWhiteSpace(json))
            {
                errors.Add("normalized document path was empty.");
                return false;
            }

            var parsed = JsonSerializer.Deserialize<JiraTestCaseDocument>(json, CompilationJson.DocumentRead);
            if (parsed is null)
            {
                errors.Add("normalized document deserialized to null.");
                return false;
            }

            if (!string.Equals(parsed.JiraKey, expectedKey, StringComparison.Ordinal))
            {
                errors.Add(
                    $"normalized document jiraKey '{parsed.JiraKey}' does not match requested '{expectedKey}'.");
                return false;
            }

            document = parsed;
            return true;
        }
        catch (JsonException ex)
        {
            errors.Add($"normalized document JSON parse failed: {ex.Message}");
            return false;
        }
    }

    private static List<string> CollectRedactedPaths(JsonElement root, JiraFieldMappingOptions mapping)
    {
        var kept = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var path in new[]
                 {
                     mapping.KeyPath,
                     mapping.SummaryPath,
                     mapping.DescriptionPath,
                     mapping.FeaturePath,
                     mapping.NormalizedDocumentPath,
                     mapping.RevisionPath,
                     mapping.UpdatedAtPath,
                     mapping.TagsPath
                 })
        {
            if (!string.IsNullOrWhiteSpace(path))
                kept.Add(path!);
        }

        var redacted = new List<string>();
        if (root.ValueKind == JsonValueKind.Object && root.TryGetProperty("fields", out var fields)
            && fields.ValueKind == JsonValueKind.Object)
        {
            foreach (var prop in fields.EnumerateObject())
            {
                var path = "fields." + prop.Name;
                if (!kept.Contains(path))
                    redacted.Add(path);
            }
        }

        return redacted;
    }

    internal static bool TryGetElement(JsonElement root, string path, out JsonElement element)
    {
        element = default;
        if (string.IsNullOrWhiteSpace(path))
            return false;

        var current = root;
        foreach (var segment in path.Split('.', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (current.ValueKind != JsonValueKind.Object || !current.TryGetProperty(segment, out current))
                return false;
        }

        element = current;
        return element.ValueKind is not JsonValueKind.Undefined and not JsonValueKind.Null;
    }

    internal static string? ReadString(JsonElement root, string path)
    {
        if (!TryGetElement(root, path, out var el))
            return null;
        return el.ValueKind switch
        {
            JsonValueKind.String => el.GetString(),
            JsonValueKind.Number => el.GetRawText(),
            JsonValueKind.True => "true",
            JsonValueKind.False => "false",
            _ => null
        };
    }

    internal static string? ReadPlainText(JsonElement root, string path, bool extractAdf)
    {
        if (!TryGetElement(root, path, out var el))
            return null;

        return el.ValueKind switch
        {
            JsonValueKind.String => el.GetString(),
            JsonValueKind.Object when extractAdf => ExtractAdf(el),
            JsonValueKind.Object => el.GetRawText(),
            JsonValueKind.Array => string.Join('\n', el.EnumerateArray().Select(e => e.GetString()).Where(s => !string.IsNullOrWhiteSpace(s))),
            _ => null
        };
    }

    internal static List<string>? ReadStringArray(JsonElement root, string? path)
    {
        if (string.IsNullOrWhiteSpace(path) || !TryGetElement(root, path!, out var el))
            return null;

        if (el.ValueKind == JsonValueKind.Array)
        {
            var list = new List<string>();
            foreach (var item in el.EnumerateArray())
            {
                if (item.ValueKind == JsonValueKind.String)
                {
                    var s = item.GetString();
                    if (!string.IsNullOrWhiteSpace(s))
                        list.Add(s!);
                }
                else if (item.ValueKind == JsonValueKind.Object && item.TryGetProperty("name", out var name)
                         && name.ValueKind == JsonValueKind.String)
                {
                    var s = name.GetString();
                    if (!string.IsNullOrWhiteSpace(s))
                        list.Add(s!);
                }
            }

            return list.Count == 0 ? null : list.Distinct(StringComparer.Ordinal).ToList();
        }

        if (el.ValueKind == JsonValueKind.String)
        {
            var raw = el.GetString();
            if (string.IsNullOrWhiteSpace(raw))
                return null;
            return raw.Split([',', ' ', ';'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Distinct(StringComparer.Ordinal)
                .ToList();
        }

        return null;
    }

    internal static string? ExtractAdf(JsonElement adf)
    {
        if (adf.ValueKind != JsonValueKind.Object)
            return null;

        var sb = new StringBuilder();
        void Walk(JsonElement node)
        {
            if (node.ValueKind == JsonValueKind.Object)
            {
                if (node.TryGetProperty("type", out var type)
                    && type.ValueKind == JsonValueKind.String
                    && type.GetString() == "text"
                    && node.TryGetProperty("text", out var text)
                    && text.ValueKind == JsonValueKind.String)
                {
                    sb.Append(text.GetString());
                }

                if (node.TryGetProperty("content", out var content) && content.ValueKind == JsonValueKind.Array)
                {
                    foreach (var child in content.EnumerateArray())
                        Walk(child);
                    if (node.TryGetProperty("type", out var blockType)
                        && blockType.ValueKind == JsonValueKind.String
                        && blockType.GetString() is "paragraph" or "heading")
                    {
                        sb.Append('\n');
                    }
                }
            }
            else if (node.ValueKind == JsonValueKind.Array)
            {
                foreach (var child in node.EnumerateArray())
                    Walk(child);
            }
        }

        Walk(adf);
        var result = sb.ToString().Trim();
        return string.IsNullOrWhiteSpace(result) ? null : result;
    }

    internal static string Sha256Hex(byte[] bytes)
    {
        var hash = SHA256.HashData(bytes);
        return Convert.ToHexString(hash).ToLowerInvariant();
    }
}
