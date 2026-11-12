using System.Globalization;
using System.Text.RegularExpressions;

namespace DesktopAutomationAgent.Compilation;

public sealed class StepMatch
{
    public required StepMapping Mapping { get; init; }

    public IReadOnlyDictionary<string, string> Parameters { get; init; } =
        new Dictionary<string, string>(StringComparer.Ordinal);
}

public sealed class StepMatchResult
{
    public StepMatch? Match { get; init; }

    public List<string> Errors { get; init; } = [];

    public bool IsSuccess => Match is not null && Errors.Count == 0;
}

public sealed class StepMatcher
{
    private readonly IReadOnlyList<StepMapping> _mappings;

    public StepMatcher(StepRegistryDocument registry)
    {
        ArgumentNullException.ThrowIfNull(registry);
        _mappings = registry.Mappings;
        foreach (var mapping in _mappings)
        {
            if (mapping.CompiledRegex is null)
            {
                var errors = new List<string>();
                StepRegistryValidator.CompilePattern(mapping, mapping.MappingId, errors);
                if (errors.Count > 0)
                    throw new InvalidOperationException(string.Join("; ", errors));
            }
        }
    }

    public StepMatchResult Match(BddIrStep step)
    {
        var candidates = new List<StepMatch>();
        var errors = new List<string>();

        foreach (var mapping in _mappings)
        {
            if (!mapping.Keywords.Any(k => k.Equals(step.Keyword, StringComparison.OrdinalIgnoreCase)))
                continue;

            Match regexMatch;
            try
            {
                regexMatch = mapping.CompiledRegex!.Match(step.Text);
            }
            catch (RegexMatchTimeoutException)
            {
                errors.Add(
                    $"step '{step.Keyword} {step.Text}' (line {step.Line}): regex timeout matching mapping '{mapping.MappingId}'.");
                continue;
            }

            if (!regexMatch.Success)
                continue;

            var parameters = new Dictionary<string, string>(StringComparer.Ordinal);
            var paramError = false;
            foreach (var parameter in mapping.Parameters)
            {
                var group = regexMatch.Groups[parameter.Name];
                if (!group.Success)
                {
                    errors.Add(
                        $"step '{step.Keyword} {step.Text}': mapping '{mapping.MappingId}' missing capture '{parameter.Name}'.");
                    paramError = true;
                    break;
                }

                if (!TryCoerce(parameter, group.Value, out var coerced, out var coerceError))
                {
                    errors.Add(
                        $"step '{step.Keyword} {step.Text}': mapping '{mapping.MappingId}' parameter '{parameter.Name}': {coerceError}");
                    paramError = true;
                    break;
                }

                parameters[parameter.Name] = coerced;
            }

            if (paramError)
                continue;

            candidates.Add(new StepMatch
            {
                Mapping = mapping,
                Parameters = parameters
            });
        }

        if (candidates.Count == 0)
        {
            errors.Add(
                $"unmapped-step: no registry mapping for '{step.Keyword} {step.Text}' (line {step.Line}).");
            return new StepMatchResult { Errors = errors };
        }

        if (candidates.Count > 1)
        {
            var ids = string.Join(", ", candidates.Select(c => c.Mapping.MappingId).OrderBy(x => x, StringComparer.Ordinal));
            errors.Add(
                $"ambiguous-step: '{step.Keyword} {step.Text}' (line {step.Line}) matched multiple mappings: {ids}.");
            return new StepMatchResult { Errors = errors };
        }

        var winner = candidates[0];
        if (string.Equals(winner.Mapping.Risk, "prohibited", StringComparison.OrdinalIgnoreCase))
        {
            errors.Add(
                $"prohibited-risk: mapping '{winner.Mapping.MappingId}' is risk=prohibited and cannot be compiled.");
            return new StepMatchResult { Errors = errors };
        }

        return new StepMatchResult { Match = winner, Errors = errors };
    }

    private static bool TryCoerce(
        StepParameter parameter,
        string raw,
        out string coerced,
        out string error)
    {
        coerced = raw;
        error = string.Empty;

        // Reject control characters / escapes that could break JSON templates.
        if (raw.IndexOfAny(['\0', '\r', '\n', '\t']) >= 0)
        {
            error = "parameter contains forbidden control characters.";
            return false;
        }

        switch (parameter.Type)
        {
            case "int":
                if (!long.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out _))
                {
                    error = $"value '{raw}' is not an invariant int.";
                    return false;
                }

                coerced = long.Parse(raw, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture);
                return true;
            case "bool":
                if (!bool.TryParse(raw, out var b))
                {
                    error = $"value '{raw}' is not a bool.";
                    return false;
                }

                coerced = b ? "true" : "false";
                return true;
            case "objectRef":
                if (!Regex.IsMatch(
                        raw,
                        @"^[a-z][a-z0-9-]{0,63}\.[a-z][a-z0-9-]{0,63}$",
                        RegexOptions.CultureInvariant,
                        BddCompilerVersions.RegexMatchTimeout))
                {
                    error = $"value '{raw}' is not a valid objectRef.";
                    return false;
                }

                coerced = raw;
                return true;
            default:
                coerced = raw;
                return true;
        }
    }
}
