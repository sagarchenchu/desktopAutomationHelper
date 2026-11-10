using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Plans;
using DesktopAutomationAgent.Workspace;
using Microsoft.Extensions.Logging;

namespace DesktopAutomationAgent.Compilation;

public sealed class DeterministicBddPlanCompiler : IDeterministicBddPlanCompiler
{
    private static readonly JsonSerializerOptions ReadOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true
    };

    private static readonly JsonSerializerOptions WriteOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
    };

    private static readonly Regex SafePlanIdPart = new(
        @"[^A-Za-z0-9._-]+",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private readonly IWorkspaceManager _workspace;
    private readonly PlanValidator _planValidator = new();
    private readonly ILogger<DeterministicBddPlanCompiler> _logger;

    public DeterministicBddPlanCompiler(
        IWorkspaceManager workspace,
        ILogger<DeterministicBddPlanCompiler> logger)
    {
        _workspace = workspace;
        _logger = logger;
    }

    public BddCompileResult Compile(BddCompileRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        var errors = new List<string>();
        var warnings = new List<string>();

        if (string.IsNullOrWhiteSpace(request.MapPath))
        {
            return Fail(request, errors: ["Map path is required."]);
        }

        string mapFullPath;
        try
        {
            mapFullPath = _workspace.ResolveSafePath(request.MapPath);
        }
        catch (WorkspaceException ex)
        {
            return Fail(request, errors: [ex.Message]);
        }

        if (!File.Exists(mapFullPath))
            return Fail(request, errors: [$"BDD action map not found: '{mapFullPath}'."]);

        BddActionMapDocument map;
        try
        {
            map = JsonSerializer.Deserialize<BddActionMapDocument>(
                      File.ReadAllText(mapFullPath), ReadOptions)
                  ?? throw new JsonException("BDD action map deserialized to null.");
        }
        catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException)
        {
            return Fail(request, errors: [$"Failed to read BDD action map: {ex.Message}"]);
        }

        ValidateMap(map, errors, warnings);
        if (errors.Count > 0)
            return Fail(request, mapPath: mapFullPath, errors: errors, warnings: warnings);

        string recordingFullPath;
        try
        {
            recordingFullPath = ResolveRecordingPath(mapFullPath, map.SourceRecording, request.RecordingPath);
        }
        catch (Exception ex) when (ex is WorkspaceException or InvalidOperationException or IOException)
        {
            return Fail(
                request,
                mapPath: mapFullPath,
                errors: [ex.Message],
                warnings: warnings,
                jiraKey: map.JiraKey,
                recordingId: map.RecordingId);
        }

        RecordingDocument recording;
        try
        {
            recording = JsonSerializer.Deserialize<RecordingDocument>(
                            File.ReadAllText(recordingFullPath), ReadOptions)
                        ?? throw new JsonException("Recording deserialized to null.");
        }
        catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException)
        {
            return Fail(
                request,
                mapPath: mapFullPath,
                recordingPath: recordingFullPath,
                errors: [$"Failed to read recording: {ex.Message}"],
                warnings: warnings,
                jiraKey: map.JiraKey,
                recordingId: map.RecordingId);
        }

        var byEventId = IndexRecording(recording, warnings);
        if (map.UnmappedEventIds.Count > 0)
        {
            warnings.Add(
                $"{map.UnmappedEventIds.Count} unmapped event(s) were not compiled " +
                "(unsupported or non-BDD Assistive actions).");
        }

        var steps = new List<PlanStep>();
        var skipped = 0;

        foreach (var group in map.BddGroups)
        {
            if (group.Actions.Count == 0)
            {
                warnings.Add($"BDD group '{group.GroupId}' has no actions.");
                continue;
            }

            var groupStepsBefore = steps.Count;
            foreach (var actionRef in group.Actions.OrderBy(a => a.Sequence).ThenBy(a => a.EventId, StringComparer.Ordinal))
            {
                if (!byEventId.TryGetValue(actionRef.EventId, out var recorded))
                {
                    skipped++;
                    errors.Add(
                        $"Group '{group.GroupId}' action '{actionRef.EventId}': " +
                        "eventId not found in source recording.");
                    continue;
                }

                var operation = string.IsNullOrWhiteSpace(actionRef.Operation)
                    ? null
                    : actionRef.Operation.Trim();
                if (string.IsNullOrWhiteSpace(operation))
                {
                    skipped++;
                    errors.Add(
                        $"Group '{group.GroupId}' action '{actionRef.EventId}': operation is required.");
                    continue;
                }

                if (!PlaybackArgumentBuilder.TryBuildArguments(
                        operation,
                        recorded,
                        preferObjectRefs: request.PreferObjectRefs && !string.IsNullOrWhiteSpace(request.ObjectRepositoryPath),
                        out var arguments,
                        out var skipReason))
                {
                    skipped++;
                    errors.Add(
                        $"Group '{group.GroupId}' action '{actionRef.EventId}' ({operation}): {skipReason}");
                    continue;
                }

                var stepId = BuildStepId(group.GroupId, actionRef.EventId);
                steps.Add(new PlanStep
                {
                    Id = stepId,
                    Operation = operation,
                    Arguments = arguments,
                    Sensitive = PlaybackArgumentBuilder.IsSensitiveOperation(operation)
                });
            }

            if (steps.Count == groupStepsBefore)
            {
                warnings.Add(
                    $"BDD group '{group.GroupId}' ('{Truncate(group.Statement, 80)}') produced no steps.");
            }
        }

        if (errors.Count > 0)
        {
            return Fail(
                request,
                mapPath: mapFullPath,
                recordingPath: recordingFullPath,
                errors: errors,
                warnings: warnings,
                jiraKey: map.JiraKey,
                recordingId: map.RecordingId,
                skipped: skipped,
                unmapped: map.UnmappedEventIds.Count);
        }

        if (steps.Count == 0)
        {
            errors.Add("Compilation produced no plan steps. Ensure bddGroups contain compilable actions.");
            return Fail(
                request,
                mapPath: mapFullPath,
                recordingPath: recordingFullPath,
                errors: errors,
                warnings: warnings,
                jiraKey: map.JiraKey,
                recordingId: map.RecordingId,
                skipped: skipped,
                unmapped: map.UnmappedEventIds.Count);
        }

        var planId = BuildPlanId(map.JiraKey, map.RecordingId);
        var plan = new PlanManifest
        {
            Schema = "../schemas/plan.schema.json",
            SchemaVersion = 1,
            CatalogSchemaVersion = 2,
            PlanId = planId,
            Name = $"{map.JiraKey} compiled from {map.RecordingId}",
            Description =
                "Deterministic Phase 4 plan compiled from Assistive BDD action map + recording. " +
                "BDD statement text is preserved as metadata only and is not interpreted.",
            Tags = ["phase4", "compiled-from-assistive", map.JiraKey],
            Metadata = BuildMetadata(map, mapFullPath, recordingFullPath),
            ObjectRepository = request.PreferObjectRefs ? request.ObjectRepositoryPath : null,
            Steps = steps
        };

        var validation = _planValidator.Validate(plan, $"plans/{planId}.plan.json");
        if (!validation.IsValid)
        {
            errors.AddRange(validation.Errors);
            return Fail(
                request,
                mapPath: mapFullPath,
                recordingPath: recordingFullPath,
                errors: errors,
                warnings: warnings,
                jiraKey: map.JiraKey,
                recordingId: map.RecordingId,
                skipped: skipped,
                unmapped: map.UnmappedEventIds.Count,
                planId: planId);
        }

        string outputFullPath;
        try
        {
            outputFullPath = ResolveOutputPath(request.OutputPath, planId);
            WritePlanAtomically(outputFullPath, plan);
        }
        catch (Exception ex) when (ex is WorkspaceException or IOException or UnauthorizedAccessException)
        {
            return Fail(
                request,
                mapPath: mapFullPath,
                recordingPath: recordingFullPath,
                errors: [$"Failed to write plan: {ex.Message}"],
                warnings: warnings,
                jiraKey: map.JiraKey,
                recordingId: map.RecordingId,
                skipped: skipped,
                unmapped: map.UnmappedEventIds.Count,
                planId: planId);
        }

        _logger.LogInformation(
            "Compiled BDD map {Map} + recording {Recording} into plan {PlanId} ({Steps} steps).",
            mapFullPath,
            recordingFullPath,
            planId,
            steps.Count);

        return new BddCompileResult
        {
            Success = true,
            MapPath = mapFullPath,
            RecordingPath = recordingFullPath,
            OutputPath = outputFullPath,
            PlanId = planId,
            JiraKey = map.JiraKey,
            RecordingId = map.RecordingId,
            StepCount = steps.Count,
            SkippedActionCount = skipped,
            UnmappedEventCount = map.UnmappedEventIds.Count,
            Plan = plan,
            Errors = errors,
            Warnings = warnings
        };
    }

    private static void ValidateMap(BddActionMapDocument map, List<string> errors, List<string> warnings)
    {
        if (map.SchemaVersion != 1)
            errors.Add($"Unsupported bdd-action-map schemaVersion {map.SchemaVersion}; expected 1.");

        if (string.IsNullOrWhiteSpace(map.RecordingId))
            errors.Add("recordingId is required.");

        if (!JiraKeyContract.TryValidate(
                map.JiraKey,
                projectRegex: null,
                out _,
                out var jiraError,
                trimSurroundingWhitespace: false))
        {
            errors.Add($"jiraKey: {jiraError}");
        }

        if (string.IsNullOrWhiteSpace(map.SourceRecording))
            errors.Add("sourceRecording is required.");

        if (map.BddGroups is null)
            errors.Add("bddGroups is required.");
        else if (map.BddGroups.Count == 0)
            warnings.Add("bddGroups is empty; compilation will produce no steps unless fixed.");

        if (map.Warnings is { Count: > 0 })
            warnings.AddRange(map.Warnings.Select(w => $"map warning: {w}"));
    }

    private string ResolveRecordingPath(string mapFullPath, string sourceRecording, string? explicitPath)
    {
        if (!string.IsNullOrWhiteSpace(explicitPath))
            return _workspace.ResolveSafePath(explicitPath);

        // Assistive layout: recording_<ts>.json sits at the recording output root while the map
        // lives under assistive-artifacts/<jira>/<recordingId>/. Walk up a few parents.
        var mapDir = Path.GetDirectoryName(mapFullPath)!;
        var searchRoots = new List<string> { mapDir };
        var cursor = mapDir;
        for (var depth = 0; depth < 4; depth++)
        {
            var parent = Path.GetDirectoryName(cursor);
            if (string.IsNullOrEmpty(parent) || string.Equals(parent, cursor, StringComparison.Ordinal))
                break;
            searchRoots.Add(parent);
            cursor = parent;
        }

        foreach (var root in searchRoots)
        {
            var candidate = Path.GetFullPath(Path.Combine(root, sourceRecording));
            if (!File.Exists(candidate))
                continue;

            if (TryMakeWorkspaceRelative(candidate, out var relative))
                return _workspace.ResolveSafePath(relative);
        }

        return _workspace.ResolveSafePath(sourceRecording);
    }

    private bool TryMakeWorkspaceRelative(string fullPath, out string relative)
    {
        var root = Path.GetFullPath(_workspace.RootPath)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var full = Path.GetFullPath(fullPath);
        var comparison = OperatingSystem.IsWindows()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;

        if (full.Equals(root, comparison))
        {
            relative = ".";
            return true;
        }

        var prefix = root + Path.DirectorySeparatorChar;
        if (full.StartsWith(prefix, comparison))
        {
            relative = Path.GetRelativePath(root, full);
            return true;
        }

        relative = string.Empty;
        return false;
    }

    private string ResolveOutputPath(string? requested, string planId)
    {
        if (!string.IsNullOrWhiteSpace(requested))
            return _workspace.ResolveSafePath(requested);

        var relative = Path.Combine("plans", $"{planId}.plan.json");
        return _workspace.ResolveSafePath(relative);
    }

    private static Dictionary<string, RecordingAction> IndexRecording(
        RecordingDocument recording,
        List<string> warnings)
    {
        var index = new Dictionary<string, RecordingAction>(StringComparer.Ordinal);
        foreach (var action in recording.Actions)
        {
            if (string.IsNullOrWhiteSpace(action.EventId))
            {
                warnings.Add("Recording contains an action without eventId; it cannot be joined.");
                continue;
            }

            if (!index.TryAdd(action.EventId, action))
                warnings.Add($"Duplicate recording eventId '{action.EventId}'; using the first occurrence.");
        }

        return index;
    }

    private static Dictionary<string, JsonElement> BuildMetadata(
        BddActionMapDocument map,
        string mapPath,
        string recordingPath)
    {
        var payload = new Dictionary<string, object?>
        {
            ["compiler"] = "phase4-deterministic-bdd",
            ["compilerVersion"] = 1,
            ["jiraKey"] = map.JiraKey,
            ["recordingId"] = map.RecordingId,
            ["sourceBddActionMap"] = mapPath,
            ["sourceRecording"] = recordingPath,
            ["bddGroupCount"] = map.BddGroups.Count,
            ["unmappedEventCount"] = map.UnmappedEventIds.Count,
            ["bddStatements"] = map.BddGroups
                .Select(g => new { g.GroupId, g.Statement, actionCount = g.Actions.Count })
                .ToArray()
        };

        using var doc = JsonDocument.Parse(JsonSerializer.Serialize(payload, WriteOptions));
        return doc.RootElement.EnumerateObject()
            .ToDictionary(p => p.Name, p => p.Value.Clone(), StringComparer.Ordinal);
    }

    private static string BuildPlanId(string jiraKey, string recordingId)
    {
        var left = SafePlanIdPart.Replace(jiraKey.ToLowerInvariant(), "-").Trim('-');
        var right = SafePlanIdPart.Replace(recordingId, "-").Trim('-');
        if (string.IsNullOrWhiteSpace(left))
            left = "compiled";
        if (string.IsNullOrWhiteSpace(right))
            right = "recording";

        var id = $"{left}.{right}";
        if (id.Length > 128)
            id = id[..128].TrimEnd('.', '-', '_');
        return id;
    }

    private static string BuildStepId(string groupId, string eventId)
    {
        var raw = $"{groupId}-{eventId}";
        var cleaned = SafePlanIdPart.Replace(raw, "-").Trim('-');
        return string.IsNullOrWhiteSpace(cleaned) ? "step" : cleaned;
    }

    private void WritePlanAtomically(string outputFullPath, PlanManifest plan)
    {
        var directory = Path.GetDirectoryName(outputFullPath)
                        ?? throw new IOException($"Invalid output path '{outputFullPath}'.");
        Directory.CreateDirectory(directory);

        var json = JsonSerializer.Serialize(plan, WriteOptions);
        var temp = Path.Combine(directory, $".{Path.GetFileName(outputFullPath)}.{Guid.NewGuid():N}.tmp");
        try
        {
            File.WriteAllText(temp, json, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            File.Move(temp, outputFullPath, overwrite: true);
        }
        finally
        {
            if (File.Exists(temp))
            {
                try { File.Delete(temp); } catch { /* ignore */ }
            }
        }
    }

    private static BddCompileResult Fail(
        BddCompileRequest request,
        string? mapPath = null,
        string? recordingPath = null,
        IReadOnlyList<string>? errors = null,
        IReadOnlyList<string>? warnings = null,
        string? jiraKey = null,
        string? recordingId = null,
        string? planId = null,
        int skipped = 0,
        int unmapped = 0) =>
        new()
        {
            Success = false,
            MapPath = mapPath ?? request.MapPath,
            RecordingPath = recordingPath ?? request.RecordingPath ?? string.Empty,
            PlanId = planId,
            JiraKey = jiraKey,
            RecordingId = recordingId,
            SkippedActionCount = skipped,
            UnmappedEventCount = unmapped,
            Errors = errors ?? [],
            Warnings = warnings ?? []
        };

    private static string Truncate(string value, int max) =>
        value.Length <= max ? value : value[..(max - 1)] + "…";
}
