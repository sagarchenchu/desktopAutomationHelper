using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using DesktopAutomationDriver.Models.Recording;
using DesktopAutomationDriver.Models.Recording.Assistive;

namespace DesktopAutomationDriver.Services.Assistive;

/// <summary>Atomic, containment-checked writer for Assistive recording sidecars.</summary>
public sealed class AssistiveArtifactWriter
{
    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    /// <summary>
    /// Test seam: invoked after staging content is written and path-checked,
    /// immediately before renaming the staging directory into place.
    /// </summary>
    internal Action<string /*stagingDir*/, string /*finalDir*/>? BeforeCommitStaging { get; set; }

    public RecordingArtifactsSummary Write(
        string outputDirectory,
        string recordingFileName,
        string recordingId,
        string? jiraKey,
        AssistiveArtifactBuilder.BuildResult buildResult)
    {
        var warnings = new List<string>(buildResult.Warnings);
        if (buildResult.Pages.Count == 0 && buildResult.BddActionMap is null)
            return new RecordingArtifactsSummary { Warnings = warnings };

        var scope = string.IsNullOrWhiteSpace(jiraKey) ? "unassigned" : jiraKey!;
        var artifactsRoot = Path.Combine(outputDirectory, "assistive-artifacts", scope, recordingId);
        AssistivePathSafety.EnsureWritablePathInside(artifactsRoot, outputDirectory);

        if (Directory.Exists(artifactsRoot))
            throw new IOException($"Assistive artifact directory already exists for recording '{recordingId}'.");

        var stagingRoot = Path.Combine(
            outputDirectory,
            "assistive-artifacts",
            scope,
            $".staging-{recordingId}-{Guid.NewGuid():N}");

        try
        {
            AssistivePathSafety.EnsureWritablePathInside(stagingRoot, outputDirectory);
            Directory.CreateDirectory(stagingRoot);

            var pageDir = Path.Combine(stagingRoot, "page-objects");
            Directory.CreateDirectory(pageDir);

            var stagedPageFiles = new List<string>();
            foreach (var page in buildResult.Pages.OrderBy(p => p.PageId, StringComparer.Ordinal))
            {
                var pagePath = Path.Combine(pageDir, $"{page.PageId}.page.json");
                AssistivePathSafety.EnsureWritablePathInside(pagePath, outputDirectory);
                AssistiveAtomicIO.ReplaceFileAtomic(
                    pagePath,
                    JsonSerializer.Serialize(page, JsonOpts));
                stagedPageFiles.Add(pagePath);
            }

            string? stagedMapPath = null;
            if (buildResult.BddActionMap is not null)
            {
                stagedMapPath = Path.Combine(stagingRoot, "bdd-action-map.json");
                AssistivePathSafety.EnsureWritablePathInside(stagedMapPath, outputDirectory);
                AssistiveAtomicIO.ReplaceFileAtomic(
                    stagedMapPath,
                    JsonSerializer.Serialize(buildResult.BddActionMap, JsonOpts));
            }

            // Final destination parents must also be free of reparse-point escapes.
            var scopeDir = Path.GetDirectoryName(artifactsRoot)!;
            Directory.CreateDirectory(scopeDir);
            AssistivePathSafety.EnsureWritablePathInside(artifactsRoot, outputDirectory);

            BeforeCommitStaging?.Invoke(stagingRoot, artifactsRoot);

            Directory.Move(stagingRoot, artifactsRoot);

            var pageFiles = stagedPageFiles
                .Select(p => Path.Combine(artifactsRoot, "page-objects", Path.GetFileName(p)))
                .ToList();
            string? mapPath = stagedMapPath is null
                ? null
                : Path.Combine(artifactsRoot, "bdd-action-map.json");

            return new RecordingArtifactsSummary
            {
                Directory = artifactsRoot,
                BddActionMapFile = mapPath,
                PageObjectFiles = pageFiles,
                Warnings = warnings
            };
        }
        catch
        {
            AssistiveAtomicIO.TryDeleteDirectory(stagingRoot);
            throw;
        }
    }
}
