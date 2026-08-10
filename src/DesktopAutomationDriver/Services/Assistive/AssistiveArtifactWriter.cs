using System.Text;
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
        EnsureInside(outputDirectory, artifactsRoot);

        if (Directory.Exists(artifactsRoot))
            throw new IOException($"Assistive artifact directory already exists for recording '{recordingId}'.");

        Directory.CreateDirectory(artifactsRoot);
        var pageDir = Path.Combine(artifactsRoot, "page-objects");
        Directory.CreateDirectory(pageDir);

        var pageFiles = new List<string>();
        foreach (var page in buildResult.Pages.OrderBy(p => p.PageId, StringComparer.Ordinal))
        {
            var pagePath = Path.Combine(pageDir, $"{page.PageId}.page.json");
            EnsureInside(outputDirectory, pagePath);
            WriteAtomic(pagePath, JsonSerializer.Serialize(page, JsonOpts));
            pageFiles.Add(pagePath);
        }

        string? mapPath = null;
        if (buildResult.BddActionMap is not null)
        {
            mapPath = Path.Combine(artifactsRoot, "bdd-action-map.json");
            EnsureInside(outputDirectory, mapPath);
            WriteAtomic(mapPath, JsonSerializer.Serialize(buildResult.BddActionMap, JsonOpts));
        }

        return new RecordingArtifactsSummary
        {
            Directory = artifactsRoot,
            BddActionMapFile = mapPath,
            PageObjectFiles = pageFiles,
            Warnings = warnings
        };
    }

    private static void WriteAtomic(string destinationPath, string contents)
    {
        var directory = Path.GetDirectoryName(destinationPath)
            ?? throw new InvalidOperationException("Destination directory is required.");
        Directory.CreateDirectory(directory);

        var tempPath = Path.Combine(directory, $".tmp-{Guid.NewGuid():N}.json");
        try
        {
            File.WriteAllText(tempPath, contents, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            File.Move(tempPath, destinationPath);
        }
        catch
        {
            try { if (File.Exists(tempPath)) File.Delete(tempPath); } catch { /* ignore */ }
            throw;
        }
    }

    private static void EnsureInside(string root, string candidate)
    {
        var normalizedRoot = Path.GetFullPath(root)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var normalizedCandidate = Path.GetFullPath(candidate);
        if (WorkspacePathsEqual(normalizedCandidate, normalizedRoot))
            return;

        var relative = Path.GetRelativePath(normalizedRoot, normalizedCandidate);
        if (Path.IsPathRooted(relative)
            || relative == ".."
            || relative.StartsWith(".." + Path.DirectorySeparatorChar, StringComparison.Ordinal)
            || relative.StartsWith(".." + Path.AltDirectorySeparatorChar, StringComparison.Ordinal))
        {
            throw new InvalidOperationException("Assistive artifact path escapes the recording output directory.");
        }
    }

    private static bool WorkspacePathsEqual(string left, string right) =>
        string.Equals(
            left.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
            right.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
            OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);
}
