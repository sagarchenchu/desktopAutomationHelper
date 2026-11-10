using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Workspace;

namespace DesktopAutomationAgent.Exploration;

public sealed class ExplorationArtifactWriter
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) }
    };

    private static readonly JsonSerializerOptions JsonlOptions = new()
    {
        WriteIndented = false,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) }
    };

    private readonly IWorkspaceManager _workspace;

    public ExplorationArtifactWriter(IWorkspaceManager workspace)
    {
        _workspace = workspace;
    }

    public string CreateRunDirectory(string explorationId)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(explorationId);
        _workspace.EnsureInitialized();

        var root = Path.Combine(_workspace.RootPath, "exploration", "runs", explorationId);
        if (Directory.Exists(root))
            throw new IOException($"Exploration artifact directory already exists: {root}");

        Directory.CreateDirectory(root);
        return root;
    }

    public void WriteAll(
        string runDirectory,
        ExplorationManifest manifest,
        IReadOnlyList<ExplorationStateRecord> states,
        IReadOnlyList<ExplorationTransitionRecord> transitions,
        ExplorationSummary summary)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(runDirectory);
        ArgumentNullException.ThrowIfNull(manifest);
        ArgumentNullException.ThrowIfNull(states);
        ArgumentNullException.ThrowIfNull(transitions);
        ArgumentNullException.ThrowIfNull(summary);

        WriteAtomicJson(Path.Combine(runDirectory, "manifest.json"), manifest);
        WriteAtomicJsonl(Path.Combine(runDirectory, "states.jsonl"), states);
        WriteAtomicJsonl(Path.Combine(runDirectory, "transitions.jsonl"), transitions);
        WriteAtomicJson(Path.Combine(runDirectory, "summary.json"), summary);
    }

    private static void WriteAtomicJson<T>(string path, T value)
    {
        if (File.Exists(path))
            throw new IOException($"Exploration artifact already exists and must not be overwritten: {path}");

        var json = SecretRedactor.Redact(JsonSerializer.Serialize(value, JsonOptions));
        WriteAtomicText(path, json);
    }

    private static void WriteAtomicJsonl<T>(string path, IReadOnlyList<T> values)
    {
        if (File.Exists(path))
            throw new IOException($"Exploration artifact already exists and must not be overwritten: {path}");

        var builder = new StringBuilder();
        foreach (var value in values)
        {
            var line = SecretRedactor.Redact(JsonSerializer.Serialize(value, JsonlOptions));
            builder.Append(line);
            builder.Append('\n');
        }

        WriteAtomicText(path, builder.ToString());
    }

    private static void WriteAtomicText(string path, string content)
    {
        var directory = Path.GetDirectoryName(path)
            ?? throw new IOException($"Unable to resolve directory for '{path}'.");
        Directory.CreateDirectory(directory);

        var tempPath = Path.Combine(directory, $".{Path.GetFileName(path)}.{Guid.NewGuid():N}.tmp");
        try
        {
            using (var stream = new FileStream(tempPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            using (var writer = new StreamWriter(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false)))
            {
                writer.Write(content);
                writer.Flush();
                stream.Flush(true);
            }

            File.Move(tempPath, path);
        }
        catch
        {
            try
            {
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
            catch
            {
                // best-effort cleanup
            }

            throw;
        }
    }
}
