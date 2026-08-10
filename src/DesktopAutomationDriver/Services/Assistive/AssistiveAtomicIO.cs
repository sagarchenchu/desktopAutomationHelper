using System.Text;

namespace DesktopAutomationDriver.Services.Assistive;

/// <summary>Atomic file helpers for Assistive primary recording and sidecar writes.</summary>
internal static class AssistiveAtomicIO
{
    private static readonly UTF8Encoding Utf8NoBom = new(encoderShouldEmitUTF8Identifier: false);

    /// <summary>
    /// Writes <paramref name="contents"/> via a temp file, then moves/replaces the destination.
    /// On failure the destination is left unchanged when it already existed.
    /// </summary>
    public static void ReplaceFileAtomic(string destinationPath, string contents)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(destinationPath);
        ArgumentNullException.ThrowIfNull(contents);

        var directory = Path.GetDirectoryName(destinationPath)
            ?? throw new InvalidOperationException("Destination directory is required.");
        Directory.CreateDirectory(directory);

        var tempPath = Path.Combine(directory, $".assistive-tmp-{Guid.NewGuid():N}.json");
        try
        {
            File.WriteAllText(tempPath, contents, Utf8NoBom);
            File.Move(tempPath, destinationPath, overwrite: true);
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
                // best-effort temp cleanup
            }

            throw;
        }
    }

    public static void TryDeleteDirectory(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return;

        try
        {
            if (Directory.Exists(path))
                Directory.Delete(path, recursive: true);
        }
        catch
        {
            // best-effort staging cleanup
        }
    }
}
