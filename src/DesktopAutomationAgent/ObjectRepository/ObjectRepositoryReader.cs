using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Plans;
using DesktopAutomationAgent.Workspace;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.ObjectRepository;

public sealed class ObjectRepositoryReader
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = false,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true
    };

    private readonly AgentOptions _options;
    private readonly IWorkspaceManager _workspace;

    public ObjectRepositoryReader(IOptions<AgentOptions> options, IWorkspaceManager workspace)
    {
        _options = options.Value;
        _workspace = workspace;
        AgentOptionsValidator.Validate(_options, OptionsValidationScope.ObjectRepository);
    }

    public ObjectRepositoryValidationResult Read(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return Failure(path ?? string.Empty, null, null, null, ["Object repository path is required."]);
        }

        string manifestFullPath;
        string manifestDisplayPath;
        try
        {
            manifestFullPath = _workspace.ResolveSafePath(path);
            manifestDisplayPath = ToDisplayRelativePath(manifestFullPath);
        }
        catch (WorkspaceException ex)
        {
            return Failure(path, null, null, null, [ex.Message]);
        }

        if (!File.Exists(manifestFullPath))
        {
            return Failure(
                manifestDisplayPath,
                null,
                null,
                null,
                [$"{manifestDisplayPath}: repository manifest not found."]);
        }

        try
        {
            ObjectRepositoryPathSafety.EnsureNotSymlinkEscape(manifestFullPath, _workspace.RootPath);
        }
        catch (RepositoryPathException ex)
        {
            return Failure(manifestDisplayPath, null, null, null, [$"{manifestDisplayPath}: {ex.Message}"]);
        }

        byte[] manifestBytes;
        try
        {
            var fileInfo = new FileInfo(manifestFullPath);
            if (fileInfo.Length > _options.ObjectRepository.MaxFileBytes)
            {
                return Failure(
                    manifestDisplayPath,
                    null,
                    null,
                    null,
                    [
                        $"{manifestDisplayPath}: repository manifest exceeds maximum size of {_options.ObjectRepository.MaxFileBytes} bytes."
                    ]);
            }

            manifestBytes = File.ReadAllBytes(manifestFullPath);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return Failure(
                manifestDisplayPath,
                null,
                null,
                null,
                [$"{manifestDisplayPath}: failed to read repository manifest ({ex.Message})."]);
        }

        var manifestSha256 = ComputeSha256(manifestBytes);
        var fileHashes = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            [manifestDisplayPath] = manifestSha256
        };

        try
        {
            var duplicateErrors = JsonDuplicatePropertyDetector.DetectDuplicates(manifestBytes)
                .Select(message => $"{manifestDisplayPath}: {message}")
                .ToList();
            if (duplicateErrors.Count > 0)
            {
                return Failure(manifestDisplayPath, manifestSha256, fileHashes, null, duplicateErrors);
            }
        }
        catch (Exception ex) when (ex is JsonException or ArgumentException or InvalidOperationException)
        {
            return Failure(
                manifestDisplayPath,
                manifestSha256,
                fileHashes,
                null,
                [$"{manifestDisplayPath}: invalid JSON ({ex.Message})."]);
        }

        var nullErrors = ObjectRepositoryNullRejector.Detect(manifestBytes, manifestDisplayPath);
        if (nullErrors.Count > 0)
        {
            return Failure(manifestDisplayPath, manifestSha256, fileHashes, null, nullErrors);
        }

        ObjectRepositoryManifest? manifest;
        try
        {
            manifest = JsonSerializer.Deserialize<ObjectRepositoryManifest>(manifestBytes, JsonOptions);
        }
        catch (JsonException ex)
        {
            return Failure(
                manifestDisplayPath,
                manifestSha256,
                fileHashes,
                null,
                [$"{manifestDisplayPath}: invalid JSON ({ex.Message})."]);
        }
        catch (NotSupportedException ex)
        {
            return Failure(
                manifestDisplayPath,
                manifestSha256,
                fileHashes,
                null,
                [$"{manifestDisplayPath}: invalid JSON ({ex.Message})."]);
        }

        if (manifest is null)
        {
            return Failure(
                manifestDisplayPath,
                manifestSha256,
                fileHashes,
                null,
                [$"{manifestDisplayPath}: repository manifest was empty."]);
        }

        var repositoryRoot = Path.GetDirectoryName(manifestFullPath)
            ?? throw new InvalidOperationException("Repository manifest path has no directory.");

        var pages = new Dictionary<string, PageObjectDocument>(StringComparer.Ordinal);
        var pagePaths = new Dictionary<string, string>(StringComparer.Ordinal);
        var errors = new List<string>();

        if (manifest.Pages is not null)
        {
            for (var pageIndex = 0; pageIndex < manifest.Pages.Count; pageIndex++)
            {
                var reference = manifest.Pages[pageIndex];
                if (reference is null)
                {
                    errors.Add($"{manifestDisplayPath}: pages[{pageIndex}] must not be null.");
                    continue;
                }

                if (string.IsNullOrWhiteSpace(reference.PageId) || string.IsNullOrWhiteSpace(reference.File))
                    continue;

                var pageLocation = $"{manifestDisplayPath}: pages entry '{reference.PageId}'";
                string pageFullPath;
                string pageDisplayPath;
                try
                {
                    pageFullPath = ResolveRepositoryPagePath(repositoryRoot, reference.File);
                    ObjectRepositoryPathSafety.EnsureNotSymlinkEscape(pageFullPath, repositoryRoot);
                    ObjectRepositoryPathSafety.EnsureNotSymlinkEscape(pageFullPath, _workspace.RootPath);
                    pageDisplayPath = ToDisplayRelativePath(pageFullPath);
                }
                catch (RepositoryPathException ex)
                {
                    errors.Add($"{pageLocation}: {ex.Message}");
                    continue;
                }

                if (!File.Exists(pageFullPath))
                {
                    errors.Add($"{pageDisplayPath}: page file not found.");
                    continue;
                }

                byte[] pageBytes;
                try
                {
                    var pageInfo = new FileInfo(pageFullPath);
                    if (pageInfo.Length > _options.ObjectRepository.MaxFileBytes)
                    {
                        errors.Add(
                            $"{pageDisplayPath}: page file exceeds maximum size of {_options.ObjectRepository.MaxFileBytes} bytes.");
                        continue;
                    }

                    pageBytes = File.ReadAllBytes(pageFullPath);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    errors.Add($"{pageDisplayPath}: failed to read page file ({ex.Message}).");
                    continue;
                }

                var pageSha256 = ComputeSha256(pageBytes);
                fileHashes[pageDisplayPath] = pageSha256;

                try
                {
                    var duplicateErrors = JsonDuplicatePropertyDetector.DetectDuplicates(pageBytes)
                        .Select(message => $"{pageDisplayPath}: {message}")
                        .ToList();
                    if (duplicateErrors.Count > 0)
                    {
                        errors.AddRange(duplicateErrors);
                        continue;
                    }
                }
                catch (Exception ex) when (ex is JsonException or ArgumentException or InvalidOperationException)
                {
                    errors.Add($"{pageDisplayPath}: invalid JSON ({ex.Message}).");
                    continue;
                }

                var pageNullErrors = ObjectRepositoryNullRejector.Detect(pageBytes, pageDisplayPath);
                if (pageNullErrors.Count > 0)
                {
                    errors.AddRange(pageNullErrors);
                    continue;
                }

                PageObjectDocument? page;
                try
                {
                    page = JsonSerializer.Deserialize<PageObjectDocument>(pageBytes, JsonOptions);
                }
                catch (JsonException ex)
                {
                    errors.Add($"{pageDisplayPath}: invalid JSON ({ex.Message}).");
                    continue;
                }
                catch (NotSupportedException ex)
                {
                    errors.Add($"{pageDisplayPath}: invalid JSON ({ex.Message}).");
                    continue;
                }

                if (page is null)
                {
                    errors.Add($"{pageDisplayPath}: page document was empty.");
                    continue;
                }

                if (!string.Equals(page.PageId, reference.PageId, StringComparison.Ordinal))
                {
                    errors.Add(
                        $"{pageDisplayPath}: pageId '{page.PageId}' does not match manifest pageId '{reference.PageId}'.");
                }

                if (pages.ContainsKey(reference.PageId))
                {
                    errors.Add($"{pageDisplayPath}: duplicate loaded pageId '{reference.PageId}'.");
                    continue;
                }

                pages[reference.PageId] = page;
                pagePaths[reference.PageId] = pageDisplayPath;
            }
        }

        var aggregateSha256 = ComputeAggregateSha256(fileHashes);
        var validator = new ObjectRepositoryValidator();
        var validation = validator.Validate(
            manifest,
            pages,
            pagePaths,
            manifestDisplayPath,
            _options.ObjectRepository);

        errors.AddRange(validation.Errors);
        var warnings = validation.Warnings.ToList();

        if (errors.Count > 0)
        {
            return new ObjectRepositoryValidationResult
            {
                RepositoryPath = manifestDisplayPath,
                Errors = errors,
                Warnings = warnings,
                ManifestSha256 = manifestSha256,
                FileHashes = fileHashes,
                AggregateSha256 = aggregateSha256
            };
        }

        var snapshot = new ObjectRepositorySnapshot(
            manifest,
            pages,
            pagePaths,
            fileHashes,
            manifestDisplayPath,
            manifestSha256,
            aggregateSha256);

        return new ObjectRepositoryValidationResult
        {
            RepositoryPath = manifestDisplayPath,
            Errors = errors,
            Warnings = warnings,
            Snapshot = snapshot,
            ManifestSha256 = manifestSha256,
            FileHashes = fileHashes,
            AggregateSha256 = aggregateSha256
        };
    }

    internal static string ResolveRepositoryPagePath(string repositoryRoot, string pageFile)
    {
        if (string.IsNullOrWhiteSpace(pageFile))
            throw new RepositoryPathException("file is required.");

        if (Path.IsPathRooted(pageFile))
            throw new RepositoryPathException("file must be a repository-relative path.");

        if (pageFile.Contains("..", StringComparison.Ordinal))
            throw new RepositoryPathException("file must not contain '..'.");

        var normalizedRoot = Path.GetFullPath(repositoryRoot)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var candidate = Path.GetFullPath(Path.Combine(normalizedRoot, pageFile));

        if (!IsInsideDirectory(candidate, normalizedRoot))
            throw new RepositoryPathException("file resolves outside the repository directory.");

        var relative = Path.GetRelativePath(normalizedRoot, candidate).Replace('\\', '/');
        if (!relative.StartsWith("pages/", StringComparison.Ordinal)
            && !string.Equals(relative, "pages", StringComparison.Ordinal))
        {
            throw new RepositoryPathException("file must be under the pages/ subdirectory.");
        }

        return candidate;
    }

    private static bool IsInsideDirectory(string fullPath, string directoryRoot)
    {
        var root = Path.GetFullPath(directoryRoot)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var normalized = Path.GetFullPath(fullPath)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

        if (WorkspaceManager.PathsEqual(normalized, root))
            return true;

        var relative = Path.GetRelativePath(root, normalized);
        return !Path.IsPathRooted(relative)
               && relative != ".."
               && !relative.StartsWith(".." + Path.DirectorySeparatorChar, StringComparison.Ordinal)
               && !relative.StartsWith(".." + Path.AltDirectorySeparatorChar, StringComparison.Ordinal);
    }

    private string ToDisplayRelativePath(string fullPath)
    {
        var root = Path.GetFullPath(_workspace.RootPath)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var normalized = Path.GetFullPath(fullPath)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

        if (WorkspaceManager.PathsEqual(normalized, root))
            return ".";

        var relative = Path.GetRelativePath(root, normalized);
        return relative.Replace(Path.DirectorySeparatorChar, '/');
    }

    private static string ComputeSha256(byte[] bytes) =>
        Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    internal static string ComputeAggregateSha256(IReadOnlyDictionary<string, string> fileHashes)
    {
        var builder = new StringBuilder();
        foreach (var entry in fileHashes.OrderBy(static pair => pair.Key, StringComparer.Ordinal))
        {
            builder.Append(entry.Key);
            builder.Append(':');
            builder.Append(entry.Value);
            builder.Append('\n');
        }

        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(builder.ToString())))
            .ToLowerInvariant();
    }

    private static ObjectRepositoryValidationResult Failure(
        string repositoryPath,
        string? manifestSha256,
        IReadOnlyDictionary<string, string>? fileHashes,
        string? aggregateSha256,
        IReadOnlyList<string> errors) =>
        new()
        {
            RepositoryPath = repositoryPath,
            ManifestSha256 = manifestSha256,
            FileHashes = fileHashes,
            AggregateSha256 = aggregateSha256,
            Errors = errors
        };
}

public sealed class RepositoryPathException : Exception
{
    public RepositoryPathException(string message) : base(message)
    {
    }
}
