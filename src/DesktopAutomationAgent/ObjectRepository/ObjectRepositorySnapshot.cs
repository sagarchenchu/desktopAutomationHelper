namespace DesktopAutomationAgent.ObjectRepository;

public sealed class ObjectRepositorySnapshot
{
    public ObjectRepositorySnapshot(
        ObjectRepositoryManifest manifest,
        IReadOnlyDictionary<string, PageObjectDocument> pages,
        IReadOnlyDictionary<string, string> pagePaths,
        IReadOnlyDictionary<string, string> fileHashes,
        string repositoryPath,
        string manifestSha256,
        string aggregateSha256)
    {
        Manifest = manifest;
        Pages = pages;
        PagePaths = pagePaths;
        FileHashes = fileHashes;
        RepositoryPath = repositoryPath;
        ManifestSha256 = manifestSha256;
        AggregateSha256 = aggregateSha256;
    }

    public ObjectRepositoryManifest Manifest { get; }

    public IReadOnlyDictionary<string, PageObjectDocument> Pages { get; }

    public IReadOnlyDictionary<string, string> PagePaths { get; }

    public IReadOnlyDictionary<string, string> FileHashes { get; }

    public string RepositoryPath { get; }

    public string ManifestSha256 { get; }

    public string AggregateSha256 { get; }
}
