using System.Text.Json;
using System.Text.RegularExpressions;
using DesktopAutomationAgent.Configuration;

namespace DesktopAutomationAgent.ObjectRepository;

public sealed class ObjectRepositoryValidator
{
    public const int SchemaVersion = 1;

    private static readonly Regex IdentifierPattern = new(
        @"^[a-z][a-z0-9-]{0,63}$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly HashSet<string> AllowedPageStates = new(StringComparer.Ordinal)
    {
        "candidate",
        "active"
    };

    private static readonly HashSet<string> AllowedSourceKinds = new(StringComparer.Ordinal)
    {
        "capture",
        "manual",
        "approved"
    };

    private static readonly HashSet<string> AllowedQualityGrades = new(StringComparer.Ordinal)
    {
        "strong",
        "medium",
        "weak"
    };

    public ObjectRepositoryValidationResult Validate(
        ObjectRepositoryManifest manifest,
        IReadOnlyDictionary<string, PageObjectDocument> pages,
        IReadOnlyDictionary<string, string> pageFilePaths,
        string repositoryPath,
        ObjectRepositoryOptions options)
    {
        ArgumentNullException.ThrowIfNull(manifest);
        ArgumentNullException.ThrowIfNull(pages);
        ArgumentNullException.ThrowIfNull(pageFilePaths);
        ArgumentNullException.ThrowIfNull(options);

        var errors = new List<string>();
        var warnings = new List<string>();

        ValidateManifest(manifest, repositoryPath, options, errors);

        var totalElements = 0;
        var seenPageIds = new Dictionary<string, string>(StringComparer.Ordinal);
        var seenPageFiles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        if (manifest.Pages is null)
            return BuildResult(errors, warnings, repositoryPath);

        for (var i = 0; i < manifest.Pages.Count; i++)
        {
            var reference = manifest.Pages[i];
            var location = $"{repositoryPath}: pages[{i}]";

            if (reference is null)
            {
                errors.Add($"{location}: page reference must not be null.");
                continue;
            }

            if (reference.ExtensionData is { Count: > 0 })
            {
                foreach (var key in reference.ExtensionData.Keys)
                    errors.Add($"{location}: unknown property '{key}'.");
            }

            if (string.IsNullOrWhiteSpace(reference.PageId) || !IdentifierPattern.IsMatch(reference.PageId))
            {
                errors.Add($"{location}: pageId must match ^[a-z][a-z0-9-]{{0,63}}$.");
                continue;
            }

            if (seenPageIds.TryGetValue(reference.PageId, out var firstLocation))
            {
                errors.Add($"{location}: pageId '{reference.PageId}' duplicates {firstLocation}.");
            }
            else
            {
                seenPageIds[reference.PageId] = location;
            }

            if (string.IsNullOrWhiteSpace(reference.File))
            {
                errors.Add($"{location}: file is required.");
                continue;
            }

            if (Path.IsPathRooted(reference.File) || reference.File.Contains("..", StringComparison.Ordinal))
            {
                errors.Add($"{location}: file must be a repository-relative path without '..'.");
                continue;
            }

            var normalizedFile = reference.File.Replace('\\', '/');
            if (!normalizedFile.StartsWith("pages/", StringComparison.Ordinal))
            {
                errors.Add($"{location}: file must be under the pages/ subdirectory.");
            }

            if (seenPageFiles.TryGetValue(normalizedFile, out var firstFileLocation))
            {
                errors.Add($"{location}: file '{reference.File}' duplicates {firstFileLocation}.");
            }
            else
            {
                seenPageFiles[normalizedFile] = location;
            }

            var preferredFile = $"pages/{reference.PageId}.page.json";
            if (!string.Equals(normalizedFile, preferredFile, StringComparison.Ordinal)
                && normalizedFile.StartsWith("pages/", StringComparison.Ordinal))
            {
                warnings.Add(
                    $"{location}: file '{reference.File}' does not follow preferred naming '{preferredFile}'.");
            }

            if (!pages.TryGetValue(reference.PageId, out var page))
            {
                errors.Add($"{location}: page file for '{reference.PageId}' was not loaded.");
                continue;
            }

            if (!pageFilePaths.TryGetValue(reference.PageId, out var pagePath))
            {
                errors.Add($"{location}: page file path for '{reference.PageId}' was not resolved.");
                continue;
            }

            ValidatePage(page, pagePath, reference.PageId, manifestReferenced: true, options, errors, warnings, ref totalElements);
        }

        if (totalElements > options.MaxTotalElements)
        {
            errors.Add(
                $"{repositoryPath}: total element count {totalElements} exceeds maximum {options.MaxTotalElements}.");
        }

        return BuildResult(errors, warnings, repositoryPath);
    }

    private static void ValidateManifest(
        ObjectRepositoryManifest manifest,
        string repositoryPath,
        ObjectRepositoryOptions options,
        List<string> errors)
    {
        if (manifest.SchemaVersion != SchemaVersion)
        {
            errors.Add(
                $"{repositoryPath}: unsupported schemaVersion {manifest.SchemaVersion}; expected {SchemaVersion}.");
        }

        if (string.IsNullOrWhiteSpace(manifest.RepositoryId) || !IdentifierPattern.IsMatch(manifest.RepositoryId))
        {
            errors.Add($"{repositoryPath}: repositoryId must match ^[a-z][a-z0-9-]{{0,63}}$.");
        }

        if (string.IsNullOrWhiteSpace(manifest.Name))
        {
            errors.Add($"{repositoryPath}: name is required.");
        }

        if (manifest.ExtensionData is { Count: > 0 })
        {
            foreach (var key in manifest.ExtensionData.Keys)
            {
                errors.Add($"{repositoryPath}: unknown top-level property '{key}'.");
            }
        }

        if (manifest.Pages is null)
        {
            errors.Add($"{repositoryPath}: pages is required.");
            return;
        }

        if (manifest.Pages.Count > options.MaxPages)
        {
            errors.Add(
                $"{repositoryPath}: page count {manifest.Pages.Count} exceeds maximum {options.MaxPages}.");
        }
    }

    private static void ValidatePage(
        PageObjectDocument page,
        string pagePath,
        string expectedPageId,
        bool manifestReferenced,
        ObjectRepositoryOptions options,
        List<string> errors,
        List<string> warnings,
        ref int totalElements)
    {
        if (page.SchemaVersion != SchemaVersion)
        {
            errors.Add($"{pagePath}: unsupported schemaVersion {page.SchemaVersion}; expected {SchemaVersion}.");
        }

        if (string.IsNullOrWhiteSpace(page.PageId) || !IdentifierPattern.IsMatch(page.PageId))
        {
            errors.Add($"{pagePath}: pageId must match ^[a-z][a-z0-9-]{{0,63}}$.");
        }
        else if (!string.Equals(page.PageId, expectedPageId, StringComparison.Ordinal))
        {
            errors.Add($"{pagePath}: pageId '{page.PageId}' does not match manifest pageId '{expectedPageId}'.");
        }

        if (string.IsNullOrWhiteSpace(page.Name))
        {
            errors.Add($"{pagePath}: name is required.");
        }

        if (string.IsNullOrWhiteSpace(page.State) || !AllowedPageStates.Contains(page.State))
        {
            errors.Add($"{pagePath}: state must be 'candidate' or 'active'.");
        }
        else if (manifestReferenced && string.Equals(page.State, "candidate", StringComparison.Ordinal))
        {
            errors.Add($"{pagePath}: manifest-referenced pages must have state 'active'.");
        }

        if (string.Equals(page.State, "active", StringComparison.Ordinal) && HasNonEmptyUnresolved(page.Unresolved))
        {
            errors.Add($"{pagePath}: active pages must not contain a non-empty unresolved section.");
        }

        if (page.ExtensionData is { Count: > 0 })
        {
            foreach (var key in page.ExtensionData.Keys)
            {
                errors.Add($"{pagePath}: unknown top-level property '{key}'.");
            }
        }

        if (page.Elements is null)
        {
            errors.Add($"{pagePath}: elements is required.");
            return;
        }

        if (page.Elements.Count > options.MaxElementsPerPage)
        {
            errors.Add(
                $"{pagePath}: element count {page.Elements.Count} exceeds maximum {options.MaxElementsPerPage}.");
        }

        totalElements += page.Elements.Count;

        var seenElementIds = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var (elementId, element) in page.Elements)
        {
            var location = $"{pagePath}: elements['{elementId}']";

            if (string.IsNullOrWhiteSpace(elementId) || !IdentifierPattern.IsMatch(elementId))
            {
                errors.Add($"{location}: element id must match ^[a-z][a-z0-9-]{{0,63}}$.");
                continue;
            }

            if (seenElementIds.TryGetValue(elementId, out var firstElementLocation))
            {
                errors.Add($"{location}: element id '{elementId}' duplicates {firstElementLocation}.");
            }
            else
            {
                seenElementIds[elementId] = location;
            }

            if (element is null)
            {
                errors.Add($"{location}: element definition must not be null.");
                continue;
            }

            ValidateElement(element, location, page.State, errors, warnings);
        }
    }

    public ObjectRepositoryValidationResult ValidateCandidatePage(
        PageObjectDocument page,
        string pagePath,
        ObjectRepositoryOptions options)
    {
        ArgumentNullException.ThrowIfNull(page);
        ArgumentNullException.ThrowIfNull(options);

        var errors = new List<string>();
        var warnings = new List<string>();
        var totalElements = 0;
        ValidatePage(
            page,
            pagePath,
            page.PageId,
            manifestReferenced: false,
            options,
            errors,
            warnings,
            ref totalElements);

        return BuildResult(errors, warnings, pagePath);
    }

    private static bool HasNonEmptyUnresolved(JsonElement? unresolved)
    {
        if (unresolved is not JsonElement element)
            return false;

        return element.ValueKind switch
        {
            JsonValueKind.Array => element.GetArrayLength() > 0,
            JsonValueKind.Object => element.EnumerateObject().Any(),
            JsonValueKind.Null or JsonValueKind.Undefined => false,
            _ => true
        };
    }

    private static void ValidateElement(
        ObjectElementDefinition element,
        string location,
        string pageState,
        List<string> errors,
        List<string> warnings)
    {
        if (element.ExtensionData is { Count: > 0 })
        {
            foreach (var key in element.ExtensionData.Keys)
            {
                errors.Add($"{location}: unknown property '{key}'.");
            }
        }

        if (element.Locator is null)
        {
            errors.Add($"{location}: locator is required.");
        }
        else
        {
            var locatorValidation = ObjectLocatorValidator.Validate(element.Locator, $"{location}.locator");
            errors.AddRange(locatorValidation.Errors);
            warnings.AddRange(locatorValidation.Warnings);
        }

        if (element.Quality is not null)
        {
            if (element.Quality.ExtensionData is { Count: > 0 })
            {
                foreach (var key in element.Quality.ExtensionData.Keys)
                    errors.Add($"{location}.quality: unknown property '{key}'.");
            }

            if (!string.IsNullOrWhiteSpace(element.Quality.Grade)
                && !AllowedQualityGrades.Contains(element.Quality.Grade))
            {
                errors.Add($"{location}: quality.grade must be strong, medium, or weak.");
            }
        }

        if (element.Source is not null)
        {
            if (element.Source.ExtensionData is { Count: > 0 })
            {
                foreach (var key in element.Source.ExtensionData.Keys)
                    errors.Add($"{location}.source: unknown property '{key}'.");
            }

            if (string.IsNullOrWhiteSpace(element.Source.Kind)
                || !AllowedSourceKinds.Contains(element.Source.Kind))
            {
                errors.Add($"{location}: source.kind must be capture, manual, or approved.");
            }

            if (string.Equals(element.Source.Kind, "capture", StringComparison.Ordinal)
                && string.Equals(pageState, "active", StringComparison.Ordinal))
            {
                errors.Add($"{location}: active pages must not contain capture-sourced elements; promote manually.");
            }
        }
    }

    private static ObjectRepositoryValidationResult BuildResult(
        List<string> errors,
        List<string> warnings,
        string repositoryPath) =>
        new()
        {
            RepositoryPath = repositoryPath,
            Errors = errors,
            Warnings = warnings
        };
}
