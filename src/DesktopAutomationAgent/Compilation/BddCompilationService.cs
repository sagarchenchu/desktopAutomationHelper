using System.Text.Json;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver.Models;
using DesktopAutomationAgent.Knowledge;
using DesktopAutomationAgent.ObjectRepository;
using DesktopAutomationAgent.Plans;
using DesktopAutomationAgent.Workspace;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Compilation;

public sealed class BddCompilationService
{
    private readonly IWorkspaceManager _workspace;
    private readonly ObjectRepositoryReader _objectRepositoryReader;
    private readonly JiraTestCaseValidator _testCaseValidator = new();
    private readonly StepRegistryValidator _registryValidator = new();
    private readonly DeterministicBddCompiler _compiler = new();
    private readonly CompilationCache _cache;
    private readonly PlanValidator _planValidator = new();
    private readonly KnowledgeValidator _knowledgeValidator;

    public BddCompilationService(
        IWorkspaceManager workspace,
        ObjectRepositoryReader objectRepositoryReader,
        IOptions<AgentOptions>? options = null)
    {
        _workspace = workspace;
        _objectRepositoryReader = objectRepositoryReader;
        _cache = new CompilationCache(workspace.RootPath);
        var schemaDir = Path.Combine(workspace.RootPath, "schemas");
        _knowledgeValidator = new KnowledgeValidator(Directory.Exists(schemaDir) ? schemaDir : null);
        _ = options; // reserved for future size limits
    }

    public JiraTestCaseValidationResult ValidateTestCase(string path)
    {
        var full = _workspace.ResolveSafePath(path);
        return _testCaseValidator.ValidateFile(full, ToRelative(full));
    }

    public StepRegistryValidationResult ValidateStepRegistry(string? path = null)
    {
        var relative = string.IsNullOrWhiteSpace(path) ? "step-registry/registry.json" : path;
        var full = _workspace.ResolveSafePath(relative);
        return _registryValidator.ValidateFile(full, ToRelative(full));
    }

    public CompilationResult CompileKey(CompileKeyRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);

        if (!JiraKeyContract.IsCanonical(request.JiraKey))
        {
            return new CompilationResult
            {
                Success = false,
                ExitCode = 2,
                Errors = [$"invalid jiraKey '{request.JiraKey}'."]
            };
        }

        var testCasePath = ResolveTestCasePath(request.JiraKey, request.TestCasePath);
        var testCaseResult = _testCaseValidator.ValidateFile(testCasePath, ToRelative(testCasePath));
        if (!testCaseResult.IsValid || testCaseResult.Document is null || testCaseResult.Sha256 is null)
        {
            return new CompilationResult
            {
                Success = false,
                ExitCode = 5,
                JiraKey = request.JiraKey,
                Errors = testCaseResult.Errors
            };
        }

        if (!string.Equals(testCaseResult.Document.JiraKey, request.JiraKey, StringComparison.Ordinal))
        {
            return new CompilationResult
            {
                Success = false,
                ExitCode = 5,
                JiraKey = request.JiraKey,
                Errors =
                [
                    $"test-case jiraKey '{testCaseResult.Document.JiraKey}' does not match requested key '{request.JiraKey}'."
                ]
            };
        }

        var registryRelative = string.IsNullOrWhiteSpace(request.RegistryPath)
            ? "step-registry/registry.json"
            : request.RegistryPath!;
        var registryFull = _workspace.ResolveSafePath(registryRelative);
        var registryResult = _registryValidator.ValidateFile(registryFull, ToRelative(registryFull));
        if (!registryResult.IsValid || registryResult.Document is null || registryResult.Sha256 is null)
        {
            return new CompilationResult
            {
                Success = false,
                ExitCode = 5,
                JiraKey = request.JiraKey,
                Errors = registryResult.Errors
            };
        }

        var objectRepoRelative = string.IsNullOrWhiteSpace(request.ObjectRepositoryPath)
            ? "object-repository/repository.json"
            : request.ObjectRepositoryPath!;
        var objectRepoResult = _objectRepositoryReader.Read(objectRepoRelative);
        if (!objectRepoResult.IsValid || objectRepoResult.Snapshot is null)
        {
            return new CompilationResult
            {
                Success = false,
                ExitCode = 5,
                JiraKey = request.JiraKey,
                Errors = objectRepoResult.Errors.ToList()
            };
        }

        var objectRepoSha = objectRepoResult.AggregateSha256
            ?? objectRepoResult.ManifestSha256
            ?? throw new InvalidOperationException("Object repository aggregate hash missing.");

        ApplicationKnowledgeDocument? knowledge = null;
        string? knowledgeSha = null;
        if (!string.IsNullOrWhiteSpace(request.KnowledgePath))
        {
            try
            {
                var resolved = _workspace.ResolveSafePath(request.KnowledgePath);
                var knowledgeDir = File.Exists(resolved)
                                   && resolved.EndsWith("application-knowledge.json", StringComparison.OrdinalIgnoreCase)
                    ? Path.GetDirectoryName(resolved)!
                    : resolved;

                // Containment under knowledge/approved — rejects knowledge/approved-rogue via path rules.
                KnowledgePathSafety.EnsureReadablePathInside(
                    knowledgeDir,
                    KnowledgePathSafety.ApprovedRoot(_workspace.RootPath));

                var validation = _knowledgeValidator.ValidateDirectory(knowledgeDir);
                if (!validation.IsValid)
                {
                    return new CompilationResult
                    {
                        Success = false,
                        ExitCode = 5,
                        JiraKey = request.JiraKey,
                        Errors = validation.Errors
                            .Select(e => $"knowledge-validation: {e}")
                            .ToList()
                    };
                }

                var knowledgeFile = Path.Combine(knowledgeDir, "application-knowledge.json");
                knowledge = JsonSerializer.Deserialize<ApplicationKnowledgeDocument>(
                    File.ReadAllBytes(knowledgeFile),
                    CompilationJson.DocumentRead);
                if (knowledge is null || !string.Equals(knowledge.Status, "approved", StringComparison.Ordinal))
                {
                    return new CompilationResult
                    {
                        Success = false,
                        ExitCode = 5,
                        JiraKey = request.JiraKey,
                        Errors = ["knowledge document must have status=approved."]
                    };
                }

                knowledgeSha = CompilationFingerprint.ComputeFileSha256(knowledgeFile);
            }
            catch (KnowledgePathException ex)
            {
                return new CompilationResult
                {
                    Success = false,
                    ExitCode = 5,
                    JiraKey = request.JiraKey,
                    Errors = [$"knowledge path must resolve under knowledge/approved: {ex.Message}"]
                };
            }
            catch (WorkspaceException ex)
            {
                return new CompilationResult
                {
                    Success = false,
                    ExitCode = 5,
                    JiraKey = request.JiraKey,
                    Errors = [ex.Message]
                };
            }
        }

        OperationsCatalogDto? catalog = null;
        string? catalogFingerprint = null;
        if (!string.IsNullOrWhiteSpace(request.CatalogPath))
        {
            var catalogFull = _workspace.ResolveSafePath(request.CatalogPath);
            catalog = JsonSerializer.Deserialize<OperationsCatalogDto>(
                File.ReadAllBytes(catalogFull),
                CompilationJson.DocumentRead);
            if (catalog is null)
            {
                return new CompilationResult
                {
                    Success = false,
                    ExitCode = 5,
                    JiraKey = request.JiraKey,
                    Errors = ["failed to deserialize operations catalog."]
                };
            }

            catalogFingerprint = CompilationFingerprint.ComputeCatalogFingerprint(catalog);
        }

        var compileResult = _compiler.Compile(new DeterministicCompileRequest
        {
            TestCase = testCaseResult.Document,
            InputSha256 = testCaseResult.Sha256,
            Registry = registryResult.Document,
            StepRegistrySha256 = registryResult.Sha256,
            ObjectRepository = objectRepoResult.Snapshot,
            ObjectRepositoryRelativePath = CompilationFingerprint.CanonicalizeWorkspaceRelativePath(objectRepoRelative),
            ObjectRepositorySha256 = objectRepoSha,
            ApprovedKnowledge = knowledge,
            ApprovedKnowledgeSha256 = knowledgeSha,
            Catalog = catalog,
            CatalogFingerprint = catalogFingerprint,
            ParameterBindings = request.ParameterBindings,
            Explain = request.Explain,
            CompiledAtUtc = request.CompiledAtUtc
        });

        if (!compileResult.Success || compileResult.Fingerprint is null || compileResult.Metadata is null || compileResult.Plan is null)
            return compileResult;

        // Cache reuse on full fingerprint match.
        if (_cache.TryReadExisting(
                request.JiraKey,
                compileResult.Fingerprint,
                out var existingPlanBytes,
                out var existingMeta,
                out var cacheErrors))
        {
            var freshPlanBytes = CompilationJson.CanonicalSerializeUtf8(compileResult.Plan);
            if (freshPlanBytes.Length == 0 || freshPlanBytes[^1] != (byte)'\n')
            {
                var withNl = new byte[freshPlanBytes.Length + 1];
                Buffer.BlockCopy(freshPlanBytes, 0, withNl, 0, freshPlanBytes.Length);
                withNl[^1] = (byte)'\n';
                freshPlanBytes = withNl;
            }

            if (existingPlanBytes is not null && existingPlanBytes.AsSpan().SequenceEqual(freshPlanBytes))
            {
                var dir = _cache.GetArtifactDirectory(request.JiraKey, compileResult.Fingerprint);
                existingMeta!.CacheHit = true;
                return new CompilationResult
                {
                    Success = true,
                    ExitCode = 0,
                    JiraKey = request.JiraKey,
                    Fingerprint = compileResult.Fingerprint,
                    OutputDirectory = ToRelative(dir),
                    PlanPath = ToRelative(Path.Combine(dir, "plan.json")),
                    CompilationPath = ToRelative(Path.Combine(dir, "compilation.json")),
                    PlanSha256 = compileResult.PlanSha256,
                    CacheHit = true,
                    Plan = compileResult.Plan,
                    Metadata = existingMeta,
                    Warnings = compileResult.Warnings,
                    Explain = compileResult.Explain
                };
            }

            return new CompilationResult
            {
                Success = false,
                ExitCode = 5,
                JiraKey = request.JiraKey,
                Fingerprint = compileResult.Fingerprint,
                Errors =
                [
                    $"cache collision: existing plan bytes differ for fingerprint '{compileResult.Fingerprint}'.",
                    ..cacheErrors
                ]
            };
        }

        string planPath;
        string compilationPath;
        try
        {
            var planBytes = CompilationJson.CanonicalSerializeUtf8(compileResult.Plan);
            if (planBytes.Length == 0 || planBytes[^1] != (byte)'\n')
            {
                var withNl = new byte[planBytes.Length + 1];
                Buffer.BlockCopy(planBytes, 0, withNl, 0, planBytes.Length);
                withNl[^1] = (byte)'\n';
                planBytes = withNl;
            }

            compileResult.Metadata.CacheHit = false;
            _cache.WriteAtomic(
                request.JiraKey,
                compileResult.Fingerprint,
                planBytes,
                compileResult.Metadata,
                out planPath,
                out compilationPath);
        }
        catch (IOException ex)
        {
            return new CompilationResult
            {
                Success = false,
                ExitCode = 5,
                JiraKey = request.JiraKey,
                Fingerprint = compileResult.Fingerprint,
                Errors = [ex.Message]
            };
        }

        return new CompilationResult
        {
            Success = true,
            ExitCode = 0,
            JiraKey = request.JiraKey,
            Fingerprint = compileResult.Fingerprint,
            OutputDirectory = ToRelative(_cache.GetArtifactDirectory(request.JiraKey, compileResult.Fingerprint)),
            PlanPath = ToRelative(planPath),
            CompilationPath = ToRelative(compilationPath),
            PlanSha256 = compileResult.PlanSha256,
            CacheHit = false,
            Plan = compileResult.Plan,
            Metadata = compileResult.Metadata,
            Warnings = compileResult.Warnings,
            Explain = compileResult.Explain
        };
    }

    public CompiledArtifactValidationResult ValidateCompiledKey(string jiraKey, string? fingerprint = null)
    {
        if (!JiraKeyContract.IsCanonical(jiraKey))
        {
            return new CompiledArtifactValidationResult
            {
                IsValid = false,
                JiraKey = jiraKey,
                Errors = [$"invalid jiraKey '{jiraKey}'."]
            };
        }

        var keyRoot = Path.Combine(_workspace.RootPath, "compiled", jiraKey);
        if (!Directory.Exists(keyRoot))
        {
            return new CompiledArtifactValidationResult
            {
                IsValid = false,
                JiraKey = jiraKey,
                Errors = [$"no compiled artifacts for '{jiraKey}'."]
            };
        }

        string artifactDir;
        if (!string.IsNullOrWhiteSpace(fingerprint))
        {
            if (!CompilationFingerprint.IsHexFingerprint(fingerprint))
            {
                return new CompiledArtifactValidationResult
                {
                    IsValid = false,
                    JiraKey = jiraKey,
                    Fingerprint = fingerprint,
                    Errors = ["fingerprint must be a 64-character lowercase hex SHA-256 (path traversal rejected)."]
                };
            }

            artifactDir = Path.Combine(keyRoot, fingerprint);
            try
            {
                KnowledgePathSafety.EnsureReadablePathInside(artifactDir, keyRoot);
            }
            catch (KnowledgePathException ex)
            {
                return new CompiledArtifactValidationResult
                {
                    IsValid = false,
                    JiraKey = jiraKey,
                    Fingerprint = fingerprint,
                    Errors = [$"fingerprint path containment failed: {ex.Message}"]
                };
            }

            if (!Directory.Exists(artifactDir))
            {
                return new CompiledArtifactValidationResult
                {
                    IsValid = false,
                    JiraKey = jiraKey,
                    Fingerprint = fingerprint,
                    Errors = [$"compiled fingerprint '{fingerprint}' not found for '{jiraKey}'."]
                };
            }
        }
        else
        {
            var dirs = Directory.GetDirectories(keyRoot)
                .Where(d => CompilationFingerprint.IsHexFingerprint(Path.GetFileName(d)))
                .OrderBy(d => d, StringComparer.Ordinal)
                .ToArray();
            if (dirs.Length == 0)
            {
                return new CompiledArtifactValidationResult
                {
                    IsValid = false,
                    JiraKey = jiraKey,
                    Errors = [$"no compiled fingerprints for '{jiraKey}'."]
                };
            }

            if (dirs.Length > 1)
            {
                return new CompiledArtifactValidationResult
                {
                    IsValid = false,
                    JiraKey = jiraKey,
                    Errors =
                    [
                        $"multiple compiled fingerprints for '{jiraKey}'; pass --fingerprint <sha>."
                    ]
                };
            }

            artifactDir = dirs[0];
            fingerprint = Path.GetFileName(artifactDir);
        }

        if (!_cache.TryReadExisting(jiraKey, fingerprint!, out var planBytes, out var metadata, out var cacheErrors)
            || planBytes is null
            || metadata is null)
        {
            return new CompiledArtifactValidationResult
            {
                IsValid = false,
                JiraKey = jiraKey,
                Fingerprint = fingerprint,
                Errors = cacheErrors.Count > 0
                    ? cacheErrors
                    : ["compiled artifact missing or corrupt."]
            };
        }

        var errors = new List<string>();
        var planPath = Path.Combine(artifactDir, "plan.json");
        var planSha = CompilationFingerprint.ComputeBytesSha256(planBytes);

        // Jira identity
        if (!string.Equals(metadata.JiraKey, jiraKey, StringComparison.Ordinal))
            errors.Add($"compilation metadata jiraKey '{metadata.JiraKey}' does not match requested '{jiraKey}'.");

        // Recompute fingerprint from sealed metadata components.
        string recomputed;
        try
        {
            recomputed = CompilationFingerprint.Compute(new CompilationFingerprintInput
            {
                InputSha256 = metadata.InputSha256,
                JiraRevision = metadata.JiraRevision,
                CompilerVersion = metadata.CompilerVersion,
                ParserVersion = metadata.ParserVersion,
                StepRegistrySha256 = metadata.StepRegistrySha256,
                ObjectRepositorySha256 = metadata.ObjectRepositorySha256,
                ObjectRepositoryPath = metadata.ObjectRepositoryPath,
                ApprovedKnowledgeSha256 = metadata.ApprovedKnowledgeSha256,
                CatalogFingerprint = metadata.CatalogFingerprint,
                ParameterBindingsSha256 = metadata.ParameterBindingsSha256
            });
        }
        catch (ArgumentException ex)
        {
            return new CompiledArtifactValidationResult
            {
                IsValid = false,
                JiraKey = jiraKey,
                Fingerprint = fingerprint,
                Errors = [$"fingerprint recompute failed: {ex.Message}"]
            };
        }

        if (!string.Equals(recomputed, fingerprint, StringComparison.OrdinalIgnoreCase)
            || !string.Equals(recomputed, metadata.Fingerprint, StringComparison.OrdinalIgnoreCase))
        {
            errors.Add(
                $"fingerprint recompute mismatch: directory/metadata '{fingerprint}' vs recomputed '{recomputed}'.");
        }

        // Dependency drift against current workspace artifacts.
        var registryFull = _workspace.ResolveSafePath("step-registry/registry.json");
        if (File.Exists(registryFull))
        {
            var liveRegistrySha = CompilationFingerprint.ComputeFileSha256(registryFull);
            if (!string.Equals(liveRegistrySha, metadata.StepRegistrySha256, StringComparison.OrdinalIgnoreCase))
                errors.Add("dependency-drift: step-registry sha256 no longer matches compilation metadata.");
        }

        var repoRelative = string.IsNullOrWhiteSpace(metadata.ObjectRepositoryPath)
            ? "object-repository/repository.json"
            : metadata.ObjectRepositoryPath;
        var liveRepo = _objectRepositoryReader.Read(repoRelative);
        if (liveRepo.IsValid)
        {
            var liveRepoSha = liveRepo.AggregateSha256 ?? liveRepo.ManifestSha256;
            if (!string.Equals(liveRepoSha, metadata.ObjectRepositorySha256, StringComparison.OrdinalIgnoreCase))
                errors.Add("dependency-drift: object-repository sha256 no longer matches compilation metadata.");
        }
        else
        {
            errors.Add("dependency-drift: object-repository could not be loaded for drift check.");
            errors.AddRange(liveRepo.Errors);
        }

        PlanManifest? plan;
        try
        {
            plan = JsonSerializer.Deserialize<PlanManifest>(planBytes, CompilationJson.DocumentRead);
        }
        catch (JsonException ex)
        {
            return new CompiledArtifactValidationResult
            {
                IsValid = false,
                JiraKey = jiraKey,
                Fingerprint = fingerprint,
                Errors = [$"invalid plan.json ({ex.Message})."]
            };
        }

        if (plan is null)
        {
            errors.Add("plan.json deserialized empty.");
        }
        else
        {
            if (!string.Equals(plan.PlanId, jiraKey, StringComparison.Ordinal))
                errors.Add($"planId '{plan.PlanId}' does not match requested jiraKey '{jiraKey}'.");

            var validation = _planValidator.Validate(plan, ToRelative(planPath));
            errors.AddRange(validation.Errors);

            if (!string.IsNullOrWhiteSpace(plan.ObjectRepository))
            {
                var repo = _objectRepositoryReader.Read(plan.ObjectRepository);
                if (!repo.IsValid || repo.Snapshot is null)
                {
                    errors.AddRange(repo.Errors);
                }
                else
                {
                    var expander = new PlanObjectReferenceExpander(new ObjectReferenceResolver());
                    var expansion = expander.Expand(plan, repo.Snapshot, ToRelative(planPath));
                    errors.AddRange(expansion.Errors);
                }
            }
        }

        return new CompiledArtifactValidationResult
        {
            IsValid = errors.Count == 0,
            JiraKey = jiraKey,
            Fingerprint = fingerprint,
            PlanPath = ToRelative(planPath),
            PlanSha256 = planSha,
            Errors = errors
        };
    }

    private string ResolveTestCasePath(string jiraKey, string? explicitPath)
    {
        if (!string.IsNullOrWhiteSpace(explicitPath))
            return _workspace.ResolveSafePath(explicitPath);

        var candidates = new[]
        {
            Path.Combine(_workspace.RootPath, "test-cases", "examples", $"{jiraKey}.json"),
            Path.Combine(_workspace.RootPath, "test-cases", $"{jiraKey}.json")
        };

        foreach (var candidate in candidates)
        {
            if (File.Exists(candidate))
                return candidate;
        }

        throw new WorkspaceException(
            $"test-case for '{jiraKey}' not found under test-cases/ or test-cases/examples/. Pass --file.");
    }

    private string ToRelative(string fullPath)
    {
        var root = Path.GetFullPath(_workspace.RootPath)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var normalized = Path.GetFullPath(fullPath);
        if (WorkspaceManager.PathsEqual(normalized, root))
            return ".";
        return Path.GetRelativePath(root, normalized).Replace('\\', '/');
    }
}

public sealed class CompileKeyRequest
{
    public required string JiraKey { get; init; }

    public string? TestCasePath { get; init; }

    public string? RegistryPath { get; init; }

    public string? ObjectRepositoryPath { get; init; }

    public string? KnowledgePath { get; init; }

    public string? CatalogPath { get; init; }

    public IReadOnlyDictionary<string, string>? ParameterBindings { get; init; }

    public bool Explain { get; init; }

    public DateTimeOffset? CompiledAtUtc { get; init; }
}
