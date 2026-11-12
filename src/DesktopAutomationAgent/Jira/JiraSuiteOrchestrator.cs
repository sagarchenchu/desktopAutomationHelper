using System.Diagnostics;
using System.Text;
using System.Text.Json;
using DesktopAutomationAgent.Cli;
using DesktopAutomationAgent.Compilation;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Execution;
using DesktopAutomationAgent.Suites;
using DesktopAutomationAgent.Workspace;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Jira;

/// <summary>
/// Sync / compile / run-key / run-suite orchestration for Phase 4E.
/// Sequential only; fail-fast by default; optional bounded continue-on-failure.
/// </summary>
public sealed class JiraSuiteOrchestrator
{
    private readonly AgentOptions _options;
    private readonly IWorkspaceManager _workspace;
    private readonly IJiraTestCaseSource _source;
    private readonly BddCompilationService _compiler;
    private readonly ISuiteManifestReader _suiteReader;
    private readonly IDeterministicPlanRunner _runner;
    private readonly JiraRawCache _cache;
    private readonly ILogger<JiraSuiteOrchestrator> _logger;

    public JiraSuiteOrchestrator(
        IOptions<AgentOptions> options,
        IWorkspaceManager workspace,
        IJiraTestCaseSource source,
        BddCompilationService compiler,
        ISuiteManifestReader suiteReader,
        IDeterministicPlanRunner runner,
        ILogger<JiraSuiteOrchestrator> logger)
    {
        _options = options.Value;
        _workspace = workspace;
        _source = source;
        _compiler = compiler;
        _suiteReader = suiteReader;
        _runner = runner;
        _cache = new JiraRawCache(workspace.RootPath);
        _logger = logger;
    }

    public async Task<JiraSyncKeyResult> SyncKeyAsync(
        string jiraKey,
        CancellationToken cancellationToken = default)
    {
        AgentOptionsValidator.Validate(_options, OptionsValidationScope.Jira);

        if (!JiraKeyContract.IsCanonical(jiraKey))
        {
            return new JiraSyncKeyResult
            {
                Success = false,
                ExitCode = ExitCodes.UsageOrConfiguration,
                JiraKey = jiraKey,
                Errors = [$"invalid jiraKey '{jiraKey}' — rejected before network."]
            };
        }

        var fetch = await _source.FetchAsync(jiraKey, cancellationToken).ConfigureAwait(false);
        return new JiraSyncKeyResult
        {
            Success = fetch.Success,
            ExitCode = fetch.ExitCode,
            JiraKey = jiraKey,
            FromCache = fetch.FromCache || fetch.NotModified,
            NotModified = fetch.NotModified,
            RevisionToken = fetch.Metadata?.RevisionToken,
            ContentSha256 = fetch.Metadata?.ContentSha256,
            CacheDirectory = fetch.Success
                ? ToRelative(_cache.GetArtifactDirectory(jiraKey, fetch.Metadata!.RevisionToken))
                : null,
            RateLimit = fetch.RateLimit,
            Errors = fetch.Errors,
            Warnings = fetch.Warnings
        };
    }

    public async Task<CompilationResult> CompileKeyFromJiraAsync(
        CompileKeyFromJiraRequest request,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);

        if (!JiraKeyContract.IsCanonical(request.JiraKey))
        {
            return new CompilationResult
            {
                Success = false,
                ExitCode = ExitCodes.UsageOrConfiguration,
                JiraKey = request.JiraKey,
                Errors = [$"invalid jiraKey '{request.JiraKey}' — rejected before network."]
            };
        }

        var sync = await SyncKeyAsync(request.JiraKey, cancellationToken).ConfigureAwait(false);
        if (!sync.Success)
        {
            return new CompilationResult
            {
                Success = false,
                ExitCode = sync.ExitCode,
                JiraKey = request.JiraKey,
                Errors = sync.Errors.ToList()
            };
        }

        if (!_cache.TryRead(request.JiraKey, sync.RevisionToken, out _, out _, out var dir, out var cacheErrors)
            || dir is null)
        {
            return new CompilationResult
            {
                Success = false,
                ExitCode = ExitCodes.SuiteOrWorkspace,
                JiraKey = request.JiraKey,
                Errors = cacheErrors
            };
        }

        var testCasePath = Path.Combine(dir, "test-case.json");
        return _compiler.CompileKey(new CompileKeyRequest
        {
            JiraKey = request.JiraKey,
            TestCasePath = testCasePath,
            RegistryPath = request.RegistryPath,
            ObjectRepositoryPath = request.ObjectRepositoryPath,
            KnowledgePath = request.KnowledgePath,
            CatalogPath = request.CatalogPath,
            Explain = request.Explain
        });
    }

    public async Task<JiraRunKeyResult> RunKeyAsync(
        JiraRunKeyRequest request,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);
        var started = DateTimeOffset.UtcNow;
        var sw = Stopwatch.StartNew();

        if (!JiraKeyContract.IsCanonical(request.JiraKey))
        {
            return new JiraRunKeyResult
            {
                Success = false,
                ExitCode = ExitCodes.UsageOrConfiguration,
                JiraKey = request.JiraKey,
                Result = "invalid-key",
                DurationMilliseconds = sw.Elapsed.TotalMilliseconds,
                Errors = [$"invalid jiraKey '{request.JiraKey}' — rejected before network."]
            };
        }

        CompilationResult compile;
        if (request.FromJira)
        {
            AgentOptionsValidator.Validate(_options, OptionsValidationScope.Jira);
            compile = await CompileKeyFromJiraAsync(
                    new CompileKeyFromJiraRequest
                    {
                        JiraKey = request.JiraKey,
                        RegistryPath = request.RegistryPath,
                        ObjectRepositoryPath = request.ObjectRepositoryPath,
                        KnowledgePath = request.KnowledgePath,
                        CatalogPath = request.CatalogPath
                    },
                    cancellationToken)
                .ConfigureAwait(false);
        }
        else
        {
            // Prefer jira-cache test-case when present; else offline test-cases/.
            string? testCasePath = null;
            if (_cache.TryRead(request.JiraKey, null, out _, out _, out var cacheDir, out _)
                && cacheDir is not null)
            {
                testCasePath = Path.Combine(cacheDir, "test-case.json");
            }

            compile = _compiler.CompileKey(new CompileKeyRequest
            {
                JiraKey = request.JiraKey,
                TestCasePath = testCasePath,
                RegistryPath = request.RegistryPath,
                ObjectRepositoryPath = request.ObjectRepositoryPath,
                KnowledgePath = request.KnowledgePath,
                CatalogPath = request.CatalogPath
            });
        }

        if (!compile.Success || string.IsNullOrWhiteSpace(compile.PlanPath))
        {
            return new JiraRunKeyResult
            {
                Success = false,
                ExitCode = compile.ExitCode == 0 ? ExitCodes.SuiteOrWorkspace : compile.ExitCode,
                JiraKey = request.JiraKey,
                CompilationFingerprint = compile.Fingerprint,
                PlanHash = compile.PlanSha256,
                Result = "compile-failed",
                DurationMilliseconds = sw.Elapsed.TotalMilliseconds,
                Errors = compile.Errors,
                StartedAtUtc = started,
                FinishedAtUtc = DateTimeOffset.UtcNow
            };
        }

        var planFull = _workspace.ResolveSafePath(compile.PlanPath);
        var report = await _runner.RunAsync(planFull, request.DryRun, cancellationToken)
            .ConfigureAwait(false);

        return new JiraRunKeyResult
        {
            Success = report.ExitCode == ExitCodes.Success,
            ExitCode = report.ExitCode,
            JiraKey = request.JiraKey,
            CompilationFingerprint = compile.Fingerprint,
            PlanHash = compile.PlanSha256 ?? report.PlanSha256,
            RunId = report.RunId,
            Result = report.Status,
            DryRun = request.DryRun,
            DurationMilliseconds = sw.Elapsed.TotalMilliseconds,
            RunReport = report,
            StartedAtUtc = started,
            FinishedAtUtc = DateTimeOffset.UtcNow,
            Errors = report.Failure is null
                ? []
                : [SecretRedactor.Redact(report.Failure.Message)]
        };
    }

    public async Task<JiraSuiteRunResult> RunSuiteAsync(
        JiraSuiteRunRequest request,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);
        var started = DateTimeOffset.UtcNow;
        var sw = Stopwatch.StartNew();

        AgentOptionsValidator.Validate(_options, OptionsValidationScope.Suites);
        var suite = _suiteReader.ValidateFile(request.SuiteFile);
        if (!suite.IsValid)
        {
            return new JiraSuiteRunResult
            {
                Success = false,
                ExitCode = ExitCodes.SuiteOrWorkspace,
                SuiteFile = request.SuiteFile,
                SuiteName = suite.SuiteName,
                Errors = suite.Errors,
                StartedAtUtc = started,
                FinishedAtUtc = DateTimeOffset.UtcNow,
                DurationMilliseconds = sw.Elapsed.TotalMilliseconds
            };
        }

        if (!suite.SuiteEnabled)
        {
            return new JiraSuiteRunResult
            {
                Success = false,
                ExitCode = ExitCodes.SuiteOrWorkspace,
                SuiteFile = suite.FilePath,
                SuiteName = suite.SuiteName,
                Errors = ["suite is disabled; refusing to execute."],
                StartedAtUtc = started,
                FinishedAtUtc = DateTimeOffset.UtcNow,
                DurationMilliseconds = sw.Elapsed.TotalMilliseconds
            };
        }

        var suiteRunId = $"suite-{DateTimeOffset.UtcNow:yyyyMMdd-HHmmss}-{Guid.NewGuid():N}"[..48];
        var suiteDir = Path.Combine(_workspace.RootPath, "runs", suiteRunId);
        Directory.CreateDirectory(suiteDir);

        var continueOnFailure = request.ContinueOnFailure;
        var maxFailures = request.MaxFailures <= 0 ? int.MaxValue : request.MaxFailures;
        var results = new List<JiraSuiteKeyResult>();
        var failureCount = 0;
        var cancelled = false;
        string? stopReason = null;

        foreach (var key in suite.EnabledJiraKeys)
        {
            JiraRunKeyResult keyResult;
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                keyResult = await RunKeyAsync(
                        new JiraRunKeyRequest
                        {
                            JiraKey = key,
                            DryRun = request.DryRun,
                            FromJira = request.FromJira,
                            RegistryPath = request.RegistryPath,
                            ObjectRepositoryPath = request.ObjectRepositoryPath,
                            KnowledgePath = request.KnowledgePath,
                            CatalogPath = request.CatalogPath
                        },
                        cancellationToken)
                    .ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                cancelled = true;
                stopReason = "cancelled";
                results.Add(new JiraSuiteKeyResult
                {
                    JiraKey = key,
                    Result = "cancelled",
                    DurationMilliseconds = 0
                });
                break;
            }

            results.Add(new JiraSuiteKeyResult
            {
                JiraKey = key,
                CompilationFingerprint = keyResult.CompilationFingerprint,
                PlanHash = keyResult.PlanHash,
                RunId = keyResult.RunId,
                Result = keyResult.Result,
                DurationMilliseconds = keyResult.DurationMilliseconds,
                ExitCode = keyResult.ExitCode,
                Errors = keyResult.Errors
            });

            if (!keyResult.Success)
            {
                failureCount++;
                if (!continueOnFailure)
                {
                    stopReason = "fail-fast";
                    break;
                }

                if (failureCount >= maxFailures)
                {
                    stopReason = "max-failures";
                    break;
                }
            }
        }

        var skippedDisabled = suite.DisabledCount;
        var allKeysOk = results.Count > 0
            && results.All(r => IsSuccessfulKeyResult(r.Result));
        var summary = new JiraSuiteRunResult
        {
            Success = !cancelled && stopReason is null && allKeysOk,
            ExitCode = cancelled
                ? ExitCodes.Cancelled
                : allKeysOk && stopReason is null
                    ? ExitCodes.Success
                    : results.LastOrDefault(r => r.ExitCode is not null and not ExitCodes.Success)?.ExitCode
                      ?? ExitCodes.ExecutionFailure,
            SuiteRunId = suiteRunId,
            SuiteFile = ToRelative(suite.FilePath),
            SuiteName = suite.SuiteName,
            DryRun = request.DryRun,
            FromJira = request.FromJira,
            FailFast = !continueOnFailure,
            ContinueOnFailure = continueOnFailure,
            MaxFailures = continueOnFailure ? (maxFailures == int.MaxValue ? null : maxFailures) : null,
            StopReason = stopReason,
            DisabledExcluded = skippedDisabled,
            Keys = results,
            StartedAtUtc = started,
            FinishedAtUtc = DateTimeOffset.UtcNow,
            DurationMilliseconds = sw.Elapsed.TotalMilliseconds,
            SuiteReportPath = ToRelative(Path.Combine(suiteDir, "suite-run.json"))
        };

        WriteSuiteReport(suiteDir, summary);
        _logger.LogInformation(
            "Suite run {SuiteRunId} finished with {KeyCount} keys (failures={Failures}, stop={Stop}).",
            suiteRunId,
            results.Count,
            failureCount,
            stopReason ?? "none");

        return summary;
    }

    private static bool IsSuccessfulKeyResult(string? result)
    {
        if (string.IsNullOrWhiteSpace(result))
            return false;
        return result.Equals("passed", StringComparison.OrdinalIgnoreCase)
            || result.Equals("success", StringComparison.OrdinalIgnoreCase)
            || result.Equals("succeeded", StringComparison.OrdinalIgnoreCase)
            || result.Equals("dry-run", StringComparison.OrdinalIgnoreCase)
            || result.Equals("dryrun", StringComparison.OrdinalIgnoreCase)
            || result.Equals("dry-run-passed", StringComparison.OrdinalIgnoreCase);
    }

    private void WriteSuiteReport(string suiteDir, JiraSuiteRunResult summary)
    {
        var path = Path.Combine(suiteDir, "suite-run.json");
        var json = SecretRedactor.Redact(JsonSerializer.Serialize(summary, CompilationJson.Indented));
        if (!json.EndsWith('\n'))
            json += "\n";
        var bytes = Encoding.UTF8.GetBytes(json);
        var temp = path + $".tmp-{Guid.NewGuid():N}";
        File.WriteAllBytes(temp, bytes);
        File.Move(temp, path, overwrite: false);
    }

    private string ToRelative(string fullPath)
    {
        var root = Path.GetFullPath(_workspace.RootPath)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var full = Path.GetFullPath(fullPath);
        if (full.StartsWith(root, StringComparison.OrdinalIgnoreCase))
        {
            var rel = full[root.Length..].TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            return rel.Replace('\\', '/');
        }

        return full.Replace('\\', '/');
    }
}

public sealed class CompileKeyFromJiraRequest
{
    public required string JiraKey { get; init; }
    public string? RegistryPath { get; init; }
    public string? ObjectRepositoryPath { get; init; }
    public string? KnowledgePath { get; init; }
    public string? CatalogPath { get; init; }
    public bool Explain { get; init; }
}

public sealed class JiraRunKeyRequest
{
    public required string JiraKey { get; init; }
    public bool DryRun { get; init; }
    public bool FromJira { get; init; }
    public string? RegistryPath { get; init; }
    public string? ObjectRepositoryPath { get; init; }
    public string? KnowledgePath { get; init; }
    public string? CatalogPath { get; init; }
}

public sealed class JiraSuiteRunRequest
{
    public required string SuiteFile { get; init; }
    public bool DryRun { get; init; }
    public bool FromJira { get; init; }
    public bool ContinueOnFailure { get; init; }
    public int MaxFailures { get; init; }
    public string? RegistryPath { get; init; }
    public string? ObjectRepositoryPath { get; init; }
    public string? KnowledgePath { get; init; }
    public string? CatalogPath { get; init; }
}

public sealed class JiraSyncKeyResult
{
    public bool Success { get; init; }
    public int ExitCode { get; init; }
    public string? JiraKey { get; init; }
    public bool FromCache { get; init; }
    public bool NotModified { get; init; }
    public string? RevisionToken { get; init; }
    public string? ContentSha256 { get; init; }
    public string? CacheDirectory { get; init; }
    public JiraRateLimitInfo? RateLimit { get; init; }
    public IReadOnlyList<string> Errors { get; init; } = [];
    public IReadOnlyList<string> Warnings { get; init; } = [];
}

public sealed class JiraRunKeyResult
{
    public bool Success { get; init; }
    public int ExitCode { get; init; }
    public string? JiraKey { get; init; }
    public string? CompilationFingerprint { get; init; }
    public string? PlanHash { get; init; }
    public string? RunId { get; init; }
    public string? Result { get; init; }
    public bool DryRun { get; init; }
    public double DurationMilliseconds { get; init; }
    public DateTimeOffset? StartedAtUtc { get; init; }
    public DateTimeOffset? FinishedAtUtc { get; init; }
    public RunReport? RunReport { get; init; }
    public IReadOnlyList<string> Errors { get; init; } = [];
}

public sealed record JiraSuiteRunResult
{
    public bool Success { get; init; }
    public int ExitCode { get; init; }
    public string? SuiteRunId { get; init; }
    public string? SuiteFile { get; init; }
    public string? SuiteName { get; init; }
    public bool DryRun { get; init; }
    public bool FromJira { get; init; }
    public bool FailFast { get; init; }
    public bool ContinueOnFailure { get; init; }
    public int? MaxFailures { get; init; }
    public string? StopReason { get; init; }
    public int DisabledExcluded { get; init; }
    public IReadOnlyList<JiraSuiteKeyResult> Keys { get; init; } = [];
    public DateTimeOffset? StartedAtUtc { get; init; }
    public DateTimeOffset? FinishedAtUtc { get; init; }
    public double DurationMilliseconds { get; init; }
    public string? SuiteReportPath { get; init; }
    public IReadOnlyList<string> Errors { get; init; } = [];
}

public sealed class JiraSuiteKeyResult
{
    public string JiraKey { get; init; } = string.Empty;
    public string? CompilationFingerprint { get; init; }
    public string? PlanHash { get; init; }
    public string? RunId { get; init; }
    public string? Result { get; init; }
    public double DurationMilliseconds { get; init; }
    public int? ExitCode { get; init; }
    public IReadOnlyList<string> Errors { get; init; } = [];
}
