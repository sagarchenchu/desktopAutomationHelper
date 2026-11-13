using System.Diagnostics;
using System.Text.Json;
using DesktopAutomationAgent.Cli;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Driver.Models;
using DesktopAutomationAgent.ObjectRepository;
using DesktopAutomationAgent.Plans;
using DesktopAutomationAgent.Recovery;
using DesktopAutomationAgent.Workspace;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Execution;

public sealed class DeterministicPlanRunner : IDeterministicPlanRunner
{
    private readonly AgentOptions _options;
    private readonly PlanManifestReader _planReader;
    private readonly PlanObjectRepositoryIntegrator _objectRepositoryIntegrator;
    private readonly IWorkspaceManager _workspace;
    private readonly IDriverConnectionResolver _connectionResolver;
    private readonly IDriverCatalogClient _catalogClient;
    private readonly IDriverUiClient _uiClient;
    private readonly AssertionEvaluator _assertionEvaluator;
    private readonly RunArtifactWriter _artifactWriter;
    private readonly RecoveryLadderService _recoveryLadder;
    private readonly RecoveryArtifactWriter _recoveryArtifactWriter;
    private readonly RecoveryManifestLoader _recoveryManifestLoader;
    private readonly ObjectRepositoryReader _objectRepositoryReader;
    private readonly ILogger<DeterministicPlanRunner> _logger;

    public DeterministicPlanRunner(
        IOptions<AgentOptions> options,
        PlanManifestReader planReader,
        PlanObjectRepositoryIntegrator objectRepositoryIntegrator,
        IWorkspaceManager workspace,
        IDriverConnectionResolver connectionResolver,
        IDriverCatalogClient catalogClient,
        IDriverUiClient uiClient,
        AssertionEvaluator assertionEvaluator,
        RunArtifactWriter artifactWriter,
        RecoveryLadderService recoveryLadder,
        RecoveryArtifactWriter recoveryArtifactWriter,
        RecoveryManifestLoader recoveryManifestLoader,
        ObjectRepositoryReader objectRepositoryReader,
        ILogger<DeterministicPlanRunner> logger)
    {
        _options = options.Value;
        _planReader = planReader;
        _objectRepositoryIntegrator = objectRepositoryIntegrator;
        _workspace = workspace;
        _connectionResolver = connectionResolver;
        _catalogClient = catalogClient;
        _uiClient = uiClient;
        _assertionEvaluator = assertionEvaluator;
        _artifactWriter = artifactWriter;
        _recoveryLadder = recoveryLadder;
        _recoveryArtifactWriter = recoveryArtifactWriter;
        _recoveryManifestLoader = recoveryManifestLoader;
        _objectRepositoryReader = objectRepositoryReader;
        _logger = logger;
    }

    public async Task<RunReport> RunAsync(
        string planPath,
        bool dryRun,
        CancellationToken cancellationToken = default)
    {
        var startedAt = DateTimeOffset.UtcNow;
        string runId;
        string runDirectory;
        try
        {
            (runId, runDirectory) = ReserveRunDirectory();
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Security.SecurityException)
        {
            _logger.LogError(ex, "Failed to reserve run artifact directory under {WorkspaceRoot}", _workspace.RootPath);
            return new RunReport
            {
                ReportSchemaVersion = 1,
                RunId = GenerateRunId(),
                Status = "failed",
                ExitCode = ExitCodes.SuiteOrWorkspace,
                PlanPath = planPath,
                DryRun = dryRun,
                StartedAtUtc = startedAt,
                FinishedAtUtc = DateTimeOffset.UtcNow,
                DurationMilliseconds = (DateTimeOffset.UtcNow - startedAt).TotalMilliseconds,
                ArtifactWriteStatus = "failed",
                Failure = new RunFailure
                {
                    Classification = UiFailureClassification.ArtifactFailure,
                    Message = $"Failed to reserve run artifact directory: {ex.Message}"
                }
            };
        }

        PlanValidationResult validation;
        try
        {
            validation = _planReader.Read(planPath);
        }
        catch (Exception ex)
        {
            var crashReport = BuildReport(
                runId,
                new PlanValidationResult
                {
                    PlanPath = planPath,
                    Errors = [$"Unexpected plan read failure: {ex.Message}"]
                },
                dryRun,
                startedAt,
                DateTimeOffset.UtcNow,
                status: "failed",
                exitCode: ExitCodes.SuiteOrWorkspace,
                steps: [],
                onFailureSteps: [],
                failure: new RunFailure
                {
                    Classification = UiFailureClassification.PlanValidation,
                    Message = $"Unexpected plan read failure: {ex.Message}"
                });
            return PersistReport(runDirectory, crashReport);
        }

        if (!validation.IsValid || validation.Plan is null)
        {
            var invalidReport = BuildReport(
                runId,
                validation,
                dryRun,
                startedAt,
                DateTimeOffset.UtcNow,
                status: "failed",
                exitCode: ExitCodes.SuiteOrWorkspace,
                steps: [],
                onFailureSteps: [],
                failure: new RunFailure
                {
                    Classification = UiFailureClassification.PlanValidation,
                    Message = validation.Errors.FirstOrDefault() ?? "Plan validation failed."
                });

            return PersistReport(runDirectory, invalidReport);
        }

        validation = _objectRepositoryIntegrator.Integrate(validation);
        if (!validation.IsValid || validation.Plan is null)
        {
            var invalidReport = BuildReport(
                runId,
                validation,
                dryRun,
                startedAt,
                DateTimeOffset.UtcNow,
                status: "failed",
                exitCode: ExitCodes.SuiteOrWorkspace,
                steps: [],
                onFailureSteps: [],
                failure: new RunFailure
                {
                    Classification = UiFailureClassification.PlanValidation,
                    Message = validation.Errors.FirstOrDefault() ?? "Object repository integration failed."
                });

            return PersistReport(runDirectory, invalidReport);
        }

        var manifest = validation.Plan;
        DriverConnection? connection = null;
        OperationsCatalogDto? catalog = null;
        var stepResults = new List<StepRunResult>();
        var onFailureResults = new List<StepRunResult>();
        RunFailure? failure = null;
        RecoveryRunSummary? recoverySummary = null;
        var recoveryAttempts = new List<RecoveryAttemptRecord>();
        var launchRequestSent = false;
        var sessionPossiblyActive = false;
        var status = "failed";
        var exitCode = ExitCodes.ExecutionFailure;
        var originalPlanSha256 = validation.Sha256 ?? string.Empty;
        var originalPlanBytes = SafeReadPlanBytes(validation.PlanPath);

        try
        {
            AgentOptionsValidator.Validate(_options, OptionsValidationScope.Driver);
            AgentOptionsValidator.ValidateRunnerOptions(_options.Runner);
            AgentOptionsValidator.ValidateRecoveryOptions(_options.Recovery);
            _workspace.EnsureInitialized();

            connection = await _connectionResolver.ResolveAsync(cancellationToken).ConfigureAwait(false);
            var driverStatus = await _catalogClient.GetStatusAsync(connection, cancellationToken).ConfigureAwait(false);
            if (!driverStatus.Ready)
            {
                throw new UiExecutionException(
                    UiFailureClassification.DriverUnavailable,
                    "Driver /status reported ready=false.");
            }

            catalog = await _catalogClient.GetOperationsAsync(connection, cancellationToken).ConfigureAwait(false);
            CatalogCompatibility.Validate(catalog, _options.Driver.ExpectedCatalogSchemaVersion);

            if (manifest.CatalogSchemaVersion != catalog.SchemaVersion)
            {
                throw new UiExecutionException(
                    UiFailureClassification.PlanValidation,
                    $"Plan catalogSchemaVersion {manifest.CatalogSchemaVersion} does not match live catalog schemaVersion {catalog.SchemaVersion}.");
            }

            var preflight = new PlanCatalogPreflight(catalog);
            var preflightErrors = preflight.Validate(manifest, validation.PlanPath);
            if (preflightErrors.Count > 0)
            {
                throw new UiExecutionException(
                    UiFailureClassification.PlanValidation,
                    string.Join(' ', preflightErrors));
            }

            if (dryRun)
            {
                status = "validated";
                exitCode = ExitCodes.Success;
            }
            else
            {
                using var recoveryBudget = new RecoveryBudgetTracker(_options.Recovery);
                var (approvedCandidateLocators, approvedCandidateIds) = _recoveryManifestLoader.LoadApprovedCandidates();
                var recoveryContext = new RecoveryContext
                {
                    RunId = runId,
                    RunDirectory = runDirectory,
                    PlanPath = validation.PlanPath,
                    OriginalPlanSha256 = originalPlanSha256,
                    ObjectRepositoryPath = validation.ObjectRepositoryPath,
                    ObjectRepositorySha256 = validation.ObjectRepositorySha256,
                    RepositorySnapshot = validation.ObjectRepositorySnapshot,
                    Catalog = catalog,
                    StepObjectBindings = validation.StepObjectBindings
                        ?? new Dictionary<string, IReadOnlyDictionary<string, string>>(StringComparer.Ordinal),
                    EquivalentTransitions = _recoveryManifestLoader.LoadEquivalentTransitions(),
                    ApprovedCandidateLocators = approvedCandidateLocators,
                    ApprovedCandidateLocatorIds = approvedCandidateIds,
                    ApplicationFingerprint = driverStatus.Build?.Version
                        ?? catalog?.DriverVersion,
                    ExpectedApplicationFingerprint = _recoveryManifestLoader.LoadExpectedApplicationFingerprint()
                };

                var steps = manifest.Steps!;
                for (var i = 0; i < steps.Count; i++)
                {
                    if (cancellationToken.IsCancellationRequested)
                    {
                        AppendSkippedSteps(stepResults, steps, i, "cancelled");
                        failure = new RunFailure
                        {
                            Classification = UiFailureClassification.Cancelled,
                            Message = "Plan execution was cancelled."
                        };
                        break;
                    }

                    var step = steps[i];
                    var isLaunch = IsLaunchOperation(step.Operation);
                    if (isLaunch)
                        launchRequestSent = true;

                    var result = await ExecutePlanStepAsync(
                        connection,
                        step,
                        phase: "steps",
                        sequence: i + 1,
                        cancellationToken).ConfigureAwait(false);

                    if (!result.Success && _options.Recovery.Enabled)
                    {
                        var activeConnection = connection;
                        var outcome = await _recoveryLadder.TryRecoverAsync(
                            activeConnection,
                            step,
                            result,
                            recoveryContext,
                            recoveryBudget,
                            async (candidateStep, ct) => await ExecutePlanStepAsync(
                                activeConnection,
                                candidateStep,
                                phase: "steps",
                                sequence: i + 1,
                                ct).ConfigureAwait(false),
                            cancellationToken).ConfigureAwait(false);

                        recoveryAttempts.AddRange(outcome.Attempts);

                        if (outcome.Recovered && outcome.RecoveredStepResult is not null)
                        {
                            result = outcome.RecoveredStepResult;
                            _logger.LogInformation(
                                "Bounded recovery recovered step {StepId} via tier {Tier} ({Classification}).",
                                step.Id,
                                outcome.TierUsed,
                                RecoveryClassificationRules.DisplayName(outcome.Classification));
                        }
                        else if (outcome.StopForApproval)
                        {
                            result = new StepRunResult
                            {
                                Sequence = result.Sequence,
                                Id = result.Id,
                                Operation = result.Operation,
                                Phase = result.Phase,
                                Status = "failed",
                                Success = false,
                                Sensitive = result.Sensitive,
                                CaptureResponse = result.CaptureResponse,
                                HttpStatusCode = result.HttpStatusCode,
                                StartedAtUtc = result.StartedAtUtc,
                                Arguments = result.Arguments,
                                ResponseValue = result.ResponseValue,
                                Error = outcome.Detail ?? result.Error,
                                DriverReason = result.DriverReason,
                                ScreenshotPath = result.ScreenshotPath,
                                Classification = result.Classification,
                                Assertions = result.Assertions,
                                Duration = result.Duration
                            };
                        }
                    }

                    stepResults.Add(result);

                    if (isLaunch && result.Success)
                        sessionPossiblyActive = true;
                    else if (isLaunch)
                        sessionPossiblyActive = true; // uncertain outcome — attempt cleanup
                    else if (result.Success && IsCloseOrQuitOperation(step.Operation))
                        sessionPossiblyActive = false;

                    if (!result.Success)
                    {
                        failure = BuildStepFailure(result);
                        AppendSkippedSteps(stepResults, steps, i + 1, "previousStepFailed");
                        break;
                    }
                }

                recoverySummary = BuildRecoverySummary(
                    runId,
                    originalPlanSha256,
                    validation.ObjectRepositorySha256,
                    originalPlanBytes,
                    validation.PlanPath,
                    validation.ObjectRepositoryPath,
                    recoveryAttempts,
                    failure is null ? "passed" : "failed",
                    _objectRepositoryReader);

                if (recoveryAttempts.Count > 0)
                {
                    try
                    {
                        var recoveryDir = _recoveryArtifactWriter.EnsureRecoveryDirectory(runDirectory);
                        _recoveryArtifactWriter.WriteSummary(recoveryDir, recoverySummary);
                    }
                    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                    {
                        _logger.LogWarning(ex, "Failed to persist recovery summary for run {RunId}", runId);
                    }
                }

                if (failure is not null || cancellationToken.IsCancellationRequested)
                {
                    // Cleanup is never subject to recovery (deterministic, at most once).
                    onFailureResults.AddRange(
                        await ExecuteOnFailureStepsAsync(
                            connection,
                            catalog,
                            manifest.OnFailureSteps,
                            sessionPossiblyActive || launchRequestSent,
                            cancellationToken).ConfigureAwait(false));
                }

                if (cancellationToken.IsCancellationRequested || failure?.Classification == UiFailureClassification.Cancelled)
                {
                    status = "cancelled";
                    exitCode = ExitCodes.Cancelled;
                    failure ??= new RunFailure
                    {
                        Classification = UiFailureClassification.Cancelled,
                        Message = "Plan execution was cancelled."
                    };
                }
                else if (failure is not null)
                {
                    status = "failed";
                    exitCode = MapExitCode(failure.Classification);
                }
                else
                {
                    status = "passed";
                    exitCode = ExitCodes.Success;
                }
            }
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            if (connection is not null && catalog is not null)
            {
                onFailureResults.AddRange(
                    await ExecuteOnFailureStepsAsync(
                        connection,
                        catalog,
                        manifest.OnFailureSteps,
                        sessionPossiblyActive || launchRequestSent,
                        CancellationToken.None).ConfigureAwait(false));
            }

            status = "cancelled";
            exitCode = ExitCodes.Cancelled;
            failure = new RunFailure
            {
                Classification = UiFailureClassification.Cancelled,
                Message = "Plan execution was cancelled."
            };
        }
        catch (UiExecutionException ex)
        {
            status = "failed";
            exitCode = MapExitCode(ex.Classification);
            failure = new RunFailure
            {
                Classification = ex.Classification,
                Message = ex.Message,
                DriverReason = ex.Response?.Reason,
                ScreenshotPath = ex.Response?.ScreenshotPath
            };
        }
        catch (DriverConnectionException ex)
        {
            status = "failed";
            exitCode = ExitCodes.DriverUnavailable;
            failure = new RunFailure
            {
                Classification = UiFailureClassification.DriverUnavailable,
                Message = ex.Message
            };
        }
        catch (DriverCatalogException ex)
        {
            status = "failed";
            exitCode = ExitCodes.AuthOrCatalog;
            failure = new RunFailure
            {
                Classification = UiFailureClassification.Catalog,
                Message = ex.Message
            };
        }
        catch (AgentConfigurationException ex)
        {
            status = "failed";
            exitCode = ExitCodes.UsageOrConfiguration;
            failure = new RunFailure
            {
                Classification = UiFailureClassification.PlanValidation,
                Message = ex.Message
            };
        }
        catch (WorkspaceException ex)
        {
            status = "failed";
            exitCode = ExitCodes.SuiteOrWorkspace;
            failure = new RunFailure
            {
                Classification = UiFailureClassification.PlanValidation,
                Message = ex.Message
            };
        }

        var report = BuildReport(
            runId,
            validation,
            dryRun,
            startedAt,
            DateTimeOffset.UtcNow,
            status,
            exitCode,
            stepResults,
            onFailureResults,
            failure,
            connection?.SafeBaseUrl,
            connection?.DiscoveryMethod,
            catalog?.DriverVersion,
            catalog?.SchemaVersion,
            recoverySummary);

        return PersistReport(runDirectory, report);
    }

    private RunReport PersistReport(string runDirectory, RunReport report)
    {
        try
        {
            _artifactWriter.WriteRunReport(runDirectory, report);
            return report;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Security.SecurityException)
        {
            _logger.LogError(ex, "Failed to write run report under {RunDirectory}", runDirectory);
            report.ArtifactWriteStatus = "failed";
            return new RunReport
            {
                ReportSchemaVersion = report.ReportSchemaVersion,
                RunId = report.RunId,
                Status = "failed",
                ExitCode = ExitCodes.SuiteOrWorkspace,
                PlanPath = report.PlanPath,
                PlanId = report.PlanId,
                PlanName = report.PlanName,
                PlanSha256 = report.PlanSha256,
                ObjectRepositoryPath = report.ObjectRepositoryPath,
                ObjectRepositoryId = report.ObjectRepositoryId,
                ObjectRepositorySha256 = report.ObjectRepositorySha256,
                ResolvedObjectReferences = report.ResolvedObjectReferences,
                DryRun = report.DryRun,
                DriverBaseUrl = report.DriverBaseUrl,
                DiscoveryMethod = report.DiscoveryMethod,
                DriverVersion = report.DriverVersion,
                CatalogSchemaVersion = report.CatalogSchemaVersion,
                StartedAtUtc = report.StartedAtUtc,
                FinishedAtUtc = report.FinishedAtUtc,
                DurationMilliseconds = report.DurationMilliseconds,
                Steps = report.Steps,
                OnFailureSteps = report.OnFailureSteps,
                Failure = new RunFailure
                {
                    Classification = UiFailureClassification.ArtifactFailure,
                    Message = $"Failed to persist run report: {ex.Message}",
                    StepId = report.Failure?.StepId,
                    DriverReason = report.Failure?.DriverReason,
                    ScreenshotPath = report.Failure?.ScreenshotPath
                },
                Recovery = report.Recovery,
                ArtifactWriteStatus = "failed"
            };
        }
    }

    private async Task<StepRunResult> ExecutePlanStepAsync(
        DriverConnection connection,
        PlanStep step,
        string phase,
        int sequence,
        CancellationToken cancellationToken)
    {
        var startedAt = DateTimeOffset.UtcNow;
        var stopwatch = Stopwatch.StartNew();
        try
        {
            var response = await _uiClient.ExecuteStepAsync(connection, step, cancellationToken).ConfigureAwait(false);
            var assertionResults = _assertionEvaluator.Evaluate(response.Value, step.Assertions);
            var failedAssertion = assertionResults.FirstOrDefault(a => !a.Passed);

            if (failedAssertion is not null)
            {
                return BuildStepResult(
                    step,
                    phase,
                    sequence,
                    status: "failed",
                    success: false,
                    skipped: false,
                    skipReason: null,
                    response,
                    assertionResults,
                    failedAssertion.Message ?? "Assertion failed.",
                    startedAt,
                    stopwatch.Elapsed,
                    UiFailureClassification.AssertionFailure);
            }

            return BuildStepResult(
                step,
                phase,
                sequence,
                status: "passed",
                success: true,
                skipped: false,
                skipReason: null,
                response,
                assertionResults,
                error: null,
                startedAt,
                stopwatch.Elapsed,
                classification: null);
        }
        catch (UiExecutionException ex)
        {
            return BuildStepResult(
                step,
                phase,
                sequence,
                status: "failed",
                success: false,
                skipped: false,
                skipReason: null,
                ex.Response,
                Array.Empty<AssertionResult>(),
                ex.Message,
                startedAt,
                stopwatch.Elapsed,
                ex.Classification);
        }
    }

    private async Task<IReadOnlyList<StepRunResult>> ExecuteOnFailureStepsAsync(
        DriverConnection connection,
        OperationsCatalogDto catalog,
        List<PlanStep>? steps,
        bool sessionPossiblyActive,
        CancellationToken cancellationToken)
    {
        if (steps is null || steps.Count == 0)
            return Array.Empty<StepRunResult>();

        var results = new List<StepRunResult>(steps.Count);
        using var timeoutCts = new CancellationTokenSource(
            TimeSpan.FromSeconds(_options.Runner.CleanupTimeoutSeconds));

        for (var i = 0; i < steps.Count; i++)
        {
            var step = steps[i];
            if (ShouldSkipOnFailureStep(step, catalog, sessionPossiblyActive, out var skipReason))
            {
                results.Add(new StepRunResult
                {
                    Sequence = i + 1,
                    Id = step.Id,
                    Operation = step.Operation,
                    Phase = "onFailureSteps",
                    Status = "skipped",
                    Success = true,
                    Sensitive = step.Sensitive,
                    CaptureResponse = step.ShouldCaptureResponse,
                    Skipped = true,
                    SkipReason = skipReason,
                    Duration = TimeSpan.Zero
                });
                continue;
            }

            try
            {
                results.Add(await ExecutePlanStepAsync(
                    connection,
                    step,
                    phase: "onFailureSteps",
                    sequence: i + 1,
                    timeoutCts.Token).ConfigureAwait(false));
            }
            catch (OperationCanceledException) when (timeoutCts.IsCancellationRequested)
            {
                results.Add(new StepRunResult
                {
                    Sequence = i + 1,
                    Id = step.Id,
                    Operation = step.Operation,
                    Phase = "onFailureSteps",
                    Status = "failed",
                    Success = false,
                    Sensitive = step.Sensitive,
                    CaptureResponse = step.ShouldCaptureResponse,
                    Error = $"Cleanup step timed out after {_options.Runner.CleanupTimeoutSeconds} seconds.",
                    Duration = TimeSpan.FromSeconds(_options.Runner.CleanupTimeoutSeconds)
                });
            }
        }

        return results;
    }

    private static bool ShouldSkipOnFailureStep(
        PlanStep step,
        OperationsCatalogDto catalog,
        bool sessionPossiblyActive,
        out string skipReason)
    {
        skipReason = string.Empty;
        if (sessionPossiblyActive || !IsCloseOrQuitOperation(step.Operation))
            return false;

        var descriptor = catalog.Operations.FirstOrDefault(
            op => string.Equals(op.Name, step.Operation, StringComparison.OrdinalIgnoreCase));
        if (descriptor is { RequiresSession: true })
        {
            skipReason = "noActiveSession";
            return true;
        }

        return false;
    }

    private static void AppendSkippedSteps(
        List<StepRunResult> results,
        IReadOnlyList<PlanStep> steps,
        int startIndex,
        string reason)
    {
        for (var i = startIndex; i < steps.Count; i++)
        {
            var step = steps[i];
            results.Add(new StepRunResult
            {
                Sequence = i + 1,
                Id = step.Id,
                Operation = step.Operation,
                Phase = "steps",
                Status = "skipped",
                Success = false,
                Sensitive = step.Sensitive,
                CaptureResponse = step.ShouldCaptureResponse,
                Skipped = true,
                SkipReason = reason,
                Duration = TimeSpan.Zero
            });
        }
    }

    private static StepRunResult BuildStepResult(
        PlanStep step,
        string phase,
        int sequence,
        string status,
        bool success,
        bool skipped,
        string? skipReason,
        UiExecutionResponse? response,
        IReadOnlyList<AssertionResult> assertions,
        string? error,
        DateTimeOffset startedAt,
        TimeSpan duration,
        UiFailureClassification? classification)
    {
        var captureValue = step.ShouldCaptureResponse && !step.Sensitive;
        return new StepRunResult
        {
            Sequence = sequence,
            Id = step.Id,
            Operation = step.Operation,
            Phase = phase,
            Status = status,
            Success = success,
            Sensitive = step.Sensitive,
            CaptureResponse = step.ShouldCaptureResponse,
            Skipped = skipped,
            SkipReason = skipReason,
            HttpStatusCode = response?.HttpStatusCode,
            StartedAtUtc = startedAt,
            Arguments = step.Sensitive ? null : step.Arguments,
            ResponseValue = captureValue ? response?.Value : null,
            Error = error,
            DriverReason = response?.Reason,
            ScreenshotPath = response?.ScreenshotPath,
            Classification = classification,
            Assertions = assertions.Select(a => new AssertionRunResult
            {
                Path = a.Path,
                Operator = a.Operator,
                Passed = a.Passed,
                Message = a.Message,
                Expected = step.Sensitive ? null : a.Expected,
                Actual = step.Sensitive ? null : a.Actual
            }).ToArray(),
            Duration = duration
        };
    }

    private static RunFailure BuildStepFailure(StepRunResult result) =>
        new()
        {
            Classification = result.Classification
                ?? (result.Assertions.Any(a => !a.Passed)
                    ? UiFailureClassification.AssertionFailure
                    : UiFailureClassification.OperationFailure),
            Message = result.Error ?? "Step failed.",
            StepId = result.Id,
            DriverReason = result.DriverReason,
            ScreenshotPath = result.ScreenshotPath
        };

    private static RunReport BuildReport(
        string runId,
        PlanValidationResult validation,
        bool dryRun,
        DateTimeOffset startedAt,
        DateTimeOffset finishedAt,
        string status,
        int exitCode,
        IReadOnlyList<StepRunResult> steps,
        IReadOnlyList<StepRunResult> onFailureSteps,
        RunFailure? failure,
        string? driverBaseUrl = null,
        string? discoveryMethod = null,
        string? driverVersion = null,
        int? catalogSchemaVersion = null,
        RecoveryRunSummary? recovery = null) =>
        new()
        {
            ReportSchemaVersion = 1,
            RunId = runId,
            Status = status,
            ExitCode = exitCode,
            PlanPath = validation.PlanPath,
            PlanId = validation.PlanId,
            PlanName = validation.Name,
            PlanSha256 = validation.Sha256,
            ObjectRepositoryPath = validation.ObjectRepositoryPath,
            ObjectRepositoryId = validation.ObjectRepositoryId,
            ObjectRepositorySha256 = validation.ObjectRepositorySha256,
            ResolvedObjectReferences = validation.ResolvedObjectReferences,
            DryRun = dryRun,
            DriverBaseUrl = driverBaseUrl,
            DiscoveryMethod = discoveryMethod,
            DriverVersion = driverVersion,
            CatalogSchemaVersion = catalogSchemaVersion,
            StartedAtUtc = startedAt,
            FinishedAtUtc = finishedAt,
            DurationMilliseconds = (finishedAt - startedAt).TotalMilliseconds,
            Steps = steps,
            OnFailureSteps = onFailureSteps,
            Failure = failure,
            Recovery = recovery,
            ArtifactWriteStatus = "pending"
        };

    private static RecoveryRunSummary BuildRecoverySummary(
        string runId,
        string originalPlanSha256,
        string? objectRepositorySha256,
        byte[]? originalPlanBytes,
        string planPath,
        string? objectRepositoryPath,
        List<RecoveryAttemptRecord> attempts,
        string finalResult,
        ObjectRepositoryReader repositoryReader)
    {
        var planMutated = false;
        if (originalPlanBytes is not null && File.Exists(planPath))
        {
            var current = File.ReadAllBytes(planPath);
            planMutated = !originalPlanBytes.AsSpan().SequenceEqual(current);
        }

        string? afterHash = null;
        var repoMutated = false;
        if (!string.IsNullOrWhiteSpace(objectRepositoryPath))
        {
            try
            {
                var after = repositoryReader.Read(objectRepositoryPath);
                if (after.IsValid && after.Snapshot is not null)
                {
                    afterHash = after.Snapshot.AggregateSha256;
                    if (!string.IsNullOrWhiteSpace(objectRepositorySha256)
                        && !string.Equals(objectRepositorySha256, afterHash, StringComparison.OrdinalIgnoreCase))
                    {
                        repoMutated = true;
                    }
                }
                else if (!string.IsNullOrWhiteSpace(objectRepositorySha256))
                {
                    // Unable to re-read a previously hashed repo — treat as mutated/fail-closed signal.
                    repoMutated = true;
                }
            }
            catch
            {
                repoMutated = !string.IsNullOrWhiteSpace(objectRepositorySha256);
            }
        }

        return new RecoveryRunSummary
        {
            RunId = runId,
            OriginalPlanSha256 = originalPlanSha256,
            ObjectRepositorySha256 = objectRepositorySha256,
            ObjectRepositorySha256After = afterHash,
            PlanMutated = planMutated,
            ApprovedRepositoryMutated = repoMutated,
            TotalRecoveryAttempts = attempts.Count,
            Attempts = attempts,
            FinalResult = finalResult
        };
    }

    private static byte[]? SafeReadPlanBytes(string planPath)
    {
        try
        {
            return File.Exists(planPath) ? File.ReadAllBytes(planPath) : null;
        }
        catch
        {
            return null;
        }
    }

    private static int MapExitCode(UiFailureClassification classification) =>
        classification switch
        {
            UiFailureClassification.DriverUnavailable => ExitCodes.DriverUnavailable,
            UiFailureClassification.Authentication => ExitCodes.AuthOrCatalog,
            UiFailureClassification.Catalog => ExitCodes.AuthOrCatalog,
            UiFailureClassification.PlanValidation => ExitCodes.SuiteOrWorkspace,
            UiFailureClassification.ArtifactFailure => ExitCodes.SuiteOrWorkspace,
            UiFailureClassification.Cancelled => ExitCodes.Cancelled,
            _ => ExitCodes.ExecutionFailure
        };

    private static bool IsLaunchOperation(string operation) =>
        string.Equals(operation, "launch", StringComparison.OrdinalIgnoreCase);

    private static bool IsCloseOrQuitOperation(string operation) =>
        string.Equals(operation, "close", StringComparison.OrdinalIgnoreCase)
        || string.Equals(operation, "quit", StringComparison.OrdinalIgnoreCase);

    private (string RunId, string Directory) ReserveRunDirectory()
    {
        var runsRoot = Path.Combine(_workspace.RootPath, "runs");
        Directory.CreateDirectory(runsRoot);

        for (var attempt = 0; attempt < 16; attempt++)
        {
            var runId = GenerateRunId();
            var directory = Path.Combine(runsRoot, runId);
            if (Directory.Exists(directory))
                continue;

            Directory.CreateDirectory(directory);
            return (runId, directory);
        }

        throw new IOException("Unable to reserve a unique run artifact directory.");
    }

    internal static string GenerateRunId()
    {
        var timestamp = DateTime.UtcNow.ToString("yyyyMMddTHHmmssfffZ");
        var suffix = Guid.NewGuid().ToString("N")[..8];
        return $"{timestamp}-{suffix}";
    }
}
