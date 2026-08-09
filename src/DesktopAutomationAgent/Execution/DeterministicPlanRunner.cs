using System.Diagnostics;
using System.Text.Json;
using DesktopAutomationAgent.Cli;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Driver.Models;
using DesktopAutomationAgent.Plans;
using DesktopAutomationAgent.Workspace;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Execution;

public sealed class DeterministicPlanRunner : IDeterministicPlanRunner
{
    private readonly AgentOptions _options;
    private readonly PlanManifestReader _planReader;
    private readonly IWorkspaceManager _workspace;
    private readonly IDriverConnectionResolver _connectionResolver;
    private readonly IDriverCatalogClient _catalogClient;
    private readonly IDriverUiClient _uiClient;
    private readonly AssertionEvaluator _assertionEvaluator;
    private readonly RunArtifactWriter _artifactWriter;
    private readonly ILogger<DeterministicPlanRunner> _logger;

    public DeterministicPlanRunner(
        IOptions<AgentOptions> options,
        PlanManifestReader planReader,
        IWorkspaceManager workspace,
        IDriverConnectionResolver connectionResolver,
        IDriverCatalogClient catalogClient,
        IDriverUiClient uiClient,
        AssertionEvaluator assertionEvaluator,
        RunArtifactWriter artifactWriter,
        ILogger<DeterministicPlanRunner> logger)
    {
        _options = options.Value;
        _planReader = planReader;
        _workspace = workspace;
        _connectionResolver = connectionResolver;
        _catalogClient = catalogClient;
        _uiClient = uiClient;
        _assertionEvaluator = assertionEvaluator;
        _artifactWriter = artifactWriter;
        _logger = logger;
    }

    public async Task<RunReport> RunAsync(
        string planPath,
        bool dryRun,
        CancellationToken cancellationToken = default)
    {
        var startedAt = DateTimeOffset.UtcNow;
        var (runId, runDirectory) = ReserveRunDirectory();

        var validation = _planReader.Read(planPath);
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

        var manifest = validation.Plan;
        DriverConnection? connection = null;
        OperationsCatalogDto? catalog = null;
        var stepResults = new List<StepRunResult>();
        var onFailureResults = new List<StepRunResult>();
        RunFailure? failure = null;
        var launchRequestSent = false;
        var sessionPossiblyActive = false;
        var status = "failed";
        var exitCode = ExitCodes.ExecutionFailure;

        try
        {
            AgentOptionsValidator.Validate(_options, OptionsValidationScope.Driver);
            AgentOptionsValidator.ValidateRunnerOptions(_options.Runner);
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

                if (failure is not null || cancellationToken.IsCancellationRequested)
                {
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
            catalog?.SchemaVersion);

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
                    stopwatch.Elapsed);
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
                stopwatch.Elapsed);
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
                stopwatch.Elapsed);
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
        TimeSpan duration)
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
            Classification = result.Assertions.Any(a => !a.Passed)
                ? UiFailureClassification.AssertionFailure
                : UiFailureClassification.OperationFailure,
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
        int? catalogSchemaVersion = null) =>
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
            ArtifactWriteStatus = "pending"
        };

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
