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
        var runId = GenerateRunId();
        var runDirectory = Path.Combine(_workspace.RootPath, "runs", runId);
        Directory.CreateDirectory(runDirectory);

        var validation = _planReader.Read(planPath);
        if (!validation.IsValid || validation.Plan is null)
        {
            var invalidReport = BuildReport(
                runId,
                planPath,
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
                    Classification = UiFailureClassification.Catalog,
                    Message = validation.Errors.FirstOrDefault() ?? "Plan validation failed."
                });

            _artifactWriter.WriteRunReport(runDirectory, invalidReport);
            return invalidReport;
        }

        var manifest = validation.Plan;
        DriverConnection? connection = null;
        OperationsCatalogDto? catalog = null;
        var stepResults = new List<StepRunResult>();
        var onFailureResults = new List<StepRunResult>();
        RunFailure? failure = null;
        var launchSent = false;
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
                    UiFailureClassification.Catalog,
                    $"Plan catalogSchemaVersion {manifest.CatalogSchemaVersion} does not match live catalog schemaVersion {catalog.SchemaVersion}.");
            }

            var preflight = new PlanCatalogPreflight(catalog);
            var preflightErrors = preflight.Validate(manifest, validation.PlanPath);
            if (preflightErrors.Count > 0)
            {
                throw new UiExecutionException(
                    UiFailureClassification.Catalog,
                    preflightErrors[0]);
            }

            if (dryRun)
            {
                status = "validated";
                exitCode = ExitCodes.Success;
            }
            else
            {
                foreach (var step in manifest.Steps!)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var result = await ExecutePlanStepAsync(
                        connection,
                        step,
                        phase: "steps",
                        cancellationToken).ConfigureAwait(false);
                    stepResults.Add(result);

                    if (IsLaunchOperation(step.Operation))
                        launchSent = true;

                    if (!result.Success)
                    {
                        failure = BuildStepFailure(result);
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
                            launchSent,
                            cancellationToken).ConfigureAwait(false));
                }

                if (cancellationToken.IsCancellationRequested)
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
                    exitCode = failure.Classification switch
                    {
                        UiFailureClassification.AssertionFailure => ExitCodes.ExecutionFailure,
                        UiFailureClassification.OperationFailure => ExitCodes.ExecutionFailure,
                        UiFailureClassification.ExecutionTimeout => ExitCodes.ExecutionFailure,
                        UiFailureClassification.Authentication => ExitCodes.AuthOrCatalog,
                        UiFailureClassification.Catalog => ExitCodes.AuthOrCatalog,
                        UiFailureClassification.DriverUnavailable => ExitCodes.DriverUnavailable,
                        _ => ExitCodes.ExecutionFailure
                    };
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
                        launchSent,
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
                Classification = UiFailureClassification.Catalog,
                Message = ex.Message
            };
        }

        var report = BuildReport(
            runId,
            planPath,
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
            catalog?.SchemaVersion);

        _artifactWriter.WriteRunReport(runDirectory, report);
        return report;
    }

    private async Task<StepRunResult> ExecutePlanStepAsync(
        DriverConnection connection,
        PlanStep step,
        string phase,
        CancellationToken cancellationToken)
    {
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
                    success: false,
                    skipped: false,
                    skipReason: null,
                    response,
                    assertionResults,
                    failedAssertion.Message ?? "Assertion failed.",
                    stopwatch.Elapsed);
            }

            return BuildStepResult(
                step,
                phase,
                success: true,
                skipped: false,
                skipReason: null,
                response,
                assertionResults,
                error: null,
                stopwatch.Elapsed);
        }
        catch (UiExecutionException ex)
        {
            return BuildStepResult(
                step,
                phase,
                success: false,
                skipped: false,
                skipReason: null,
                ex.Response,
                Array.Empty<AssertionResult>(),
                ex.Message,
                stopwatch.Elapsed);
        }
    }

    private async Task<IReadOnlyList<StepRunResult>> ExecuteOnFailureStepsAsync(
        DriverConnection connection,
        OperationsCatalogDto catalog,
        List<PlanStep>? steps,
        bool launchSent,
        CancellationToken cancellationToken)
    {
        if (steps is null || steps.Count == 0)
            return Array.Empty<StepRunResult>();

        var results = new List<StepRunResult>(steps.Count);
        using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeoutCts.CancelAfter(TimeSpan.FromSeconds(_options.Runner.CleanupTimeoutSeconds));

        foreach (var step in steps)
        {
            if (ShouldSkipOnFailureStep(step, catalog, launchSent, out var skipReason))
            {
                results.Add(new StepRunResult
                {
                    Id = step.Id,
                    Operation = step.Operation,
                    Phase = "onFailureSteps",
                    Success = true,
                    Sensitive = step.Sensitive,
                    CaptureResponse = step.CaptureResponse,
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
                    timeoutCts.Token).ConfigureAwait(false));
            }
            catch (OperationCanceledException) when (timeoutCts.IsCancellationRequested)
            {
                results.Add(new StepRunResult
                {
                    Id = step.Id,
                    Operation = step.Operation,
                    Phase = "onFailureSteps",
                    Success = false,
                    Sensitive = step.Sensitive,
                    CaptureResponse = step.CaptureResponse,
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
        bool launchSent,
        out string skipReason)
    {
        skipReason = string.Empty;
        if (launchSent || !IsCloseOrQuitOperation(step.Operation))
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

    private static StepRunResult BuildStepResult(
        PlanStep step,
        string phase,
        bool success,
        bool skipped,
        string? skipReason,
        UiExecutionResponse? response,
        IReadOnlyList<AssertionResult> assertions,
        string? error,
        TimeSpan duration) =>
        new()
        {
            Id = step.Id,
            Operation = step.Operation,
            Phase = phase,
            Success = success,
            Sensitive = step.Sensitive,
            CaptureResponse = step.CaptureResponse,
            Skipped = skipped,
            SkipReason = skipReason,
            Arguments = step.Sensitive ? null : step.Arguments,
            ResponseValue = step.CaptureResponse ? response?.Value : null,
            Error = error,
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

    private static RunFailure BuildStepFailure(StepRunResult result) =>
        new()
        {
            Classification = result.Assertions.Any(a => !a.Passed)
                ? UiFailureClassification.AssertionFailure
                : UiFailureClassification.OperationFailure,
            Message = result.Error ?? "Step failed.",
            StepId = result.Id,
            ScreenshotPath = result.ScreenshotPath
        };

    private static RunReport BuildReport(
        string runId,
        string planPath,
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
        int? catalogSchemaVersion = null) =>
        new()
        {
            RunId = runId,
            Status = status,
            ExitCode = exitCode,
            PlanPath = validation.PlanPath,
            PlanId = validation.PlanId,
            PlanName = validation.Name,
            PlanSha256 = validation.Sha256,
            DryRun = dryRun,
            DriverBaseUrl = driverBaseUrl,
            CatalogSchemaVersion = catalogSchemaVersion,
            StartedAtUtc = startedAt,
            FinishedAtUtc = finishedAt,
            Steps = steps,
            OnFailureSteps = onFailureSteps,
            Failure = failure
        };

    private static int MapExitCode(UiFailureClassification classification) =>
        classification switch
        {
            UiFailureClassification.DriverUnavailable => ExitCodes.DriverUnavailable,
            UiFailureClassification.Authentication => ExitCodes.AuthOrCatalog,
            UiFailureClassification.Catalog => ExitCodes.AuthOrCatalog,
            UiFailureClassification.Cancelled => ExitCodes.Cancelled,
            _ => ExitCodes.ExecutionFailure
        };

    private static bool IsLaunchOperation(string operation) =>
        string.Equals(operation, "launch", StringComparison.OrdinalIgnoreCase);

    private static bool IsCloseOrQuitOperation(string operation) =>
        string.Equals(operation, "close", StringComparison.OrdinalIgnoreCase)
        || string.Equals(operation, "quit", StringComparison.OrdinalIgnoreCase);

    internal static string GenerateRunId()
    {
        var timestamp = DateTime.UtcNow.ToString("yyyyMMddTHHmmssfffZ");
        var suffix = Guid.NewGuid().ToString("N")[..8];
        return $"{timestamp}-{suffix}";
    }
}
