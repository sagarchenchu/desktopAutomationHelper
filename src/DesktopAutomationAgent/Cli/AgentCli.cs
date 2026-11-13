using System.Text.Json;
using System.Text.Json.Serialization;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Execution;
using DesktopAutomationAgent.ObjectRepository;
using DesktopAutomationAgent.Exploration;
using DesktopAutomationAgent.Compilation;
using DesktopAutomationAgent.Jira;
using DesktopAutomationAgent.Knowledge;
using DesktopAutomationAgent.Mcp;
using DesktopAutomationAgent.Observation;
using DesktopAutomationAgent.Plans;
using DesktopAutomationAgent.Planning;
using DesktopAutomationAgent.Readiness;
using DesktopAutomationAgent.Suites;
using DesktopAutomationAgent.Workspace;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Cli;

public static class AgentCli
{
    private static readonly JsonSerializerOptions JsonOutputOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) }
    };

    public static Task<int> RunAsync(string[] args, CancellationToken cancellationToken = default) =>
        RunAsync(args, BuildHost, cancellationToken);

    /// <summary>
    /// Testable entry point. <paramref name="hostBuilder"/> receives configuration
    /// args and whether JSON-only stdout mode is active.
    /// </summary>
    internal static async Task<int> RunAsync(
        string[] args,
        Func<string[], bool, IHost> hostBuilder,
        CancellationToken cancellationToken = default)
    {
        var parsed = CommandLine.Parse(args);
        // Also honor a raw --json flag so usage errors never fall back to plain text
        // when the caller requested machine-readable output.
        var jsonRequested = parsed.Json
            || args.Any(arg => string.Equals(arg, "--json", StringComparison.Ordinal));
        var jsonStdoutMode = jsonRequested
            && parsed.Kind is AgentCommandKind.Doctor
                or AgentCommandKind.ValidatePlan
                or AgentCommandKind.RunPlan
                or AgentCommandKind.ValidateObjectRepository
                or AgentCommandKind.ResolveObject
                or AgentCommandKind.CapturePage
                or AgentCommandKind.VerifyObjectRepository
                or AgentCommandKind.ObserveApp
                or AgentCommandKind.ExploreApp
                or AgentCommandKind.ValidateKnowledge
                or AgentCommandKind.BuildKnowledge
                or AgentCommandKind.CompareKnowledge
                or AgentCommandKind.SearchKnowledge
                or AgentCommandKind.ApproveKnowledge
                or AgentCommandKind.ValidateTestCase
                or AgentCommandKind.ValidateStepRegistry
                or AgentCommandKind.CompileKey
                or AgentCommandKind.ValidateCompiledKey
                or AgentCommandKind.SyncKey
                or AgentCommandKind.RunKey
                or AgentCommandKind.RunSuite
                or AgentCommandKind.ProposePlan;

        if (parsed.Error is not null)
        {
            if (jsonStdoutMode)
            {
                WriteJsonError(ExitCodes.UsageOrConfiguration, parsed.Error);
            }
            else
            {
                await Console.Error.WriteLineAsync(SecretRedactor.Redact(parsed.Error)).ConfigureAwait(false);
                await Console.Error.WriteLineAsync(CommandLine.HelpText).ConfigureAwait(false);
            }

            return ExitCodes.UsageOrConfiguration;
        }

        if (parsed.Kind == AgentCommandKind.Help)
        {
            await Console.Out.WriteLineAsync(CommandLine.HelpText).ConfigureAwait(false);
            return ExitCodes.Success;
        }

        // MCP owns its own stdio host; do not construct the CLI Generic Host first
        // (that would risk polluting stdout before JSON-RPC begins).
        if (parsed.Kind == AgentCommandKind.McpServer)
        {
            return await McpServerHost.RunAsync(parsed.ConfigurationArgs, cancellationToken)
                .ConfigureAwait(false);
        }

        using var host = hostBuilder(parsed.ConfigurationArgs, jsonStdoutMode);

        try
        {
            return parsed.Kind switch
            {
                AgentCommandKind.Init => RunInit(host.Services),
                AgentCommandKind.ValidateSuite => RunValidateSuite(host.Services, parsed.SuiteFile!),
                AgentCommandKind.ValidateKeys => RunValidateKeys(host.Services, parsed.Keys),
                AgentCommandKind.ValidatePlan => RunValidatePlan(host.Services, parsed.PlanFile!, parsed.Json),
                AgentCommandKind.RunPlan => await RunPlanAsync(
                        host.Services,
                        parsed.PlanFile!,
                        parsed.DryRun,
                        parsed.Json,
                        cancellationToken)
                    .ConfigureAwait(false),
                AgentCommandKind.Doctor => await RunDoctorAsync(host.Services, parsed.Json, cancellationToken)
                    .ConfigureAwait(false),
                AgentCommandKind.ValidateObjectRepository => RunValidateObjectRepository(
                    host.Services,
                    parsed.RepositoryFile!,
                    parsed.Json),
                AgentCommandKind.ResolveObject => RunResolveObject(
                    host.Services,
                    parsed.RepositoryFile!,
                    parsed.ObjectRef!,
                    parsed.Json),
                AgentCommandKind.CapturePage => await RunCapturePageAsync(
                        host.Services,
                        parsed,
                        cancellationToken)
                    .ConfigureAwait(false),
                AgentCommandKind.VerifyObjectRepository => await RunVerifyObjectRepositoryAsync(
                        host.Services,
                        parsed,
                        cancellationToken)
                    .ConfigureAwait(false),
                AgentCommandKind.ObserveApp => await RunObserveAppAsync(
                        host.Services,
                        parsed,
                        cancellationToken)
                    .ConfigureAwait(false),
                AgentCommandKind.ExploreApp => await RunExploreAppAsync(
                        host.Services,
                        parsed,
                        cancellationToken)
                    .ConfigureAwait(false),
                AgentCommandKind.ValidateKnowledge => RunValidateKnowledge(host.Services, parsed),
                AgentCommandKind.BuildKnowledge => RunBuildKnowledge(host.Services, parsed),
                AgentCommandKind.CompareKnowledge => RunCompareKnowledge(host.Services, parsed),
                AgentCommandKind.SearchKnowledge => RunSearchKnowledge(host.Services, parsed),
                AgentCommandKind.ApproveKnowledge => RunApproveKnowledge(host.Services, parsed),
                AgentCommandKind.ValidateTestCase => RunValidateTestCase(host.Services, parsed),
                AgentCommandKind.ValidateStepRegistry => RunValidateStepRegistry(host.Services, parsed),
                AgentCommandKind.CompileKey => parsed.FromJira
                    ? await RunCompileKeyFromJiraAsync(host.Services, parsed, cancellationToken)
                        .ConfigureAwait(false)
                    : RunCompileKey(host.Services, parsed),
                AgentCommandKind.ValidateCompiledKey => RunValidateCompiledKey(host.Services, parsed),
                AgentCommandKind.SyncKey => await RunSyncKeyAsync(host.Services, parsed, cancellationToken)
                    .ConfigureAwait(false),
                AgentCommandKind.RunKey => await RunRunKeyAsync(host.Services, parsed, cancellationToken)
                    .ConfigureAwait(false),
                AgentCommandKind.RunSuite => await RunRunSuiteAsync(host.Services, parsed, cancellationToken)
                    .ConfigureAwait(false),
                AgentCommandKind.ProposePlan => await RunProposePlanAsync(host.Services, parsed, cancellationToken)
                    .ConfigureAwait(false),
                _ => ExitCodes.UsageOrConfiguration
            };
        }
        catch (AgentConfigurationException ex)
        {
            if (jsonStdoutMode)
            {
                WriteJsonError(ExitCodes.UsageOrConfiguration, ex.Message);
            }
            else
            {
                await Console.Error.WriteLineAsync(SecretRedactor.Redact(ex.Message)).ConfigureAwait(false);
            }

            return ExitCodes.UsageOrConfiguration;
        }
        catch (WorkspaceException ex)
        {
            if (jsonStdoutMode)
            {
                WriteJsonError(ExitCodes.SuiteOrWorkspace, ex.Message);
            }
            else
            {
                await Console.Error.WriteLineAsync(ex.Message).ConfigureAwait(false);
            }

            return ExitCodes.SuiteOrWorkspace;
        }
    }

    private static void WriteJsonError(int exitCode, string message)
    {
        var payload = new
        {
            success = false,
            exitCode,
            error = SecretRedactor.Redact(message)
        };
        Console.WriteLine(JsonSerializer.Serialize(payload, JsonOutputOptions));
    }

    internal static IHost BuildHost(string[] configurationArgs, bool jsonStdoutMode = false)
    {
        var builder = Host.CreateApplicationBuilder(new HostApplicationBuilderSettings
        {
            Args = configurationArgs,
            ContentRootPath = Directory.GetCurrentDirectory()
        });

        AgentServiceRegistrar.ConfigureAgentConfiguration(builder, configurationArgs);
        // ProcessIdentityValidator is crawl-scoped (created per ExploreAsync), not a singleton.
        AgentServiceRegistrar.AddAgentCoreServices(builder, quietAgentLogs: jsonStdoutMode);

        return builder.Build();
    }

    private static int RunInit(IServiceProvider services)
    {
        AgentOptionsValidator.Validate(services.GetRequiredService<IOptions<AgentOptions>>().Value, OptionsValidationScope.Workspace);
        var workspace = services.GetRequiredService<IWorkspaceManager>();
        var result = workspace.Initialize();

        Console.WriteLine($"Workspace root: {result.RootPath}");
        Console.WriteLine($"Created: {result.CreatedPaths.Count}");
        foreach (var path in result.CreatedPaths)
            Console.WriteLine($"  + {path}");
        Console.WriteLine($"Skipped existing: {result.SkippedExistingPaths.Count}");
        foreach (var path in result.SkippedExistingPaths)
            Console.WriteLine($"  = {path}");

        return ExitCodes.Success;
    }

    private static int RunValidateSuite(IServiceProvider services, string file)
    {
        AgentOptionsValidator.Validate(services.GetRequiredService<IOptions<AgentOptions>>().Value, OptionsValidationScope.Suites);
        var reader = services.GetRequiredService<ISuiteManifestReader>();
        var result = reader.ValidateFile(file);

        Console.WriteLine($"Suite file : {result.FilePath}");
        Console.WriteLine($"Suite name : {result.SuiteName}");
        Console.WriteLine($"Enabled    : {result.SuiteEnabled}");
        Console.WriteLine($"Total      : {result.TotalCount}");
        Console.WriteLine($"Enabled TC : {result.EnabledCount}");
        Console.WriteLine($"Disabled   : {result.DisabledCount}");
        Console.WriteLine($"Duplicates : {result.DuplicateCount}");

        if (!result.IsValid)
        {
            Console.Error.WriteLine("Validation failed:");
            foreach (var error in result.Errors)
                Console.Error.WriteLine($"  - {error}");
            return ExitCodes.SuiteOrWorkspace;
        }

        Console.WriteLine("Suite validation succeeded.");
        return ExitCodes.Success;
    }

    private static int RunValidateKeys(IServiceProvider services, IReadOnlyList<string> keys)
    {
        AgentOptionsValidator.Validate(services.GetRequiredService<IOptions<AgentOptions>>().Value, OptionsValidationScope.Suites);
        var reader = services.GetRequiredService<ISuiteManifestReader>();
        var result = reader.ValidateKeys(keys);

        Console.WriteLine($"Valid keys: {result.ValidKeys.Count}");
        foreach (var key in result.ValidKeys)
            Console.WriteLine($"  - {key}");

        if (!result.IsValid)
        {
            Console.Error.WriteLine("Key validation failed:");
            foreach (var error in result.Errors)
                Console.Error.WriteLine($"  - {error}");
            return ExitCodes.SuiteOrWorkspace;
        }

        Console.WriteLine("Key validation succeeded.");
        return ExitCodes.Success;
    }

    private static int RunValidatePlan(IServiceProvider services, string file, bool json)
    {
        AgentOptionsValidator.Validate(
            services.GetRequiredService<IOptions<AgentOptions>>().Value,
            OptionsValidationScope.Runner);

        var reader = services.GetRequiredService<PlanManifestReader>();
        var integrator = services.GetRequiredService<PlanObjectRepositoryIntegrator>();
        var result = integrator.Integrate(reader.Read(file));

        if (json)
        {
            var payload = new
            {
                result.IsValid,
                result.PlanPath,
                result.PlanId,
                result.Name,
                result.StepCount,
                result.OnFailureStepCount,
                result.TotalStepCount,
                result.Sha256,
                result.ObjectRepositoryPath,
                result.ObjectRepositoryId,
                result.ObjectRepositorySha256,
                result.ResolvedObjectReferences,
                result.Errors,
                result.Warnings
            };
            Console.WriteLine(SecretRedactor.Redact(JsonSerializer.Serialize(payload, JsonOutputOptions)));
        }
        else
        {
            Console.WriteLine($"Plan file : {result.PlanPath}");
            Console.WriteLine($"Plan ID   : {result.PlanId}");
            Console.WriteLine($"Name      : {result.Name}");
            Console.WriteLine($"Steps     : {result.StepCount}");
            Console.WriteLine($"OnFailure : {result.OnFailureStepCount}");
            if (!string.IsNullOrWhiteSpace(result.Sha256))
                Console.WriteLine($"SHA-256   : {result.Sha256}");

            if (!result.IsValid)
            {
                Console.Error.WriteLine("Plan validation failed:");
                foreach (var error in result.Errors)
                    Console.Error.WriteLine($"  - {error}");
            }
            else
            {
                Console.WriteLine("Offline plan validation succeeded.");
                Console.WriteLine("Note: operation support requires a live catalog (use run-plan --dry-run).");
            }
        }

        return result.IsValid ? ExitCodes.Success : ExitCodes.SuiteOrWorkspace;
    }

    private static async Task<int> RunPlanAsync(
        IServiceProvider services,
        string file,
        bool dryRun,
        bool json,
        CancellationToken cancellationToken)
    {
        AgentOptionsValidator.Validate(
            services.GetRequiredService<IOptions<AgentOptions>>().Value,
            OptionsValidationScope.Runner);

        var runner = services.GetRequiredService<IDeterministicPlanRunner>();
        var report = await runner.RunAsync(file, dryRun, cancellationToken).ConfigureAwait(false);

        if (json)
        {
            var redacted = RunArtifactWriter.RedactReport(report);
            Console.WriteLine(JsonSerializer.Serialize(redacted, JsonOutputOptions));
        }
        else
        {
            Console.WriteLine(dryRun ? "Desktop Automation Agent — run-plan (dry-run)" : "Desktop Automation Agent — run-plan");
            Console.WriteLine($"Run ID     : {report.RunId}");
            Console.WriteLine($"Status     : {report.Status}");
            Console.WriteLine($"Exit code  : {report.ExitCode}");
            Console.WriteLine($"Plan       : {report.PlanPath}");
            Console.WriteLine($"Plan ID    : {report.PlanId}");
            if (!string.IsNullOrWhiteSpace(report.PlanSha256))
                Console.WriteLine($"Plan SHA   : {report.PlanSha256}");
            if (!string.IsNullOrWhiteSpace(report.DriverBaseUrl))
                Console.WriteLine($"Driver     : {report.DriverBaseUrl}");
            Console.WriteLine($"Steps      : {report.Steps.Count}");
            Console.WriteLine($"Cleanup    : {report.OnFailureSteps.Count}");
            Console.WriteLine($"Artifacts  : {report.ArtifactWriteStatus}");

            if (report.Failure is not null)
            {
                Console.Error.WriteLine(
                    $"Failure [{report.Failure.Classification}]: {SecretRedactor.Redact(report.Failure.Message)}");
            }
        }

        return report.ExitCode;
    }

    private static int RunValidateObjectRepository(IServiceProvider services, string file, bool json)
    {
        AgentOptionsValidator.Validate(
            services.GetRequiredService<IOptions<AgentOptions>>().Value,
            OptionsValidationScope.ObjectRepository);

        var reader = services.GetRequiredService<ObjectRepositoryReader>();
        var result = reader.Read(file);

        if (json)
        {
            var payload = new
            {
                result.IsValid,
                result.RepositoryPath,
                result.ManifestSha256,
                result.AggregateSha256,
                result.Errors,
                result.Warnings
            };
            Console.WriteLine(SecretRedactor.Redact(JsonSerializer.Serialize(payload, JsonOutputOptions)));
        }
        else
        {
            Console.WriteLine($"Repository : {result.RepositoryPath}");
            if (!string.IsNullOrWhiteSpace(result.AggregateSha256))
                Console.WriteLine($"SHA-256    : {result.AggregateSha256}");

            if (!result.IsValid)
            {
                Console.Error.WriteLine("Object repository validation failed:");
                foreach (var error in result.Errors)
                    Console.Error.WriteLine($"  - {SecretRedactor.Redact(error)}");
            }
            else
            {
                Console.WriteLine("Object repository validation succeeded.");
            }

            foreach (var warning in result.Warnings)
                Console.Error.WriteLine($"Warning: {SecretRedactor.Redact(warning)}");
        }

        return result.IsValid ? ExitCodes.Success : ExitCodes.SuiteOrWorkspace;
    }

    private static int RunResolveObject(IServiceProvider services, string file, string objectRef, bool json)
    {
        AgentOptionsValidator.Validate(
            services.GetRequiredService<IOptions<AgentOptions>>().Value,
            OptionsValidationScope.ObjectRepository);

        var reader = services.GetRequiredService<ObjectRepositoryReader>();
        var resolver = services.GetRequiredService<ObjectReferenceResolver>();
        var validation = reader.Read(file);
        if (!validation.IsValid || validation.Snapshot is null)
        {
            if (json)
            {
                WriteJsonError(ExitCodes.SuiteOrWorkspace, validation.Errors.FirstOrDefault() ?? "Repository validation failed.");
            }
            else
            {
                foreach (var error in validation.Errors)
                    Console.Error.WriteLine(SecretRedactor.Redact(error));
            }

            return ExitCodes.SuiteOrWorkspace;
        }

        var resolution = resolver.Resolve(validation.Snapshot, objectRef);
        if (json)
        {
            var payload = new
            {
                success = resolution.IsResolved,
                reference = resolution.Reference,
                repositoryId = validation.Snapshot.Manifest.RepositoryId,
                repositorySha256 = validation.Snapshot.AggregateSha256,
                pageId = resolution.PageId,
                elementId = resolution.ElementId,
                locator = resolution.Locator is null
                    ? (JsonElement?)null
                    : ObjectLocatorSerializer.ToJsonElement(resolution.Locator),
                warnings = resolution.Warnings,
                errors = resolution.Errors,
                repositoryPath = validation.RepositoryPath
            };
            Console.WriteLine(SecretRedactor.Redact(JsonSerializer.Serialize(payload, JsonOutputOptions)));
        }
        else
        {
            Console.WriteLine($"Reference  : {resolution.Reference}");
            if (resolution.Locator is not null)
                Console.WriteLine($"Locator    : {JsonSerializer.Serialize(resolution.Locator, JsonOutputOptions)}");

            foreach (var warning in resolution.Warnings)
                Console.Error.WriteLine($"Warning: {SecretRedactor.Redact(warning)}");

            if (!resolution.IsResolved)
            {
                foreach (var error in resolution.Errors)
                    Console.Error.WriteLine(SecretRedactor.Redact(error));
            }
        }

        return resolution.IsResolved ? ExitCodes.Success : ExitCodes.SuiteOrWorkspace;
    }

    private static async Task<int> RunCapturePageAsync(
        IServiceProvider services,
        ParsedCommand parsed,
        CancellationToken cancellationToken)
    {
        var service = services.GetRequiredService<ObjectCaptureService>();
        var result = await service.CaptureAsync(
            parsed.RepositoryFile!,
            parsed.PageId!,
            parsed.PageName!,
            new ObjectCaptureOptions
            {
                View = parsed.View ?? "control",
                Root = parsed.Root ?? "activeWindow",
                MaxDepth = parsed.MaxDepth,
                MaxChildren = parsed.MaxChildren,
                IncludeOffscreen = parsed.IncludeOffscreen
            },
            cancellationToken).ConfigureAwait(false);

        if (parsed.Json)
        {
            Console.WriteLine(JsonSerializer.Serialize(result, JsonOutputOptions));
        }
        else
        {
            Console.WriteLine("Desktop Automation Agent — capture-page");
            Console.WriteLine($"Success    : {result.Success}");
            Console.WriteLine($"Capture ID : {result.CaptureId}");
            Console.WriteLine($"Nodes      : {result.NodeCount}");
            Console.WriteLine($"Elements   : {result.ElementCount}");
            Console.WriteLine($"Unresolved : {result.UnresolvedCount}");
            if (!string.IsNullOrWhiteSpace(result.CaptureFilePath))
                Console.WriteLine($"Capture    : {result.CaptureFilePath}");
            if (!string.IsNullOrWhiteSpace(result.CandidateFilePath))
                Console.WriteLine($"Candidate  : {result.CandidateFilePath}");
            if (!string.IsNullOrWhiteSpace(result.Error))
                Console.Error.WriteLine(result.Error);
        }

        return result.ExitCode;
    }

    private static async Task<int> RunVerifyObjectRepositoryAsync(
        IServiceProvider services,
        ParsedCommand parsed,
        CancellationToken cancellationToken)
    {
        var service = services.GetRequiredService<ObjectVerificationService>();
        var result = await service.VerifyAsync(
            parsed.RepositoryFile!,
            new ObjectVerificationOptions
            {
                PageId = parsed.PageId,
                ObjectRef = parsed.ObjectRef,
                View = parsed.View ?? "control",
                Root = parsed.Root ?? "activeWindow",
                MaxDepth = parsed.MaxDepth,
                MaxChildren = parsed.MaxChildren,
                IncludeOffscreen = parsed.IncludeOffscreen
            },
            cancellationToken).ConfigureAwait(false);

        if (parsed.Json)
        {
            Console.WriteLine(JsonSerializer.Serialize(result, JsonOutputOptions));
        }
        else
        {
            Console.WriteLine("Desktop Automation Agent — verify-object-repository");
            Console.WriteLine($"Success    : {result.Success}");
            Console.WriteLine($"Total      : {result.Total}");
            Console.WriteLine($"Passed     : {result.Passed}");
            Console.WriteLine($"Missing    : {result.Missing}");
            Console.WriteLine($"Ambiguous  : {result.Ambiguous}");
            Console.WriteLine($"Fragile    : {result.Fragile}");
            Console.WriteLine($"Failed     : {result.Failed}");
            if (!string.IsNullOrWhiteSpace(result.Error))
                Console.Error.WriteLine(result.Error);
        }

        return result.ExitCode;
    }

    private static async Task<int> RunObserveAppAsync(
        IServiceProvider services,
        ParsedCommand parsed,
        CancellationToken cancellationToken)
    {
        var service = services.GetRequiredService<ISemanticObservationService>();
        var result = await service.ObserveAsync(
            new ObservationOptions
            {
                View = parsed.View ?? "control",
                Root = parsed.Root ?? "activeWindow",
                MaxDepth = parsed.MaxDepth,
                MaxChildren = parsed.MaxChildren,
                MaxNodes = parsed.MaxNodes,
                IncludeOffscreen = parsed.IncludeOffscreen,
                IncludeBounds = parsed.IncludeBounds,
                IncludeHelpText = parsed.IncludeHelpText,
                IncludeFocusState = parsed.IncludeFocusState ?? true,
                ActionableOnly = parsed.ActionableOnly ?? true,
                OutputPath = parsed.OutputPath
            },
            cancellationToken).ConfigureAwait(false);

        if (parsed.Json)
        {
            // Structured JSON on stdout. Runtime IDs stay in-memory only and are omitted from
            // machine-readable stdout to match the persisted artifact privacy rules.
            var jsonPayload = result.Snapshot is null
                ? result
                : result with
                {
                    Snapshot = new SemanticSnapshot
                    {
                        SnapshotSchemaVersion = result.Snapshot.SnapshotSchemaVersion,
                        SnapshotId = result.Snapshot.SnapshotId,
                        CapturedAtUtc = result.Snapshot.CapturedAtUtc,
                        DriverVersion = result.Snapshot.DriverVersion,
                        SnapshotHash = result.Snapshot.SnapshotHash,
                        Application = result.Snapshot.Application,
                        WindowsSession = result.Snapshot.WindowsSession,
                        Window = result.Snapshot.Window,
                        Nodes = result.Snapshot.Nodes.Select(static n => new SemanticNode
                        {
                            Ref = n.Ref,
                            ParentRef = n.ParentRef,
                            Ordinal = n.Ordinal,
                            Depth = n.Depth,
                            ControlType = n.ControlType,
                            Role = n.Role,
                            Name = n.Name,
                            AutomationId = n.AutomationId,
                            ClassName = n.ClassName,
                            HelpText = n.HelpText,
                            Enabled = n.Enabled,
                            Offscreen = n.Offscreen,
                            Focusable = n.Focusable,
                            HasKeyboardFocus = n.HasKeyboardFocus,
                            Patterns = n.Patterns,
                            Capabilities = n.Capabilities,
                            Bounds = n.Bounds
                        }).ToList(),
                        Truncated = result.Snapshot.Truncated,
                        Warnings = result.Snapshot.Warnings,
                        Ordering = result.Snapshot.Ordering
                    }
                };
            Console.WriteLine(JsonSerializer.Serialize(jsonPayload, JsonOutputOptions));
        }
        else
        {
            Console.WriteLine("Desktop Automation Agent — observe-app");
            Console.WriteLine($"Success         : {result.Success}");
            Console.WriteLine($"Observation ID  : {result.ObservationId}");
            Console.WriteLine($"Nodes           : {result.NodeCount}");
            Console.WriteLine($"Truncated       : {result.Truncated}");
            Console.WriteLine($"Snapshot hash   : {result.SnapshotHash}");
            if (!string.IsNullOrWhiteSpace(result.SnapshotFilePath))
                Console.WriteLine($"Artifact        : {result.SnapshotFilePath}");
            if (!string.IsNullOrWhiteSpace(result.Error))
                Console.Error.WriteLine(result.Error);
        }

        return result.ExitCode;
    }

    private static async Task<int> RunExploreAppAsync(
        IServiceProvider services,
        ParsedCommand parsed,
        CancellationToken cancellationToken)
    {
        var coordinator = services.GetRequiredService<IExplorationCoordinator>();
        var result = await coordinator.ExploreAsync(
            new ExplorationRequest
            {
                PolicyPath = parsed.PolicyFile!,
                DryRun = parsed.DryRun
            },
            cancellationToken).ConfigureAwait(false);

        if (parsed.Json)
        {
            var payload = new
            {
                result.Success,
                result.ExitCode,
                result.Error,
                result.ExplorationId,
                result.ArtifactDirectory,
                summary = result.Summary,
                dryRunCandidates = result.DryRunCandidates.Select(c => new
                {
                    c.ActionKey,
                    kind = c.Kind.ToString(),
                    c.Category,
                    c.DriverOperation,
                    c.ElementKey,
                    c.AutomationId,
                    c.Name,
                    c.ControlType,
                    c.Depth,
                    c.Priority
                })
            };
            Console.WriteLine(SecretRedactor.Redact(JsonSerializer.Serialize(payload, JsonOutputOptions)));
        }
        else
        {
            Console.WriteLine("Desktop Automation Agent — explore-app");
            Console.WriteLine($"Success           : {result.Success}");
            Console.WriteLine($"Exit code         : {result.ExitCode}");
            Console.WriteLine($"Exploration id    : {result.ExplorationId}");
            if (result.Summary is not null)
            {
                Console.WriteLine($"Status            : {result.Summary.Status}");
                Console.WriteLine($"Mode              : {result.Summary.Mode}");
                Console.WriteLine($"Dry run           : {result.Summary.DryRun}");
                Console.WriteLine($"States visited    : {result.Summary.StatesVisited}");
                Console.WriteLine($"Actions attempted : {result.Summary.ActionsAttempted}");
                Console.WriteLine($"Transitions       : {result.Summary.TransitionsRecorded}");
                Console.WriteLine($"Max depth         : {result.Summary.MaxDepthReached}");
                Console.WriteLine($"Stop reason       : {result.Summary.StopReason}");
            }

            if (!string.IsNullOrWhiteSpace(result.ArtifactDirectory))
                Console.WriteLine($"Artifacts         : {result.ArtifactDirectory}");

            if (parsed.DryRun && result.DryRunCandidates.Count > 0)
            {
                Console.WriteLine("Permitted candidates:");
                foreach (var candidate in result.DryRunCandidates.Take(50))
                {
                    Console.WriteLine(
                        $"  - [{candidate.Category}] {candidate.Kind} {candidate.ElementKey} via {candidate.DriverOperation}");
                }
            }

            if (!string.IsNullOrWhiteSpace(result.Error))
                Console.Error.WriteLine(SecretRedactor.Redact(result.Error));
        }

        return result.ExitCode;
    }

    private static int RunValidateKnowledge(IServiceProvider services, ParsedCommand parsed)
    {
        AgentOptionsValidator.Validate(
            services.GetRequiredService<IOptions<AgentOptions>>().Value,
            OptionsValidationScope.Knowledge);

        var knowledge = services.GetRequiredService<KnowledgeService>();
        var result = knowledge.ValidatePath(parsed.KnowledgePath!);

        if (parsed.Json)
        {
            var payload = new
            {
                success = result.IsValid,
                exitCode = result.IsValid ? ExitCodes.Success : ExitCodes.SuiteOrWorkspace,
                errors = result.Errors,
                warnings = result.Warnings
            };
            Console.WriteLine(SecretRedactor.Redact(JsonSerializer.Serialize(payload, JsonOutputOptions)));
        }
        else
        {
            Console.WriteLine("Desktop Automation Agent — validate-knowledge");
            Console.WriteLine($"Valid    : {result.IsValid}");
            foreach (var error in result.Errors)
                Console.Error.WriteLine($"  - {SecretRedactor.Redact(error)}");
            foreach (var warning in result.Warnings)
                Console.Error.WriteLine($"Warning: {SecretRedactor.Redact(warning)}");
        }

        return result.IsValid ? ExitCodes.Success : ExitCodes.SuiteOrWorkspace;
    }

    private static int RunBuildKnowledge(IServiceProvider services, ParsedCommand parsed)
    {
        AgentOptionsValidator.Validate(
            services.GetRequiredService<IOptions<AgentOptions>>().Value,
            OptionsValidationScope.Knowledge);

        var knowledge = services.GetRequiredService<KnowledgeService>();
        var result = knowledge.Build(
            parsed.ApplicationId!,
            parsed.ExplorationRuns,
            parsed.SuggestPageObjects,
            parsed.OutputPath);

        if (parsed.Json)
        {
            var payload = new
            {
                success = result.Success,
                exitCode = result.ExitCode,
                error = result.Error,
                outputDirectory = result.OutputDirectory,
                knowledgeId = result.Knowledge?.KnowledgeId,
                buildFingerprint = result.Knowledge?.BuildFingerprint,
                contentHash = result.ContentHash,
                approvalTemplate = result.Success
                    ? Path.Combine(result.OutputDirectory ?? string.Empty, "approval-template.json")
                    : null,
                coverage = result.Knowledge?.Coverage,
                conflicts = result.Knowledge?.Coverage.Conflicts,
                suggestedPageObjectPaths = result.SuggestedPageObjectPaths,
                warnings = result.Knowledge?.Warnings
            };
            Console.WriteLine(SecretRedactor.Redact(JsonSerializer.Serialize(payload, JsonOutputOptions)));
        }
        else
        {
            Console.WriteLine("Desktop Automation Agent — build-knowledge");
            Console.WriteLine($"Success           : {result.Success}");
            Console.WriteLine($"Exit code         : {result.ExitCode}");
            if (result.Knowledge is not null)
            {
                Console.WriteLine($"Knowledge id      : {result.Knowledge.KnowledgeId}");
                Console.WriteLine($"Build fingerprint : {result.Knowledge.BuildFingerprint}");
                Console.WriteLine($"Content hash      : {result.ContentHash}");
                Console.WriteLine($"States            : {result.Knowledge.Coverage.UniqueStates}");
                Console.WriteLine($"Transitions       : {result.Knowledge.Coverage.PermittedTransitionsAttempted}");
                Console.WriteLine($"Conflicts         : {result.Knowledge.Coverage.Conflicts}");
            }

            if (!string.IsNullOrWhiteSpace(result.OutputDirectory))
            {
                Console.WriteLine($"Output            : {result.OutputDirectory}");
                Console.WriteLine($"Approval template : {Path.Combine(result.OutputDirectory, "approval-template.json")}");
            }

            foreach (var path in result.SuggestedPageObjectPaths)
                Console.WriteLine($"Candidate PO      : {path}");
            if (!string.IsNullOrWhiteSpace(result.Error))
                Console.Error.WriteLine(SecretRedactor.Redact(result.Error));
        }

        return result.ExitCode;
    }

    private static int RunCompareKnowledge(IServiceProvider services, ParsedCommand parsed)
    {
        AgentOptionsValidator.Validate(
            services.GetRequiredService<IOptions<AgentOptions>>().Value,
            OptionsValidationScope.Knowledge);

        try
        {
            var knowledge = services.GetRequiredService<KnowledgeService>();
            var result = knowledge.Compare(parsed.KnowledgeLeftPath!, parsed.KnowledgeRightPath!);

            if (parsed.Json)
            {
                Console.WriteLine(SecretRedactor.Redact(JsonSerializer.Serialize(result, JsonOutputOptions)));
            }
            else
            {
                Console.WriteLine("Desktop Automation Agent — compare-knowledge");
                Console.WriteLine($"Compatible fingerprints : {result.CompatibleFingerprints}");
                Console.WriteLine($"Graphs equal            : {result.GraphsEqual}");
                Console.WriteLine($"Left fingerprint        : {result.LeftBuildFingerprint}");
                Console.WriteLine($"Right fingerprint       : {result.RightBuildFingerprint}");
                foreach (var diff in result.Differences)
                    Console.WriteLine($"  - {diff}");
            }

            if (!result.CompatibleFingerprints)
                return ExitCodes.SuiteOrWorkspace;
            return ExitCodes.Success;
        }
        catch (Exception ex) when (ex is FileNotFoundException or InvalidOperationException or KnowledgePathException or IOException)
        {
            if (parsed.Json)
                WriteJsonError(ExitCodes.SuiteOrWorkspace, ex.Message);
            else
                Console.Error.WriteLine(SecretRedactor.Redact(ex.Message));
            return ExitCodes.SuiteOrWorkspace;
        }
    }

    private static int RunSearchKnowledge(IServiceProvider services, ParsedCommand parsed)
    {
        AgentOptionsValidator.Validate(
            services.GetRequiredService<IOptions<AgentOptions>>().Value,
            OptionsValidationScope.Knowledge);

        try
        {
            var knowledge = services.GetRequiredService<KnowledgeService>();
            var hits = knowledge.Search(parsed.KnowledgePath!, parsed.KnowledgeQuery!);

            if (parsed.Json)
            {
                var payload = new { success = true, query = parsed.KnowledgeQuery, hits };
                Console.WriteLine(SecretRedactor.Redact(JsonSerializer.Serialize(payload, JsonOutputOptions)));
            }
            else
            {
                Console.WriteLine("Desktop Automation Agent — search-knowledge");
                Console.WriteLine($"Query : {parsed.KnowledgeQuery}");
                Console.WriteLine($"Hits  : {hits.Count}");
                foreach (var hit in hits)
                    Console.WriteLine($"  [{hit.Score}] {hit.Kind} {hit.Id}: {hit.Text}");
            }

            return ExitCodes.Success;
        }
        catch (Exception ex) when (ex is FileNotFoundException or InvalidOperationException or KnowledgePathException or IOException or ArgumentException)
        {
            if (parsed.Json)
                WriteJsonError(ExitCodes.SuiteOrWorkspace, ex.Message);
            else
                Console.Error.WriteLine(SecretRedactor.Redact(ex.Message));
            return ExitCodes.SuiteOrWorkspace;
        }
    }

    private static int RunApproveKnowledge(IServiceProvider services, ParsedCommand parsed)
    {
        AgentOptionsValidator.Validate(
            services.GetRequiredService<IOptions<AgentOptions>>().Value,
            OptionsValidationScope.Knowledge);

        var knowledge = services.GetRequiredService<KnowledgeService>();
        var result = knowledge.Approve(
            parsed.KnowledgeSourcePath!,
            parsed.Reviewer!,
            parsed.ReviewedInputPath!);

        if (parsed.Json)
        {
            var payload = new
            {
                success = result.Success,
                exitCode = result.ExitCode,
                error = result.Error,
                outputDirectory = result.OutputDirectory,
                knowledgeId = result.Knowledge?.KnowledgeId,
                status = result.Knowledge?.Status,
                promotesPageObjects = result.Knowledge?.Approval?.PromotesPageObjects
            };
            Console.WriteLine(SecretRedactor.Redact(JsonSerializer.Serialize(payload, JsonOutputOptions)));
        }
        else
        {
            Console.WriteLine("Desktop Automation Agent — approve-knowledge");
            Console.WriteLine($"Success : {result.Success}");
            if (!string.IsNullOrWhiteSpace(result.OutputDirectory))
                Console.WriteLine($"Output  : {result.OutputDirectory}");
            Console.WriteLine("Note    : Knowledge approval never promotes Page Objects.");
            if (!string.IsNullOrWhiteSpace(result.Error))
                Console.Error.WriteLine(SecretRedactor.Redact(result.Error));
        }

        return result.ExitCode;
    }

    private static int RunValidateTestCase(IServiceProvider services, ParsedCommand parsed)
    {
        AgentOptionsValidator.Validate(
            services.GetRequiredService<IOptions<AgentOptions>>().Value,
            OptionsValidationScope.Workspace);

        var service = services.GetRequiredService<BddCompilationService>();
        var result = service.ValidateTestCase(parsed.TestCaseFile!);

        if (parsed.Json)
        {
            var payload = new
            {
                success = result.IsValid,
                exitCode = result.IsValid ? ExitCodes.Success : ExitCodes.SuiteOrWorkspace,
                path = result.Path,
                sha256 = result.Sha256,
                jiraKey = result.Document?.JiraKey,
                errors = result.Errors
            };
            Console.WriteLine(SecretRedactor.Redact(JsonSerializer.Serialize(payload, JsonOutputOptions)));
        }
        else
        {
            Console.WriteLine("Desktop Automation Agent — validate-test-case");
            Console.WriteLine($"Path    : {result.Path}");
            if (!string.IsNullOrWhiteSpace(result.Sha256))
                Console.WriteLine($"SHA-256 : {result.Sha256}");
            if (!result.IsValid)
            {
                Console.Error.WriteLine("Validation failed:");
                foreach (var error in result.Errors)
                    Console.Error.WriteLine($"  - {SecretRedactor.Redact(error)}");
            }
            else
            {
                Console.WriteLine("Test-case validation succeeded.");
            }
        }

        return result.IsValid ? ExitCodes.Success : ExitCodes.SuiteOrWorkspace;
    }

    private static int RunValidateStepRegistry(IServiceProvider services, ParsedCommand parsed)
    {
        AgentOptionsValidator.Validate(
            services.GetRequiredService<IOptions<AgentOptions>>().Value,
            OptionsValidationScope.Workspace);

        var service = services.GetRequiredService<BddCompilationService>();
        var result = service.ValidateStepRegistry(parsed.StepRegistryFile);

        if (parsed.Json)
        {
            var payload = new
            {
                success = result.IsValid,
                exitCode = result.IsValid ? ExitCodes.Success : ExitCodes.SuiteOrWorkspace,
                path = result.Path,
                sha256 = result.Sha256,
                registryId = result.Document?.RegistryId,
                mappingCount = result.Document?.Mappings.Count,
                errors = result.Errors
            };
            Console.WriteLine(SecretRedactor.Redact(JsonSerializer.Serialize(payload, JsonOutputOptions)));
        }
        else
        {
            Console.WriteLine("Desktop Automation Agent — validate-step-registry");
            Console.WriteLine($"Path    : {result.Path}");
            if (!string.IsNullOrWhiteSpace(result.Sha256))
                Console.WriteLine($"SHA-256 : {result.Sha256}");
            if (!result.IsValid)
            {
                Console.Error.WriteLine("Validation failed:");
                foreach (var error in result.Errors)
                    Console.Error.WriteLine($"  - {SecretRedactor.Redact(error)}");
            }
            else
            {
                Console.WriteLine($"Step registry validation succeeded ({result.Document!.Mappings.Count} mappings).");
            }
        }

        return result.IsValid ? ExitCodes.Success : ExitCodes.SuiteOrWorkspace;
    }

    private static int RunCompileKey(IServiceProvider services, ParsedCommand parsed)
    {
        AgentOptionsValidator.Validate(
            services.GetRequiredService<IOptions<AgentOptions>>().Value,
            OptionsValidationScope.Workspace);

        var service = services.GetRequiredService<BddCompilationService>();
        CompilationResult result;
        try
        {
            result = service.CompileKey(new CompileKeyRequest
            {
                JiraKey = parsed.CompileJiraKey!,
                TestCasePath = parsed.TestCaseFile,
                RegistryPath = parsed.StepRegistryFile,
                ObjectRepositoryPath = parsed.RepositoryFile,
                KnowledgePath = parsed.KnowledgePath,
                CatalogPath = parsed.CatalogFile,
                Explain = parsed.Explain
            });
        }
        catch (WorkspaceException ex)
        {
            if (parsed.Json)
                WriteJsonError(ExitCodes.SuiteOrWorkspace, ex.Message);
            else
                Console.Error.WriteLine(SecretRedactor.Redact(ex.Message));
            return ExitCodes.SuiteOrWorkspace;
        }

        return WriteCompileResult(result, parsed, fromJira: false);
    }

    private static async Task<int> RunCompileKeyFromJiraAsync(
        IServiceProvider services,
        ParsedCommand parsed,
        CancellationToken cancellationToken)
    {
        AgentOptionsValidator.Validate(
            services.GetRequiredService<IOptions<AgentOptions>>().Value,
            OptionsValidationScope.Jira);

        var orchestrator = services.GetRequiredService<JiraSuiteOrchestrator>();
        CompilationResult result;
        try
        {
            result = await orchestrator.CompileKeyFromJiraAsync(
                    new CompileKeyFromJiraRequest
                    {
                        JiraKey = parsed.CompileJiraKey!,
                        RegistryPath = parsed.StepRegistryFile,
                        ObjectRepositoryPath = parsed.RepositoryFile,
                        KnowledgePath = parsed.KnowledgePath,
                        CatalogPath = parsed.CatalogFile,
                        Explain = parsed.Explain
                    },
                    cancellationToken)
                .ConfigureAwait(false);
        }
        catch (WorkspaceException ex)
        {
            if (parsed.Json)
                WriteJsonError(ExitCodes.SuiteOrWorkspace, ex.Message);
            else
                Console.Error.WriteLine(SecretRedactor.Redact(ex.Message));
            return ExitCodes.SuiteOrWorkspace;
        }

        return WriteCompileResult(result, parsed, fromJira: true);
    }

    private static int WriteCompileResult(CompilationResult result, ParsedCommand parsed, bool fromJira)
    {
        if (parsed.Json)
        {
            var payload = new
            {
                success = result.Success,
                exitCode = result.ExitCode == 0
                    ? (result.Success ? ExitCodes.Success : ExitCodes.SuiteOrWorkspace)
                    : result.ExitCode,
                jiraKey = result.JiraKey,
                fingerprint = result.Fingerprint,
                cacheHit = result.CacheHit,
                fromJira,
                planPath = result.PlanPath,
                compilationPath = result.CompilationPath,
                planSha256 = result.PlanSha256,
                errors = result.Errors,
                warnings = result.Warnings,
                explain = result.Explain
            };
            Console.WriteLine(SecretRedactor.Redact(JsonSerializer.Serialize(payload, JsonOutputOptions)));
        }
        else
        {
            Console.WriteLine(fromJira
                ? "Desktop Automation Agent — compile-key --from-jira"
                : "Desktop Automation Agent — compile-key");
            Console.WriteLine($"Jira key     : {result.JiraKey}");
            Console.WriteLine($"Success      : {result.Success}");
            if (!string.IsNullOrWhiteSpace(result.Fingerprint))
                Console.WriteLine($"Fingerprint  : {result.Fingerprint}");
            Console.WriteLine($"Cache hit    : {result.CacheHit}");
            if (!string.IsNullOrWhiteSpace(result.PlanPath))
                Console.WriteLine($"Plan         : {result.PlanPath}");
            if (!string.IsNullOrWhiteSpace(result.CompilationPath))
                Console.WriteLine($"Compilation  : {result.CompilationPath}");
            if (!string.IsNullOrWhiteSpace(result.PlanSha256))
                Console.WriteLine($"Plan SHA-256 : {result.PlanSha256}");
            Console.WriteLine(fromJira
                ? "Note         : --from-jira is read-only; never writes to Jira or calls POST /ui."
                : "Note         : compile-key never calls Jira, AI, or POST /ui.");
            if (result.Explain is { Count: > 0 })
            {
                Console.WriteLine("Explain:");
                foreach (var entry in result.Explain)
                {
                    Console.WriteLine(
                        $"  [{entry.Ordinal}] {entry.Keyword} {entry.Text} -> {entry.Outcome}" +
                        (string.IsNullOrWhiteSpace(entry.MappingId) ? string.Empty : $" ({entry.MappingId})"));
                    foreach (var detail in entry.Details)
                        Console.WriteLine($"      - {SecretRedactor.Redact(detail)}");
                }
            }

            if (!result.Success)
            {
                Console.Error.WriteLine("Compilation failed:");
                foreach (var error in result.Errors)
                    Console.Error.WriteLine($"  - {SecretRedactor.Redact(error)}");
            }
        }

        if (result.ExitCode is 2)
            return ExitCodes.UsageOrConfiguration;
        if (result.ExitCode is 3 or 4 or 6 or 7)
            return result.ExitCode;
        return result.Success ? ExitCodes.Success : ExitCodes.SuiteOrWorkspace;
    }

    private static async Task<int> RunSyncKeyAsync(
        IServiceProvider services,
        ParsedCommand parsed,
        CancellationToken cancellationToken)
    {
        var orchestrator = services.GetRequiredService<JiraSuiteOrchestrator>();
        var result = await orchestrator.SyncKeyAsync(parsed.CompileJiraKey!, cancellationToken)
            .ConfigureAwait(false);

        if (parsed.Json)
        {
            Console.WriteLine(SecretRedactor.Redact(JsonSerializer.Serialize(result, JsonOutputOptions)));
        }
        else
        {
            Console.WriteLine("Desktop Automation Agent — sync-key");
            Console.WriteLine($"Jira key   : {result.JiraKey}");
            Console.WriteLine($"Success    : {result.Success}");
            Console.WriteLine($"From cache : {result.FromCache}");
            if (!string.IsNullOrWhiteSpace(result.RevisionToken))
                Console.WriteLine($"Revision   : {result.RevisionToken}");
            if (!string.IsNullOrWhiteSpace(result.CacheDirectory))
                Console.WriteLine($"Cache      : {result.CacheDirectory}");
            if (result.RateLimit is not null)
            {
                Console.WriteLine(
                    $"Rate limit : remaining={result.RateLimit.Remaining?.ToString() ?? "?"} " +
                    $"limit={result.RateLimit.Limit?.ToString() ?? "?"}");
            }

            if (!result.Success)
            {
                Console.Error.WriteLine("Sync failed:");
                foreach (var error in result.Errors)
                    Console.Error.WriteLine($"  - {SecretRedactor.Redact(error)}");
            }
        }

        return result.ExitCode;
    }

    private static async Task<int> RunRunKeyAsync(
        IServiceProvider services,
        ParsedCommand parsed,
        CancellationToken cancellationToken)
    {
        if (parsed.FromJira)
            AgentOptionsValidator.Validate(services.GetRequiredService<IOptions<AgentOptions>>().Value, OptionsValidationScope.Jira);
        else
            AgentOptionsValidator.Validate(services.GetRequiredService<IOptions<AgentOptions>>().Value, OptionsValidationScope.Workspace);

        var orchestrator = services.GetRequiredService<JiraSuiteOrchestrator>();
        var result = await orchestrator.RunKeyAsync(
                new JiraRunKeyRequest
                {
                    JiraKey = parsed.CompileJiraKey!,
                    DryRun = parsed.DryRun,
                    FromJira = parsed.FromJira,
                    RegistryPath = parsed.StepRegistryFile,
                    ObjectRepositoryPath = parsed.RepositoryFile,
                    KnowledgePath = parsed.KnowledgePath,
                    CatalogPath = parsed.CatalogFile
                },
                cancellationToken)
            .ConfigureAwait(false);

        if (parsed.Json)
        {
            var payload = new
            {
                success = result.Success,
                exitCode = result.ExitCode,
                jiraKey = result.JiraKey,
                compilationFingerprint = result.CompilationFingerprint,
                planHash = result.PlanHash,
                runId = result.RunId,
                result = result.Result,
                dryRun = result.DryRun,
                durationMilliseconds = result.DurationMilliseconds,
                errors = result.Errors
            };
            Console.WriteLine(SecretRedactor.Redact(JsonSerializer.Serialize(payload, JsonOutputOptions)));
        }
        else
        {
            Console.WriteLine(parsed.DryRun
                ? "Desktop Automation Agent — run-key (dry-run)"
                : "Desktop Automation Agent — run-key");
            Console.WriteLine($"Jira key    : {result.JiraKey}");
            Console.WriteLine($"Success     : {result.Success}");
            Console.WriteLine($"Result      : {result.Result}");
            if (!string.IsNullOrWhiteSpace(result.CompilationFingerprint))
                Console.WriteLine($"Fingerprint : {result.CompilationFingerprint}");
            if (!string.IsNullOrWhiteSpace(result.PlanHash))
                Console.WriteLine($"Plan SHA    : {result.PlanHash}");
            if (!string.IsNullOrWhiteSpace(result.RunId))
                Console.WriteLine($"Run ID      : {result.RunId}");
            if (!result.Success)
            {
                foreach (var error in result.Errors)
                    Console.Error.WriteLine($"  - {SecretRedactor.Redact(error)}");
            }
        }

        return result.ExitCode;
    }

    private static async Task<int> RunRunSuiteAsync(
        IServiceProvider services,
        ParsedCommand parsed,
        CancellationToken cancellationToken)
    {
        if (parsed.FromJira)
            AgentOptionsValidator.Validate(services.GetRequiredService<IOptions<AgentOptions>>().Value, OptionsValidationScope.Jira);
        else
            AgentOptionsValidator.Validate(services.GetRequiredService<IOptions<AgentOptions>>().Value, OptionsValidationScope.Suites);

        var orchestrator = services.GetRequiredService<JiraSuiteOrchestrator>();
        var result = await orchestrator.RunSuiteAsync(
                new JiraSuiteRunRequest
                {
                    SuiteFile = parsed.SuiteFile!,
                    DryRun = parsed.DryRun,
                    FromJira = parsed.FromJira,
                    ContinueOnFailure = parsed.ContinueOnFailure,
                    MaxFailures = parsed.MaxFailures ?? 0,
                    RegistryPath = parsed.StepRegistryFile,
                    ObjectRepositoryPath = parsed.RepositoryFile,
                    KnowledgePath = parsed.KnowledgePath,
                    CatalogPath = parsed.CatalogFile
                },
                cancellationToken)
            .ConfigureAwait(false);

        if (parsed.Json)
        {
            Console.WriteLine(SecretRedactor.Redact(JsonSerializer.Serialize(result, JsonOutputOptions)));
        }
        else
        {
            Console.WriteLine(parsed.DryRun
                ? "Desktop Automation Agent — run-suite (dry-run)"
                : "Desktop Automation Agent — run-suite");
            Console.WriteLine($"Suite run   : {result.SuiteRunId}");
            Console.WriteLine($"Suite       : {result.SuiteName}");
            Console.WriteLine($"Success     : {result.Success}");
            Console.WriteLine($"Fail-fast   : {result.FailFast}");
            Console.WriteLine($"Disabled    : {result.DisabledExcluded} excluded");
            if (!string.IsNullOrWhiteSpace(result.SuiteReportPath))
                Console.WriteLine($"Report      : {result.SuiteReportPath}");
            foreach (var key in result.Keys)
            {
                Console.WriteLine(
                    $"  - {key.JiraKey}: {key.Result} " +
                    $"(fp={key.CompilationFingerprint ?? "-"}, run={key.RunId ?? "-"}, ms={key.DurationMilliseconds:0})");
            }

            if (!result.Success)
            {
                foreach (var error in result.Errors)
                    Console.Error.WriteLine($"  - {SecretRedactor.Redact(error)}");
            }
        }

        return result.ExitCode;
    }

    private static int RunValidateCompiledKey(IServiceProvider services, ParsedCommand parsed)
    {
        AgentOptionsValidator.Validate(
            services.GetRequiredService<IOptions<AgentOptions>>().Value,
            OptionsValidationScope.Workspace);

        var service = services.GetRequiredService<BddCompilationService>();
        var result = service.ValidateCompiledKey(parsed.CompileJiraKey!, parsed.CompileFingerprint);

        if (parsed.Json)
        {
            var payload = new
            {
                success = result.IsValid,
                exitCode = result.IsValid ? ExitCodes.Success : ExitCodes.SuiteOrWorkspace,
                jiraKey = result.JiraKey,
                fingerprint = result.Fingerprint,
                planPath = result.PlanPath,
                planSha256 = result.PlanSha256,
                errors = result.Errors,
                warnings = result.Warnings
            };
            Console.WriteLine(SecretRedactor.Redact(JsonSerializer.Serialize(payload, JsonOutputOptions)));
        }
        else
        {
            Console.WriteLine("Desktop Automation Agent — validate-compiled-key");
            Console.WriteLine($"Jira key    : {result.JiraKey}");
            if (!string.IsNullOrWhiteSpace(result.Fingerprint))
                Console.WriteLine($"Fingerprint : {result.Fingerprint}");
            if (!string.IsNullOrWhiteSpace(result.PlanPath))
                Console.WriteLine($"Plan        : {result.PlanPath}");
            if (!result.IsValid)
            {
                Console.Error.WriteLine("Validation failed:");
                foreach (var error in result.Errors)
                    Console.Error.WriteLine($"  - {SecretRedactor.Redact(error)}");
            }
            else
            {
                Console.WriteLine("Compiled key validation succeeded.");
            }
        }

        return result.IsValid ? ExitCodes.Success : ExitCodes.SuiteOrWorkspace;
    }

    private static async Task<int> RunProposePlanAsync(
        IServiceProvider services,
        ParsedCommand parsed,
        CancellationToken cancellationToken)
    {
        AgentOptionsValidator.Validate(
            services.GetRequiredService<IOptions<AgentOptions>>().Value,
            OptionsValidationScope.Workspace);

        if (!PlanningRiskPolicy.TryParseMode(
                parsed.ApprovalMode
                ?? services.GetRequiredService<IOptions<AgentOptions>>().Value.Planning.DefaultApprovalMode,
                out var mode))
        {
            if (parsed.Json)
                WriteJsonError(ExitCodes.UsageOrConfiguration, $"Unknown approval mode '{parsed.ApprovalMode}'.");
            else
                await Console.Error.WriteLineAsync($"Unknown approval mode '{parsed.ApprovalMode}'.").ConfigureAwait(false);
            return ExitCodes.UsageOrConfiguration;
        }

        var planning = services.GetRequiredService<PlanningService>();
        var result = await planning.ProposeAsync(new PlanningRequest
        {
            Goal = parsed.Goal ?? string.Empty,
            UntrustedContext = parsed.UntrustedContext,
            ApplicationId = parsed.ApplicationId,
            ExpectedBuildFingerprint = parsed.ExpectedBuildFingerprint,
            KnowledgePath = parsed.KnowledgePath,
            ObjectRepositoryPath = parsed.RepositoryFile,
            StepRegistryPath = parsed.StepRegistryFile,
            CatalogPath = parsed.CatalogFile,
            RequestedApprovalMode = mode,
            ExecuteIfAuthorized = parsed.ExecuteIfAuthorized,
            DryRun = parsed.DryRun,
            PlanningId = parsed.PlanningId,
            CancellationToken = cancellationToken
        }).ConfigureAwait(false);

        // Explicit execute only when policy authorized auto-run-known-safe and a candidate exists.
        // Proposal-only / approval-required / denied never POST /ui from this command.
        var executed = false;
        string? runId = null;
        int? runExit = null;
        if (parsed.ExecuteIfAuthorized
            && result.EffectiveApprovalMode == PlanningApprovalMode.AutoRunKnownSafe
            && result.Success
            && !string.IsNullOrWhiteSpace(result.ArtifactDirectory))
        {
            var candidatePath = Path.Combine(result.ArtifactDirectory, "candidate.plan.json");
            if (File.Exists(candidatePath))
            {
                var runner = services.GetRequiredService<IDeterministicPlanRunner>();
                var report = await runner.RunAsync(candidatePath, dryRun: parsed.DryRun, cancellationToken)
                    .ConfigureAwait(false);
                executed = true;
                runId = report.RunId;
                runExit = report.ExitCode;
            }
        }
        else if (parsed.ExecuteIfAuthorized
                 && result.EffectiveApprovalMode != PlanningApprovalMode.AutoRunKnownSafe)
        {
            // Hard gate: never execute outside auto-run-known-safe.
        }

        if (parsed.Json)
        {
            var payload = new
            {
                success = result.Success && (runExit is null or 0),
                exitCode = runExit ?? result.ExitCode,
                planningId = result.PlanningId,
                artifactDirectory = result.ArtifactDirectory,
                effectiveApprovalMode = PlanningRiskPolicy.ToWire(result.EffectiveApprovalMode),
                candidatePlanHash = result.Validation.CandidatePlanHash,
                proposalHash = result.Validation.ProposalHash,
                retrievalManifestHash = result.Validation.RetrievalManifestHash,
                errors = result.Validation.Errors,
                warnings = result.Validation.Warnings,
                error = result.Error,
                executed,
                runId,
                dryRun = parsed.DryRun
            };
            Console.WriteLine(SecretRedactor.Redact(JsonSerializer.Serialize(payload, JsonOutputOptions)));
        }
        else
        {
            Console.WriteLine("Desktop Automation Agent — propose-plan");
            Console.WriteLine($"Planning id      : {result.PlanningId}");
            Console.WriteLine($"Approval mode    : {PlanningRiskPolicy.ToWire(result.EffectiveApprovalMode)}");
            Console.WriteLine($"Artifacts        : {result.ArtifactDirectory}");
            Console.WriteLine($"Success          : {result.Success}");
            if (!string.IsNullOrWhiteSpace(result.Error))
                Console.WriteLine($"Error            : {SecretRedactor.Redact(result.Error)}");
            foreach (var err in result.Validation.Errors)
                Console.WriteLine($"  - {SecretRedactor.Redact(err)}");
            if (executed)
                Console.WriteLine($"Executed run     : {runId} (dryRun={parsed.DryRun})");
        }

        return runExit ?? result.ExitCode;
    }

    private static async Task<int> RunDoctorAsync(
        IServiceProvider services,
        bool json,
        CancellationToken cancellationToken)
    {
        var readiness = services.GetRequiredService<IAgentReadinessService>();
        var report = await readiness.RunDoctorAsync(cancellationToken).ConfigureAwait(false);

        if (json)
        {
            var payload = JsonSerializer.Serialize(report, JsonOutputOptions);
            Console.WriteLine(SecretRedactor.Redact(payload));
        }
        else
        {
            WriteHumanDoctor(report);
        }

        return report.ExitCode;
    }

    private static void WriteHumanDoctor(ReadinessReport report)
    {
        Console.WriteLine("Desktop Automation Agent — doctor");
        Console.WriteLine($"Success          : {report.Success}");
        Console.WriteLine($"Exit code        : {report.ExitCode}");
        Console.WriteLine($"Workspace        : {report.WorkspaceRoot}");
        if (!string.IsNullOrWhiteSpace(report.Username))
            Console.WriteLine($"Username         : {report.Username}");
        if (!string.IsNullOrWhiteSpace(report.DriverBaseUrl))
            Console.WriteLine($"Driver base URL  : {report.DriverBaseUrl}");
        if (!string.IsNullOrWhiteSpace(report.DiscoveryMethod))
            Console.WriteLine($"Discovery        : {report.DiscoveryMethod}");
        if (!string.IsNullOrWhiteSpace(report.DriverVersion))
            Console.WriteLine($"Driver version   : {report.DriverVersion}");
        if (report.CatalogSchemaVersion.HasValue)
            Console.WriteLine($"Catalog schema   : {report.CatalogSchemaVersion}");
        if (report.OperationCount.HasValue)
            Console.WriteLine($"Operation count  : {report.OperationCount}");

        Console.WriteLine("Checks:");
        foreach (var check in report.Checks)
            Console.WriteLine($"  [{check.Status}] {check.Name}: {check.Detail}");

        if (report.Errors.Count > 0)
        {
            Console.WriteLine("Errors:");
            foreach (var error in report.Errors)
                Console.WriteLine($"  - {error}");
        }
    }
}
