using System.Text.Json;
using System.Text.Json.Serialization;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Execution;
using DesktopAutomationAgent.Plans;
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
        var jsonStdoutMode = parsed.Json
            && parsed.Kind is AgentCommandKind.Doctor or AgentCommandKind.ValidatePlan or AgentCommandKind.RunPlan;

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

        builder.Configuration.Sources.Clear();
        builder.Configuration
            .AddJsonFile(
                Path.Combine(AppContext.BaseDirectory, "appsettings.json"),
                optional: true,
                reloadOnChange: false)
            .AddJsonFile(
                Path.Combine(Directory.GetCurrentDirectory(), "automation", "config", "agentsettings.local.json"),
                optional: true,
                reloadOnChange: false)
            .AddEnvironmentVariables(prefix: "DA_AGENT__")
            .AddCommandLine(configurationArgs);

        builder.Logging.ClearProviders();
        // Keep machine-readable stdout clean: all log records go to stderr.
        builder.Logging.AddSimpleConsole(options =>
        {
            options.TimestampFormat = "yyyy-MM-ddTHH:mm:ss.fffZ ";
            options.SingleLine = true;
            options.UseUtcTimestamp = true;
        });
        builder.Services.Configure<Microsoft.Extensions.Logging.Console.ConsoleLoggerOptions>(options =>
        {
            options.LogToStandardErrorThreshold = LogLevel.Trace;
        });
        builder.Logging.AddFilter("Microsoft.Hosting.Lifetime", LogLevel.Warning);

        if (jsonStdoutMode)
        {
            builder.Logging.AddFilter("DesktopAutomationAgent", LogLevel.Warning);
        }

        builder.Services.Configure<AgentOptions>(builder.Configuration);

        builder.Services.AddHttpClient("driver-unauthenticated");
        builder.Services.AddHttpClient("driver-authenticated");

        builder.Services.AddSingleton<IWorkspaceManager, WorkspaceManager>();
        builder.Services.AddSingleton<ISuiteManifestReader, SuiteManifestReader>();
        builder.Services.AddSingleton<PlanManifestReader>();
        builder.Services.AddSingleton<IDriverConnectionResolver, DriverConnectionResolver>();
        builder.Services.AddSingleton<IDriverCatalogClient, DriverCatalogClient>();
        builder.Services.AddSingleton<IDriverUiClient, DriverUiClient>();
        builder.Services.AddSingleton<AssertionEvaluator>();
        builder.Services.AddSingleton<RunArtifactWriter>();
        builder.Services.AddSingleton<IDeterministicPlanRunner, DeterministicPlanRunner>();
        builder.Services.AddSingleton<IAgentReadinessService, AgentReadinessService>();

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
        var result = reader.Read(file);

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
