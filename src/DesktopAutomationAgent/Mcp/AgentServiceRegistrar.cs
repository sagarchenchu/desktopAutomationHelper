using DesktopAutomationAgent.Compilation;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Execution;
using DesktopAutomationAgent.Exploration;
using DesktopAutomationAgent.Jira;
using DesktopAutomationAgent.Knowledge;
using DesktopAutomationAgent.ObjectRepository;
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

namespace DesktopAutomationAgent.Mcp;

/// <summary>Shared DI registration for CLI and MCP hosts.</summary>
public static class AgentServiceRegistrar
{
    public static void AddAgentCoreServices(IHostApplicationBuilder builder, bool quietAgentLogs = false)
    {
        builder.Logging.ClearProviders();
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
        if (quietAgentLogs)
            builder.Logging.AddFilter("DesktopAutomationAgent", LogLevel.Warning);

        builder.Services.Configure<AgentOptions>(builder.Configuration);

        builder.Services.AddHttpClient("driver-unauthenticated");
        builder.Services.AddHttpClient("driver-authenticated");
        builder.Services.AddHttpClient(JiraHttpClientNames.Default);

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
        builder.Services.AddSingleton<ObjectRepositoryReader>();
        builder.Services.AddSingleton<ObjectReferenceResolver>();
        builder.Services.AddSingleton<PlanObjectReferenceExpander>();
        builder.Services.AddSingleton<PlanObjectRepositoryIntegrator>();
        builder.Services.AddSingleton<ObjectArtifactWriter>();
        builder.Services.AddSingleton<ObjectCandidateGenerator>();
        builder.Services.AddSingleton<ObjectCaptureService>();
        builder.Services.AddSingleton<ObjectVerificationService>();
        builder.Services.AddSingleton<SemanticSnapshotNormalizer>();
        builder.Services.AddSingleton<ISemanticObservationService, SemanticObservationService>();
        builder.Services.AddSingleton<ExplorationPolicyReader>();
        builder.Services.AddSingleton<StateFingerprinter>();
        builder.Services.AddSingleton<CapabilityResolver>();
        builder.Services.AddSingleton<IProcessIdentityProbe, SystemProcessIdentityProbe>();
        builder.Services.AddSingleton<UiStabilizationService>();
        builder.Services.AddSingleton<WindowTransitionDetector>();
        builder.Services.AddSingleton<ExplorationArtifactWriter>();
        builder.Services.AddSingleton<IExplorationCoordinator, ExplorationCoordinator>();
        builder.Services.AddSingleton<KnowledgeService>();
        builder.Services.AddSingleton<BddCompilationService>();
        builder.Services.AddSingleton<IJiraTestCaseSource, JiraRestTestCaseSource>();
        builder.Services.AddSingleton<JiraSuiteOrchestrator>();
        builder.Services.AddSingleton<McpOpaqueRefStore>();
        builder.Services.AddSingleton<IDesktopMcpFacade, DesktopMcpFacade>();
        builder.Services.AddSingleton<IPlanningModelClient>(sp =>
        {
            var planning = sp.GetRequiredService<IOptions<AgentOptions>>().Value.Planning;
            // CI and default: fake deterministic model. Real providers are adapter-only and
            // must not be referenced from DesktopAutomationDriver.
            if (string.IsNullOrWhiteSpace(planning.Provider)
                || string.Equals(planning.Provider, "fake", StringComparison.OrdinalIgnoreCase))
            {
                return new FakePlanningModelClient();
            }

            throw new AgentConfigurationException(
                $"Planning:Provider '{planning.Provider}' is not registered. Use 'fake' for CI or register a provider-neutral adapter.");
        });
        builder.Services.AddSingleton<PlanningRetrievalService>();
        builder.Services.AddSingleton<PlanningArtifactWriter>();
        builder.Services.AddSingleton<PlanningService>();
    }

    public static void ConfigureAgentConfiguration(HostApplicationBuilder builder, string[] configurationArgs)
    {
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
    }
}
