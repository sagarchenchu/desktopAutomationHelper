using DesktopAutomationAgent.Cli;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;

namespace DesktopAutomationAgent.Mcp;

/// <summary>
/// Local stdio MCP sidecar. stdout is reserved for JSON-RPC; all logs go to stderr.
/// No unauthenticated network listener is started.
/// </summary>
public static class McpServerHost
{
    public const string ServerName = "desktop-automation-agent";
    public const string ServerVersion = "0.5.0";

    public static async Task<int> RunAsync(
        string[] configurationArgs,
        CancellationToken cancellationToken = default)
    {
        var builder = Host.CreateApplicationBuilder(new HostApplicationBuilderSettings
        {
            Args = configurationArgs,
            ContentRootPath = Directory.GetCurrentDirectory()
        });

        AgentServiceRegistrar.ConfigureAgentConfiguration(builder, configurationArgs);
        AgentServiceRegistrar.AddAgentCoreServices(builder, quietAgentLogs: true);

        builder.Services
            .AddMcpServer(options =>
            {
                options.ServerInfo = new()
                {
                    Name = ServerName,
                    Version = ServerVersion
                };
            })
            .WithStdioServerTransport()
            .WithTools<DesktopMcpTools>()
            .WithResources<DesktopMcpResources>();

        // Explicit: never bind an HTTP MCP listener in this phase.
        using var host = builder.Build();
        var logger = host.Services.GetRequiredService<ILoggerFactory>().CreateLogger("McpServerHost");
        logger.LogInformation(
            "Starting local stdio MCP server {Name} {Version}. Same-user local sidecar; no network listener.",
            ServerName,
            ServerVersion);

        await host.RunAsync(cancellationToken).ConfigureAwait(false);
        return ExitCodes.Success;
    }

    /// <summary>Test helper: build an in-process host with a custom stream transport factory hook.</summary>
    internal static IHost BuildTestHost(
        string[] configurationArgs,
        Action<IServiceCollection>? configure = null)
    {
        var builder = Host.CreateApplicationBuilder(new HostApplicationBuilderSettings
        {
            Args = configurationArgs,
            ContentRootPath = Directory.GetCurrentDirectory()
        });

        AgentServiceRegistrar.ConfigureAgentConfiguration(builder, configurationArgs);
        AgentServiceRegistrar.AddAgentCoreServices(builder, quietAgentLogs: true);
        configure?.Invoke(builder.Services);

        builder.Services
            .AddMcpServer(options =>
            {
                options.ServerInfo = new()
                {
                    Name = ServerName,
                    Version = ServerVersion
                };
            })
            .WithTools<DesktopMcpTools>()
            .WithResources<DesktopMcpResources>();

        return builder.Build();
    }
}
