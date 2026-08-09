using DesktopAutomationAgent.Cli;
using DesktopAutomationAgent.Driver;

try
{
    return await AgentCli.RunAsync(args).ConfigureAwait(false);
}
catch (Exception ex)
{
    await Console.Error.WriteLineAsync($"Unhandled error: {SecretRedactor.Redact(ex.Message)}").ConfigureAwait(false);
    return ExitCodes.UsageOrConfiguration;
}
