namespace DesktopAutomationAgent.Execution;

public interface IDeterministicPlanRunner
{
    Task<RunReport> RunAsync(
        string planPath,
        bool dryRun,
        CancellationToken cancellationToken = default);
}
