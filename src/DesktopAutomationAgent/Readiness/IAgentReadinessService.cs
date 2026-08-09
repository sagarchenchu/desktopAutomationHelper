namespace DesktopAutomationAgent.Readiness;

public interface IAgentReadinessService
{
    Task<ReadinessReport> RunDoctorAsync(CancellationToken cancellationToken = default);
}
