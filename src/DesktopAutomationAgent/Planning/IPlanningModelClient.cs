namespace DesktopAutomationAgent.Planning;

/// <summary>
/// Provider-neutral planning model client. Implementations must never call the driver
/// or mutate workspace state. Credentials come from env/secret configuration only.
/// </summary>
public interface IPlanningModelClient
{
    string ProviderName { get; }

    Task<PlanningModelResponse> ProposeAsync(PlanningModelRequest request, CancellationToken cancellationToken = default);
}
