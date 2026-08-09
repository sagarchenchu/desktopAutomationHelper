using DesktopAutomationAgent.Plans;

namespace DesktopAutomationAgent.Driver;

public interface IDriverUiClient
{
    Task<UiExecutionResponse> ExecuteStepAsync(
        DriverConnection connection,
        PlanStep step,
        CancellationToken cancellationToken = default);
}
