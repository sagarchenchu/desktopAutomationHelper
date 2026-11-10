namespace DesktopAutomationAgent.Observation;

public interface ISemanticObservationService
{
    Task<ObservationResult> ObserveAsync(
        ObservationOptions options,
        CancellationToken cancellationToken = default);
}
