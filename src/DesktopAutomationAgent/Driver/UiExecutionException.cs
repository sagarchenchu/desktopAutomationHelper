namespace DesktopAutomationAgent.Driver;

public sealed class UiExecutionException : Exception
{
    public UiExecutionException(
        UiFailureClassification classification,
        string message,
        UiExecutionResponse? response = null,
        Exception? innerException = null)
        : base(SecretRedactor.Redact(message), innerException)
    {
        Classification = classification;
        Response = response;
    }

    public UiFailureClassification Classification { get; }

    public UiExecutionResponse? Response { get; }
}
