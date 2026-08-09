namespace DesktopAutomationAgent.Driver;

public sealed class DriverConnectionException : Exception
{
    public DriverConnectionException(string message) : base(SecretRedactor.Redact(message))
    {
    }

    public DriverConnectionException(string message, Exception inner)
        : base(SecretRedactor.Redact(message), inner)
    {
    }
}

public sealed class DriverCatalogException : Exception
{
    public DriverCatalogException(string message) : base(SecretRedactor.Redact(message))
    {
    }

    public DriverCatalogException(string message, Exception inner)
        : base(SecretRedactor.Redact(message), inner)
    {
    }
}
