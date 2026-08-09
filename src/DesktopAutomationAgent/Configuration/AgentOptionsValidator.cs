using System.Text.RegularExpressions;

namespace DesktopAutomationAgent.Configuration;

public enum OptionsValidationScope
{
    Workspace,
    Suites,
    Driver
}

public sealed class AgentConfigurationException : Exception
{
    public AgentConfigurationException(string message) : base(message)
    {
    }
}

public static class AgentOptionsValidator
{
    public static void Validate(AgentOptions options, OptionsValidationScope scope)
    {
        ArgumentNullException.ThrowIfNull(options);

        if (scope is OptionsValidationScope.Workspace or OptionsValidationScope.Suites or OptionsValidationScope.Driver)
        {
            if (string.IsNullOrWhiteSpace(options.Workspace.Root))
                throw new AgentConfigurationException("Workspace:Root is required.");
        }

        if (scope is OptionsValidationScope.Suites)
        {
            if (string.IsNullOrWhiteSpace(options.Suites.JiraKeyPattern))
                throw new AgentConfigurationException("Suites:JiraKeyPattern is required.");

            try
            {
                _ = new Regex(options.Suites.JiraKeyPattern, RegexOptions.CultureInvariant | RegexOptions.Compiled);
            }
            catch (ArgumentException ex)
            {
                throw new AgentConfigurationException($"Suites:JiraKeyPattern is not a valid regular expression: {ex.Message}");
            }
        }

        if (scope is OptionsValidationScope.Driver)
        {
            ValidateDriverOptions(options.Driver);
        }
    }

    public static void ValidateDriverOptions(DriverOptions driver)
    {
        ArgumentNullException.ThrowIfNull(driver);

        var hasUrl = !string.IsNullOrWhiteSpace(driver.BaseUrl);
        var hasToken = !string.IsNullOrWhiteSpace(driver.BearerToken);

        if (hasUrl ^ hasToken)
        {
            throw new AgentConfigurationException(
                "Driver:BaseUrl and Driver:BearerToken must be supplied together. " +
                "Provide both for an explicit connection, or omit both to use verify-endpoint discovery.");
        }

        if (driver.RequestTimeoutSeconds <= 0)
            throw new AgentConfigurationException("Driver:RequestTimeoutSeconds must be greater than zero.");

        if (driver.ExpectedCatalogSchemaVersion <= 0)
            throw new AgentConfigurationException("Driver:ExpectedCatalogSchemaVersion must be greater than zero.");

        if (!hasUrl && string.IsNullOrWhiteSpace(driver.VerifyUrl))
            throw new AgentConfigurationException("Driver:VerifyUrl is required when explicit BaseUrl/BearerToken are not set.");

        if (hasUrl)
            DriverUrlRules.EnsureAllowed(driver.BaseUrl!, driver.AllowRemoteDriver);

        if (!hasUrl)
            DriverUrlRules.EnsureAllowed(driver.VerifyUrl, driver.AllowRemoteDriver);
    }
}

public static class DriverUrlRules
{
    public static bool IsLoopbackHost(string host) =>
        string.Equals(host, "localhost", StringComparison.OrdinalIgnoreCase)
        || string.Equals(host, "127.0.0.1", StringComparison.OrdinalIgnoreCase)
        || string.Equals(host, "::1", StringComparison.OrdinalIgnoreCase)
        || string.Equals(host, "[::1]", StringComparison.OrdinalIgnoreCase);

    public static Uri EnsureAllowed(string url, bool allowRemote)
    {
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri)
            || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
        {
            throw new AgentConfigurationException($"URL '{url}' is not a valid absolute http(s) URL.");
        }

        if (!allowRemote && !IsLoopbackHost(uri.Host))
        {
            throw new AgentConfigurationException(
                $"Remote driver URL host '{uri.Host}' is not allowed. " +
                "Use a loopback address or set Driver:AllowRemoteDriver=true.");
        }

        return uri;
    }
}
