using System.Text.RegularExpressions;

namespace DesktopAutomationAgent.Configuration;

public enum OptionsValidationScope
{
    Workspace,
    Suites,
    Driver,
    Runner,
    ObjectRepository
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

        if (scope is OptionsValidationScope.Runner or OptionsValidationScope.Driver or OptionsValidationScope.ObjectRepository)
        {
            ValidateRunnerOptions(options.Runner);
        }

        if (scope is OptionsValidationScope.ObjectRepository)
        {
            ValidateObjectRepositoryOptions(options.ObjectRepository);
        }
    }

    public static void ValidateObjectRepositoryOptions(ObjectRepositoryOptions repository)
    {
        ArgumentNullException.ThrowIfNull(repository);

        const int maxFiftyMegabytes = 52_428_800;

        if (repository.MaxFileBytes <= 0 || repository.MaxFileBytes > maxFiftyMegabytes)
        {
            throw new AgentConfigurationException(
                $"ObjectRepository:MaxFileBytes must be between 1 and {maxFiftyMegabytes}.");
        }

        if (repository.MaxPages <= 0 || repository.MaxPages > 10_000)
        {
            throw new AgentConfigurationException("ObjectRepository:MaxPages must be between 1 and 10000.");
        }

        if (repository.MaxElementsPerPage <= 0 || repository.MaxElementsPerPage > 100_000)
        {
            throw new AgentConfigurationException(
                "ObjectRepository:MaxElementsPerPage must be between 1 and 100000.");
        }

        if (repository.MaxTotalElements <= 0 || repository.MaxTotalElements > 1_000_000)
        {
            throw new AgentConfigurationException(
                "ObjectRepository:MaxTotalElements must be between 1 and 1000000.");
        }

        if (repository.DiagnosticTimeoutMilliseconds < 500 || repository.DiagnosticTimeoutMilliseconds > 15_000)
        {
            throw new AgentConfigurationException(
                "ObjectRepository:DiagnosticTimeoutMilliseconds must be between 500 and 15000.");
        }
    }

    public static void ValidateRunnerOptions(RunnerOptions runner)
    {
        ArgumentNullException.ThrowIfNull(runner);

        const int maxFiftyMegabytes = 52_428_800;
        const int maxTimeoutSeconds = 3_600;

        if (runner.StepTransportTimeoutSeconds <= 0 || runner.StepTransportTimeoutSeconds > maxTimeoutSeconds)
        {
            throw new AgentConfigurationException(
                $"Runner:StepTransportTimeoutSeconds must be between 1 and {maxTimeoutSeconds}.");
        }

        if (runner.CleanupTimeoutSeconds <= 0 || runner.CleanupTimeoutSeconds > maxTimeoutSeconds)
        {
            throw new AgentConfigurationException(
                $"Runner:CleanupTimeoutSeconds must be between 1 and {maxTimeoutSeconds}.");
        }

        if (runner.MaxPlanBytes <= 0 || runner.MaxPlanBytes > maxFiftyMegabytes)
        {
            throw new AgentConfigurationException(
                $"Runner:MaxPlanBytes must be between 1 and {maxFiftyMegabytes}.");
        }

        if (runner.MaxResponseBytes <= 0 || runner.MaxResponseBytes > maxFiftyMegabytes)
        {
            throw new AgentConfigurationException(
                $"Runner:MaxResponseBytes must be between 1 and {maxFiftyMegabytes}.");
        }

        if (runner.RegexTimeoutMilliseconds <= 0 || runner.RegexTimeoutMilliseconds > 60_000)
        {
            throw new AgentConfigurationException(
                "Runner:RegexTimeoutMilliseconds must be between 1 and 60000.");
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
        {
            // Verify discovery always targets a loopback probe and then builds
            // http://127.0.0.1:{port}/. Remote verify hosts are rejected even when
            // AllowRemoteDriver is true; use explicit BaseUrl+BearerToken instead.
            DriverUrlRules.EnsureAllowed(driver.VerifyUrl, allowRemote: false);
        }
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
