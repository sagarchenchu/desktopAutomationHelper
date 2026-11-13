namespace DesktopAutomationAgent.Configuration;

public enum OptionsValidationScope
{
    Workspace,
    Suites,
    Driver,
    Runner,
    ObjectRepository,
    Knowledge,
    Jira,
    Recovery
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
            // Empty / whitespace means "no project-specific additional restriction"
            // beyond the canonical Jira contract.
            if (!string.IsNullOrWhiteSpace(options.Suites.JiraKeyPattern))
            {
                try
                {
                    _ = JiraKeyContract.CompileProjectPattern(options.Suites.JiraKeyPattern);
                }
                catch (ArgumentException ex)
                {
                    throw new AgentConfigurationException(
                        $"Suites:JiraKeyPattern is not a valid regular expression: {ex.Message}");
                }
            }
        }

        if (scope is OptionsValidationScope.Driver)
        {
            ValidateDriverOptions(options.Driver);
        }

        if (scope is OptionsValidationScope.Runner or OptionsValidationScope.Driver or OptionsValidationScope.ObjectRepository)
        {
            ValidateRunnerOptions(options.Runner);
            ValidateRecoveryOptions(options.Recovery);
        }

        if (scope is OptionsValidationScope.Recovery)
        {
            ValidateRecoveryOptions(options.Recovery);
        }

        if (scope is OptionsValidationScope.ObjectRepository)
        {
            ValidateObjectRepositoryOptions(options.ObjectRepository);
        }

        if (scope is OptionsValidationScope.Knowledge)
        {
            ValidateKnowledgeOptions(options.Knowledge);
        }

        if (scope is OptionsValidationScope.Jira)
        {
            ValidateJiraOptions(options.Jira);
        }
    }

    public static void ValidateJiraOptions(JiraOptions jira)
    {
        ArgumentNullException.ThrowIfNull(jira);

        if (string.IsNullOrWhiteSpace(jira.BaseUrl))
            throw new AgentConfigurationException("Jira:BaseUrl is required for Jira sync commands.");

        _ = JiraUrlRules.EnsureAllowed(jira.BaseUrl!, jira.AllowLocalHttpTestServer);

        if (jira.RequestTimeoutSeconds <= 0 || jira.RequestTimeoutSeconds > 300)
            throw new AgentConfigurationException("Jira:RequestTimeoutSeconds must be between 1 and 300.");

        if (jira.MaxResponseBytes <= 0 || jira.MaxResponseBytes > 16_777_216)
            throw new AgentConfigurationException("Jira:MaxResponseBytes must be between 1 and 16777216.");

        if (jira.MaxRetries < 0 || jira.MaxRetries > 3)
            throw new AgentConfigurationException("Jira:MaxRetries must be between 0 and 3.");

        if (jira.PageSize <= 0 || jira.PageSize > 100)
            throw new AgentConfigurationException("Jira:PageSize must be between 1 and 100.");

        if (jira.MaxPages <= 0 || jira.MaxPages > 100)
            throw new AgentConfigurationException("Jira:MaxPages must be between 1 and 100.");

        var searchPagination = (jira.SearchPagination ?? string.Empty).Trim();
        if (!string.Equals(searchPagination, "jql-token", StringComparison.OrdinalIgnoreCase))
        {
            throw new AgentConfigurationException(
                "Jira:SearchPagination must be 'jql-token' (Cloud /rest/api/3/search/jql). " +
                "Legacy DC offset pagination is not enabled.");
        }

        jira.SearchPagination = "jql-token";

        var mode = (jira.AuthMode ?? "basic").Trim().ToLowerInvariant();
        if (mode is not ("basic" or "bearer" or "pat"))
            throw new AgentConfigurationException("Jira:AuthMode must be one of: basic, bearer, pat.");

        var profile = (jira.MappingProfile ?? "plain-fields").Trim().ToLowerInvariant();
        if (profile is not ("plain-fields" or "xray" or "zephyr" or "normalized" or "custom"))
        {
            throw new AgentConfigurationException(
                "Jira:MappingProfile must be one of: plain-fields, xray, zephyr, normalized, custom.");
        }
    }

    public static void ValidateKnowledgeOptions(KnowledgeOptions knowledge)
    {
        ArgumentNullException.ThrowIfNull(knowledge);

        const int maxFiftyMegabytes = 52_428_800;
        if (knowledge.MaxFileBytes <= 0 || knowledge.MaxFileBytes > maxFiftyMegabytes)
        {
            throw new AgentConfigurationException(
                $"Knowledge:MaxFileBytes must be between 1 and {maxFiftyMegabytes}.");
        }

        if (knowledge.MaxExplorationRuns <= 0 || knowledge.MaxExplorationRuns > 500)
        {
            throw new AgentConfigurationException("Knowledge:MaxExplorationRuns must be between 1 and 500.");
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

    public static void ValidateRecoveryOptions(RecoveryOptions recovery)
    {
        ArgumentNullException.ThrowIfNull(recovery);

        if (recovery.MaxTotalRecoveriesPerRun < 0 || recovery.MaxTotalRecoveriesPerRun > 50)
        {
            throw new AgentConfigurationException(
                "Recovery:MaxTotalRecoveriesPerRun must be between 0 and 50.");
        }

        if (recovery.MaxRecoveryAttemptsPerStep < 0 || recovery.MaxRecoveryAttemptsPerStep > 10)
        {
            throw new AgentConfigurationException(
                "Recovery:MaxRecoveryAttemptsPerStep must be between 0 and 10.");
        }

        if (recovery.MaxAdditionalDurationSeconds < 0 || recovery.MaxAdditionalDurationSeconds > 3_600)
        {
            throw new AgentConfigurationException(
                "Recovery:MaxAdditionalDurationSeconds must be between 0 and 3600.");
        }

        if (recovery.BoundedWaitMilliseconds < 0 || recovery.BoundedWaitMilliseconds > 60_000)
        {
            throw new AgentConfigurationException(
                "Recovery:BoundedWaitMilliseconds must be between 0 and 60000.");
        }

        if (string.IsNullOrWhiteSpace(recovery.Provider))
        {
            throw new AgentConfigurationException("Recovery:Provider is required (use 'fake' for CI).");
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

/// <summary>
/// Jira remote URL policy: HTTPS required for remote hosts; non-HTTPS allowed only
/// for explicit local test servers when <see cref="JiraOptions.AllowLocalHttpTestServer"/> is true.
/// </summary>
public static class JiraUrlRules
{
    public static Uri EnsureAllowed(string url, bool allowLocalHttpTestServer)
    {
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri)
            || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
        {
            throw new AgentConfigurationException($"Jira:BaseUrl '{url}' is not a valid absolute http(s) URL.");
        }

        if (!string.IsNullOrEmpty(uri.UserInfo))
        {
            throw new AgentConfigurationException(
                "Jira:BaseUrl must not embed credentials (userinfo). " +
                "Use Jira auth options / env / OS secret store instead.");
        }

        var loopback = DriverUrlRules.IsLoopbackHost(uri.Host);
        if (uri.Scheme == Uri.UriSchemeHttp)
        {
            if (!(allowLocalHttpTestServer && loopback))
            {
                throw new AgentConfigurationException(
                    "Jira:BaseUrl must use HTTPS for remote hosts. " +
                    "Non-HTTPS is allowed only when Jira:AllowLocalHttpTestServer=true and the host is loopback.");
            }
        }

        return uri;
    }
}
