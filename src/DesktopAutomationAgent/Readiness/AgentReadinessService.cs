using DesktopAutomationAgent.Cli;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Workspace;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Readiness;

public sealed class AgentReadinessService : IAgentReadinessService
{
    private readonly AgentOptions _options;
    private readonly IWorkspaceManager _workspace;
    private readonly IDriverConnectionResolver _connectionResolver;
    private readonly IDriverCatalogClient _catalogClient;
    private readonly ILogger<AgentReadinessService> _logger;

    public AgentReadinessService(
        IOptions<AgentOptions> options,
        IWorkspaceManager workspace,
        IDriverConnectionResolver connectionResolver,
        IDriverCatalogClient catalogClient,
        ILogger<AgentReadinessService> logger)
    {
        _options = options.Value;
        _workspace = workspace;
        _connectionResolver = connectionResolver;
        _catalogClient = catalogClient;
        _logger = logger;
    }

    public async Task<ReadinessReport> RunDoctorAsync(CancellationToken cancellationToken = default)
    {
        var checks = new List<ReadinessCheck>();
        var errors = new List<string>();

        string? username = null;
        string? driverBaseUrl = null;
        string? discoveryMethod = null;
        string? driverVersion = null;
        int? schemaVersion = null;
        int? operationCount = null;

        try
        {
            AgentOptionsValidator.Validate(_options, OptionsValidationScope.Driver);
            checks.Add(Pass("configuration", "Driver and workspace options are valid."));
        }
        catch (AgentConfigurationException ex)
        {
            checks.Add(Fail("configuration", ex.Message));
            errors.Add(ex.Message);
            return Build(false, ExitCodes.UsageOrConfiguration, checks, errors, username, driverBaseUrl, discoveryMethod, driverVersion, schemaVersion, operationCount);
        }

        try
        {
            _workspace.EnsureInitialized();
            checks.Add(Pass("workspace", $"Workspace ready at {_workspace.RootPath}."));
        }
        catch (WorkspaceException ex)
        {
            checks.Add(Fail("workspace", ex.Message));
            errors.Add(ex.Message);
            return Build(false, ExitCodes.SuiteOrWorkspace, checks, errors, username, driverBaseUrl, discoveryMethod, driverVersion, schemaVersion, operationCount);
        }

        DriverConnection connection;
        try
        {
            connection = await _connectionResolver.ResolveAsync(cancellationToken).ConfigureAwait(false);
            username = connection.DiscoveredUsername ?? Environment.UserName;
            driverBaseUrl = connection.SafeBaseUrl;
            discoveryMethod = connection.DiscoveryMethod;
            checks.Add(Pass(
                "driver-discovery",
                $"Connected via {connection.DiscoveryMethod} to {connection.SafeBaseUrl}."));
            checks.Add(Pass(
                "username-safety",
                $"Current user '{Environment.UserName}' is accepted for this connection."));
        }
        catch (AgentConfigurationException ex)
        {
            checks.Add(Fail("driver-discovery", ex.Message));
            errors.Add(ex.Message);
            return Build(false, ExitCodes.UsageOrConfiguration, checks, errors, username, driverBaseUrl, discoveryMethod, driverVersion, schemaVersion, operationCount);
        }
        catch (DriverConnectionException ex)
        {
            var detail = SecretRedactor.Redact(ex.Message);
            checks.Add(Fail("driver-discovery", detail));
            errors.Add(detail);
            return Build(false, ExitCodes.DriverUnavailable, checks, errors, username, driverBaseUrl, discoveryMethod, driverVersion, schemaVersion, operationCount);
        }

        try
        {
            var status = await _catalogClient.GetStatusAsync(connection, cancellationToken).ConfigureAwait(false);
            if (!status.Ready)
            {
                var detail = "Driver /status reported ready=false.";
                checks.Add(Fail("driver-status", detail));
                errors.Add(detail);
                return Build(false, ExitCodes.DriverUnavailable, checks, errors, username, driverBaseUrl, discoveryMethod, driverVersion, schemaVersion, operationCount);
            }

            driverVersion = status.Build?.Version;
            checks.Add(Pass("driver-status", status.Message));
            checks.Add(Pass("authentication", "Bearer authentication succeeded for GET /status."));
        }
        catch (DriverCatalogException ex) when (ex.Message.Contains("401", StringComparison.Ordinal))
        {
            var detail = SecretRedactor.Redact(ex.Message);
            checks.Add(Fail("authentication", detail));
            errors.Add(detail);
            return Build(false, ExitCodes.AuthOrCatalog, checks, errors, username, driverBaseUrl, discoveryMethod, driverVersion, schemaVersion, operationCount);
        }
        catch (DriverCatalogException ex)
        {
            var detail = SecretRedactor.Redact(ex.Message);
            checks.Add(Fail("driver-status", detail));
            errors.Add(detail);
            return Build(false, ExitCodes.AuthOrCatalog, checks, errors, username, driverBaseUrl, discoveryMethod, driverVersion, schemaVersion, operationCount);
        }
        catch (DriverConnectionException ex)
        {
            var detail = SecretRedactor.Redact(ex.Message);
            checks.Add(Fail("driver-status", detail));
            errors.Add(detail);
            return Build(false, ExitCodes.DriverUnavailable, checks, errors, username, driverBaseUrl, discoveryMethod, driverVersion, schemaVersion, operationCount);
        }

        try
        {
            var catalog = await _catalogClient.GetOperationsAsync(connection, cancellationToken).ConfigureAwait(false);
            CatalogCompatibility.Validate(catalog, _options.Driver.ExpectedCatalogSchemaVersion);
            schemaVersion = catalog.SchemaVersion;
            operationCount = catalog.Operations.Count;
            driverVersion ??= catalog.DriverVersion;

            checks.Add(Pass(
                "operation-catalog",
                $"schemaVersion={catalog.SchemaVersion}, driverVersion={catalog.DriverVersion}, operations={catalog.Operations.Count}."));
        }
        catch (DriverCatalogException ex) when (ex.Message.Contains("401", StringComparison.Ordinal))
        {
            var detail = SecretRedactor.Redact(ex.Message);
            checks.Add(Fail("authentication", detail));
            errors.Add(detail);
            return Build(false, ExitCodes.AuthOrCatalog, checks, errors, username, driverBaseUrl, discoveryMethod, driverVersion, schemaVersion, operationCount);
        }
        catch (DriverCatalogException ex)
        {
            var detail = SecretRedactor.Redact(ex.Message);
            checks.Add(Fail("operation-catalog", detail));
            errors.Add(detail);
            return Build(false, ExitCodes.AuthOrCatalog, checks, errors, username, driverBaseUrl, discoveryMethod, driverVersion, schemaVersion, operationCount);
        }
        catch (DriverConnectionException ex)
        {
            var detail = SecretRedactor.Redact(ex.Message);
            checks.Add(Fail("operation-catalog", detail));
            errors.Add(detail);
            return Build(false, ExitCodes.DriverUnavailable, checks, errors, username, driverBaseUrl, discoveryMethod, driverVersion, schemaVersion, operationCount);
        }

        _logger.LogInformation(
            "Doctor succeeded for {BaseUrl} schema={SchemaVersion} operations={OperationCount}",
            driverBaseUrl,
            schemaVersion,
            operationCount);

        return Build(true, ExitCodes.Success, checks, errors, username, driverBaseUrl, discoveryMethod, driverVersion, schemaVersion, operationCount);
    }

    private ReadinessReport Build(
        bool success,
        int exitCode,
        List<ReadinessCheck> checks,
        List<string> errors,
        string? username,
        string? driverBaseUrl,
        string? discoveryMethod,
        string? driverVersion,
        int? schemaVersion,
        int? operationCount) =>
        new()
        {
            Success = success,
            ExitCode = exitCode,
            Username = username,
            DriverBaseUrl = driverBaseUrl,
            DiscoveryMethod = discoveryMethod,
            DriverVersion = driverVersion,
            CatalogSchemaVersion = schemaVersion,
            OperationCount = operationCount,
            WorkspaceRoot = _workspace.RootPath,
            Checks = checks,
            Errors = errors.Select(SecretRedactor.Redact).ToArray()
        };

    private static ReadinessCheck Pass(string name, string detail) =>
        new() { Name = name, Status = ReadinessCheckStatus.Passed, Detail = detail };

    private static ReadinessCheck Fail(string name, string detail) =>
        new() { Name = name, Status = ReadinessCheckStatus.Failed, Detail = SecretRedactor.Redact(detail) };
}
