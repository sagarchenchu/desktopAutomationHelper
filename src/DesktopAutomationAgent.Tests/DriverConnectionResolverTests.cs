using System.Net;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Driver.Bootstrap;
using Microsoft.Extensions.Logging.Abstractions;

namespace DesktopAutomationAgent.Tests;

public class DriverConnectionResolverTests
{
    [Fact]
    public async Task ExplicitConfiguration_Succeeds()
    {
        var options = TestSupport.CreateOptions(
            baseUrl: "http://127.0.0.1:33201",
            bearerToken: "secret-token");
        var resolver = new DriverConnectionResolver(
            TestSupport.Wrap(options),
            TestSupport.CreateFactory(new FakeHttpMessageHandler(_ => throw new InvalidOperationException("should not call verify"))),
            NullLogger<DriverConnectionResolver>.Instance);

        var connection = await resolver.ResolveAsync();

        Assert.Equal("explicit", connection.DiscoveryMethod);
        Assert.Equal("http://127.0.0.1:33201", connection.SafeBaseUrl);
        Assert.Equal("secret-token", connection.BearerToken);
    }

    [Theory]
    [InlineData("http://127.0.0.1:33201", null)]
    [InlineData(null, "secret-token")]
    [InlineData("http://127.0.0.1:33201", "")]
    [InlineData("", "secret-token")]
    public void PartialExplicitConfiguration_IsRejected(string? baseUrl, string? token)
    {
        var options = TestSupport.CreateOptions(baseUrl: baseUrl, bearerToken: token);
        var ex = Assert.Throws<AgentConfigurationException>(() => AgentOptionsValidator.ValidateDriverOptions(options.Driver));
        Assert.Contains("together", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task SecureDiscovery_IsDefaultAndUsesSecureBootstrap()
    {
        const string sid = "S-1-5-21-100-200-300-400";
        const string token = "secure-discovered-token";
        var processStart = DateTimeOffset.UtcNow;
        var secureClient = new SecureBootstrapClient(
            currentSid: () => sid,
            localAppData: () => Path.GetTempPath(),
            pipeReaderOverride: (_, _) => Task.FromResult<SecureBootstrapPayload?>(new SecureBootstrapPayload
            {
                SchemaVersion = SecureBootstrapConstants.SchemaVersion,
                UserSid = sid,
                Username = "alice",
                Port = 33201,
                Token = token,
                AuthorizationHeader = $"Bearer {token}",
                Pid = Environment.ProcessId,
                ProcessStartTimeUtc = processStart,
                StartedAtUtc = DateTimeOffset.UtcNow,
                PipeName = SecureBootstrapConstants.BuildPipeName(sid)
            }),
            descriptorReaderOverride: _ => null,
            getProcessStartTimeUtc: _ => processStart);

        var options = TestSupport.CreateOptions(); // DiscoveryMode defaults to Secure
        var resolver = new DriverConnectionResolver(
            TestSupport.Wrap(options),
            TestSupport.CreateFactory(new FakeHttpMessageHandler(_ => throw new InvalidOperationException("HTTP verify must not be used in Secure mode"))),
            NullLogger<DriverConnectionResolver>.Instance,
            () => "alice",
            secureClient);

        var connection = await resolver.ResolveAsync();

        Assert.Equal("secure-bootstrap", connection.DiscoveryMethod);
        Assert.Equal("http://127.0.0.1:33201", connection.SafeBaseUrl);
        Assert.Equal(token, connection.BearerToken);
        Assert.Equal("alice", connection.DiscoveredUsername);
    }

    [Fact]
    public async Task SecureDiscovery_PipeTimeoutWithoutCallerCancel_MapsToDriverConnectionException()
    {
        const string sid = "S-1-5-21-100-200-300-402";
        var secureClient = new SecureBootstrapClient(
            currentSid: () => sid,
            localAppData: () => Path.GetTempPath(),
            pipeReaderOverride: async (_, _) =>
            {
                // Simulate CancelAfter timeout: OCE while caller token is still open.
                var timedOut = new CancellationTokenSource();
                timedOut.Cancel();
                timedOut.Token.ThrowIfCancellationRequested();
                await Task.CompletedTask;
                return null;
            },
            descriptorReaderOverride: _ => null,
            getProcessStartTimeUtc: _ => null);

        var options = TestSupport.CreateOptions();
        var resolver = new DriverConnectionResolver(
            TestSupport.Wrap(options),
            TestSupport.CreateFactory(new FakeHttpMessageHandler(_ => throw new InvalidOperationException("unused"))),
            NullLogger<DriverConnectionResolver>.Instance,
            () => "alice",
            secureClient);

        var ex = await Assert.ThrowsAsync<DriverConnectionException>(() => resolver.ResolveAsync());
        Assert.Contains("secure bootstrap", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task SecureDiscovery_RejectsStaleDescriptor()
    {
        const string sid = "S-1-5-21-100-200-300-401";
        const string token = "stale-token-must-not-leak";
        var secureClient = new SecureBootstrapClient(
            currentSid: () => sid,
            localAppData: () => Path.GetTempPath(),
            pipeReaderOverride: (_, _) => Task.FromResult<SecureBootstrapPayload?>(null),
            descriptorReaderOverride: _ => new SecureBootstrapPayload
            {
                SchemaVersion = SecureBootstrapConstants.SchemaVersion,
                UserSid = sid,
                Username = "alice",
                Port = 33201,
                Token = token,
                Pid = 424242,
                ProcessStartTimeUtc = DateTimeOffset.UtcNow.AddMinutes(-10),
                StartedAtUtc = DateTimeOffset.UtcNow
            },
            getProcessStartTimeUtc: _ => null);

        var options = TestSupport.CreateOptions();
        var resolver = new DriverConnectionResolver(
            TestSupport.Wrap(options),
            TestSupport.CreateFactory(new FakeHttpMessageHandler(_ => throw new InvalidOperationException("unused"))),
            NullLogger<DriverConnectionResolver>.Instance,
            () => "alice",
            secureClient);

        var ex = await Assert.ThrowsAsync<DriverConnectionException>(() => resolver.ResolveAsync());
        Assert.Contains("stale", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(token, ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task LegacyVerifyDiscovery_SucceedsWhenExplicitlyConfigured()
    {
        var user = Environment.UserName;
        var handler = new FakeHttpMessageHandler(_ => FakeHttpMessageHandler.Json(new
        {
            status = 0,
            value = new
            {
                running = true,
                username = user,
                port = 33201,
                probePort = 9102,
                token = "discovered-token",
                authorizationHeader = "Bearer discovered-token"
            }
        }));

        var options = TestSupport.CreateOptions(
            discoveryMode: "LegacyVerify",
            acknowledgeLegacyVerifyRisk: true);
        var resolver = new DriverConnectionResolver(
            TestSupport.Wrap(options),
            TestSupport.CreateFactory(handler),
            NullLogger<DriverConnectionResolver>.Instance,
            () => user);

        var connection = await resolver.ResolveAsync();

        Assert.Equal("verify", connection.DiscoveryMethod);
        Assert.Equal("http://127.0.0.1:33201", connection.SafeBaseUrl);
        Assert.Equal("discovered-token", connection.BearerToken);
        Assert.Equal(user, connection.DiscoveredUsername);
    }

    [Fact]
    public void LegacyVerifyDiscovery_RequiresRiskAcknowledgement()
    {
        var options = TestSupport.CreateOptions(
            discoveryMode: "LegacyVerify",
            acknowledgeLegacyVerifyRisk: false);

        var ex = Assert.Throws<AgentConfigurationException>(() =>
            AgentOptionsValidator.ValidateDriverOptions(options.Driver));
        Assert.Contains("AcknowledgeLegacyVerifyRisk", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task LegacyVerifyDiscovery_RejectsUsernameMismatch()
    {
        var handler = new FakeHttpMessageHandler(_ => FakeHttpMessageHandler.Json(new
        {
            status = 0,
            value = new
            {
                running = true,
                username = "other-user",
                port = 33201,
                token = "should-not-use",
                authorizationHeader = "Bearer should-not-use"
            }
        }));

        var options = TestSupport.CreateOptions(
            discoveryMode: "LegacyVerify",
            acknowledgeLegacyVerifyRisk: true);
        var resolver = new DriverConnectionResolver(
            TestSupport.Wrap(options),
            TestSupport.CreateFactory(handler),
            NullLogger<DriverConnectionResolver>.Instance,
            () => "local-user");

        var ex = await Assert.ThrowsAsync<DriverConnectionException>(() => resolver.ResolveAsync());
        Assert.Contains("does not match", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("should-not-use", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void LoopbackEnforcement_RejectsRemoteByDefault()
    {
        var options = TestSupport.CreateOptions(
            baseUrl: "http://example.com:33201",
            bearerToken: "token");

        var ex = Assert.Throws<AgentConfigurationException>(() => AgentOptionsValidator.ValidateDriverOptions(options.Driver));
        Assert.Contains("Remote driver URL", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void LoopbackEnforcement_AllowsRemoteWhenEnabled()
    {
        var options = TestSupport.CreateOptions(
            baseUrl: "http://example.com:33201",
            bearerToken: "token",
            allowRemote: true);

        AgentOptionsValidator.ValidateDriverOptions(options.Driver);
    }

    [Fact]
    public void RemoteVerifyUrl_IsRejectedEvenWhenAllowRemoteDriverEnabled()
    {
        var options = TestSupport.CreateOptions(
            allowRemote: true,
            discoveryMode: "LegacyVerify",
            acknowledgeLegacyVerifyRisk: true,
            verifyUrl: "http://example.com:9102/verify");

        var ex = Assert.Throws<AgentConfigurationException>(() =>
            AgentOptionsValidator.ValidateDriverOptions(options.Driver));
        Assert.Contains("Remote driver URL", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("alice", "alice", true)]
    [InlineData("DOMAIN\\alice", "alice", true)]
    [InlineData("alice@contoso.com", "alice", true)]
    [InlineData("bob", "alice", false)]
    public void UsernameMatching_IsCaseInsensitiveAndDomainAware(string discovered, string current, bool expected)
    {
        Assert.Equal(expected, DriverConnectionResolver.UsernamesMatch(discovered, current));
    }
}
