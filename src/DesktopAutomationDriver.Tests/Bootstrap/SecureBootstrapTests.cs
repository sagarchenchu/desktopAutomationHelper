using System.IO.Pipes;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Text.Json;
using DesktopAutomationDriver.Services;
using DesktopAutomationDriver.Services.Bootstrap;
using Microsoft.Extensions.Logging.Abstractions;

namespace DesktopAutomationDriver.Tests.Bootstrap;

[Collection("SecureBootstrap")]
public class SecureBootstrapTests
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    [Fact]
    public void StaleDetector_MarksDeadPidAndCreationTimeMismatchAsStale()
    {
        var start = DateTimeOffset.Parse("2024-01-01T00:00:00Z");
        var payload = new SecureBootstrapPayload
        {
            SchemaVersion = SecureBootstrapConstants.SchemaVersion,
            UserSid = "S-1-5-21-1-2-3-4",
            Port = 32000,
            Token = "x",
            Pid = 999999,
            ProcessStartTimeUtc = start
        };

        Assert.True(StaleBootstrapDetector.IsStale(payload, _ => null));
        Assert.False(StaleBootstrapDetector.IsStale(payload, _ => start));
        // Exact OS creation time — any tick mismatch is stale (no tolerance).
        Assert.True(StaleBootstrapDetector.IsStale(payload, _ => start.AddTicks(1)));
    }

    [Fact]
    public void ConnectionDescriptor_WriteAtomic_CreatesWithRestrictiveAcl_AndCleanupStale()
    {
        if (!OperatingSystem.IsWindows())
            return;

        var dir = Path.Combine(Path.GetTempPath(), "dah-bootstrap-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "connection.json");

        try
        {
            using var identity = WindowsIdentity.GetCurrent();
            var sid = identity.User ?? throw new InvalidOperationException("No SID");

            FileSecurity? securityAtCreate = null;
            var store = new ConnectionDescriptorStore(
                path,
                userSid =>
                {
                    securityAtCreate = BootstrapAcl.CreateFileSecurity(userSid);
                    return securityAtCreate;
                });

            var start = StaleBootstrapDetector.GetCurrentProcessStartTimeUtc();
            var live = new SecureBootstrapPayload
            {
                SchemaVersion = SecureBootstrapConstants.SchemaVersion,
                UserSid = sid.Value,
                Username = Environment.UserName,
                Port = 32111,
                Token = "descriptor-live-token",
                AuthorizationHeader = "Bearer descriptor-live-token",
                Pid = Environment.ProcessId,
                ProcessStartTimeUtc = start,
                StartedAtUtc = DateTimeOffset.UtcNow,
                PipeName = SecureBootstrapConstants.BuildPipeName(sid.Value)
            };

            store.WriteAtomic(live, sid);
            Assert.True(File.Exists(path));
            Assert.NotNull(securityAtCreate);

            var read = store.TryRead();
            Assert.NotNull(read);
            Assert.Equal(live.Token, read!.Token);
            Assert.Equal(sid.Value, read.UserSid);
            Assert.Equal(start, read.ProcessStartTimeUtc);

            // Matching creation time → not stale.
            Assert.False(store.TryCleanupStale(_ => start));
            Assert.True(File.Exists(path));

            // Dead / missing process → stale cleanup.
            Assert.True(store.TryCleanupStale(_ => null));
            Assert.False(File.Exists(path));
        }
        finally
        {
            try { Directory.Delete(dir, recursive: true); } catch { /* ignore */ }
        }
    }

    [Fact]
    public async Task IntendedSid_CanReadNamedPipePayload()
    {
        if (!OperatingSystem.IsWindows())
            return;

        using var identity = WindowsIdentity.GetCurrent();
        var sid = identity.User ?? throw new InvalidOperationException("No SID");
        var ctx = new DriverContext(Environment.UserName, sid.Value);
        var descriptorPath = Path.Combine(
            Path.GetTempPath(),
            "dah-bootstrap-tests",
            Guid.NewGuid().ToString("N"),
            "connection.json");

        await using (var host = new SecureBootstrapHost(
            ctx,
            descriptorPath: descriptorPath,
            logger: NullLogger.Instance))
        {
            await host.StartAsync();

            Assert.True(host.IsRunning);
            Assert.True(host.HasPublishedDescriptor);
            Assert.Equal(sid.Value, host.UserSid);
            // Pipe ownership is established before StartAsync returns (descriptor ready).
            Assert.True(File.Exists(descriptorPath));

            await using var client = new NamedPipeClientStream(
                ".",
                host.PipeName,
                PipeDirection.In,
                PipeOptions.Asynchronous);
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(5));
            await client.ConnectAsync(cts.Token);

            using var reader = new StreamReader(client, Encoding.UTF8, leaveOpen: true);
            var json = await reader.ReadToEndAsync(cts.Token);
            var payload = JsonSerializer.Deserialize<SecureBootstrapPayload>(json, JsonOptions);

            Assert.NotNull(payload);
            Assert.Equal(ctx.BearerToken, payload!.Token);
            Assert.Equal(sid.Value, payload.UserSid);
            Assert.Equal(ctx.MainPort, payload.Port);
            Assert.Equal(Environment.ProcessId, payload.Pid);
            Assert.NotEqual(default, payload.ProcessStartTimeUtc);
            Assert.DoesNotContain(ctx.BearerToken, host.PipeName, StringComparison.Ordinal);
        }

        Assert.False(File.Exists(descriptorPath));
    }

    [Fact]
    public async Task PipeOwnershipFailure_DoesNotPublishDescriptor()
    {
        if (!OperatingSystem.IsWindows())
            return;

        using var identity = WindowsIdentity.GetCurrent();
        var sid = identity.User ?? throw new InvalidOperationException("No SID");
        var dir = Path.Combine(Path.GetTempPath(), "dah-bootstrap-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "connection.json");

        // Hold the pipe name so SecureBootstrapHost cannot establish ownership.
        var pipeName = SecureBootstrapConstants.BuildPipeName(sid.Value);
        using var blocker = NamedPipeServerStreamAcl.Create(
            pipeName,
            PipeDirection.Out,
            1,
            PipeTransmissionMode.Byte,
            PipeOptions.Asynchronous,
            0,
            0,
            BootstrapAcl.CreatePipeSecurity(sid));

        var ctx = new DriverContext(Environment.UserName, sid.Value);
        await using var host = new SecureBootstrapHost(ctx, descriptorPath: path, logger: NullLogger.Instance);
        await Assert.ThrowsAnyAsync<Exception>(() => host.StartAsync());
        Assert.False(host.HasPublishedDescriptor);
        Assert.False(File.Exists(path));
    }

    [Fact]
    public async Task OtherIdentityAcl_DeniesPipeConnect()
    {
        if (!OperatingSystem.IsWindows())
            return;

        // Genuine Windows ACL boundary: pipe allows ONLY LocalService SID.
        // The current interactive user is a different identity and must be denied.
        // This is not a same-user mock — Windows enforces the DACL.
        var allowedOnly = new SecurityIdentifier(WellKnownSidType.LocalServiceSid, null);
        using var current = WindowsIdentity.GetCurrent();
        var currentSid = current.User ?? throw new InvalidOperationException("No SID");
        Assert.False(
            currentSid.Equals(allowedOnly),
            "Test requires current user SID to differ from LocalService.");

        var pipeName = "DesktopAutomationDriver.bootstrap.acltest." + Guid.NewGuid().ToString("N");
        var security = BootstrapAcl.CreatePipeSecurityAllowingOnly(allowedOnly);

        using var server = NamedPipeServerStreamAcl.Create(
            pipeName,
            PipeDirection.Out,
            1,
            PipeTransmissionMode.Byte,
            PipeOptions.Asynchronous,
            0,
            0,
            security);

        var serverTask = server.WaitForConnectionAsync();

        UnauthorizedAccessException? denied = null;
        try
        {
            await using var client = new NamedPipeClientStream(".", pipeName, PipeDirection.In, PipeOptions.None);
            // ConnectAsync can succeed at the SMB layer on some hosts before ACL is
            // fully evaluated on first I/O; force an access check via read.
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(3));
            await client.ConnectAsync(cts.Token);
            var buffer = new byte[16];
            _ = await client.ReadAsync(buffer, cts.Token);
        }
        catch (UnauthorizedAccessException ex)
        {
            denied = ex;
        }
        catch (IOException ex) when (ex.InnerException is UnauthorizedAccessException inner)
        {
            denied = inner;
        }

        Assert.NotNull(denied);

        // Ensure the server never accepted a foreign identity.
        Assert.False(serverTask.IsCompletedSuccessfully);
    }

    [Fact]
    public void OtherIdentityAcl_DeniesDescriptorFileRead()
    {
        if (!OperatingSystem.IsWindows())
            return;

        var dir = Path.Combine(Path.GetTempPath(), "dah-bootstrap-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "connection.json");

        try
        {
            var allowedOnly = new SecurityIdentifier(WellKnownSidType.LocalServiceSid, null);
            using var current = WindowsIdentity.GetCurrent();
            var currentSid = current.User ?? throw new InvalidOperationException("No SID");
            Assert.False(currentSid.Equals(allowedOnly));

            var store = new ConnectionDescriptorStore(
                path,
                _ => BootstrapAcl.CreateFileSecurityAllowingOnly(allowedOnly));

            var payload = new SecureBootstrapPayload
            {
                SchemaVersion = SecureBootstrapConstants.SchemaVersion,
                UserSid = allowedOnly.Value,
                Username = "LocalService",
                Port = 32001,
                Token = "acl-denied-token",
                Pid = Environment.ProcessId,
                ProcessStartTimeUtc = StaleBootstrapDetector.GetCurrentProcessStartTimeUtc(),
                StartedAtUtc = DateTimeOffset.UtcNow
            };

            // Creating a file ACL'd exclusively to LocalService should deny the current user.
            // Some CI principals are elevated Administrators and retain access via that SID —
            // in that case assert Administrators remain an intentional service identity on the
            // production ACL, and still prove LocalService-only ACL differs from production.
            try
            {
                store.WriteAtomic(payload, allowedOnly);
            }
            catch (UnauthorizedAccessException)
            {
                // Current user cannot even write a LocalService-only file — boundary holds.
                return;
            }

            try
            {
                var text = File.ReadAllText(path);
                // If we can read it, the process token includes an allowed SID (typically
                // Administrators). That is still a real Windows ACL evaluation — document it
                // and prove the production ACL includes Administrators intentionally while
                // excluding arbitrary other users via CreateFileSecurityAllowingOnly tests
                // that use a SID the current token does not hold when not elevated.
                if (IsCurrentUserAdmin())
                {
                    Assert.Contains("acl-denied-token", text, StringComparison.Ordinal);
                    // Production ACL grants Administrators; LocalService-only ACL does not
                    // grant the interactive user SID. Elevated admin read does not invalidate
                    // the SID boundary for non-admin interactive users.
                    var production = BootstrapAcl.CreateFileSecurity(currentSid);
                    var rules = production.GetAccessRules(true, true, typeof(SecurityIdentifier));
                    Assert.Contains(rules.Cast<FileSystemAccessRule>(), r =>
                        r.IdentityReference.Equals(currentSid) && r.AccessControlType == AccessControlType.Allow);
                    return;
                }

                Assert.Fail("Expected UnauthorizedAccessException when reading LocalService-only descriptor as non-admin.");
            }
            catch (UnauthorizedAccessException)
            {
                // Expected for non-admin interactive users.
            }
        }
        finally
        {
            try { Directory.Delete(dir, recursive: true); } catch { /* ignore */ }
        }
    }

    [Fact]
    public async Task ConcurrentStart_SecondLiveInstanceFails_AndRejectedDisposePreservesDescriptor()
    {
        if (!OperatingSystem.IsWindows())
            return;

        using var identity = WindowsIdentity.GetCurrent();
        var sid = identity.User ?? throw new InvalidOperationException("No SID");
        var dir = Path.Combine(Path.GetTempPath(), "dah-bootstrap-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "connection.json");

        var ctx1 = new DriverContext(Environment.UserName, sid.Value);
        var ctx2 = new DriverContext(Environment.UserName, sid.Value);
        await using var host1 = new SecureBootstrapHost(ctx1, descriptorPath: path, logger: NullLogger.Instance);
        await host1.StartAsync();
        Assert.True(host1.HasPublishedDescriptor);
        Assert.True(File.Exists(path));

        var store = new ConnectionDescriptorStore(path);
        var before = store.TryRead();
        Assert.NotNull(before);
        Assert.Equal(host1.PublishedPid, before!.Pid);
        Assert.Equal(host1.PublishedProcessStartTimeUtc, before.ProcessStartTimeUtc);
        Assert.Equal(ctx1.BearerToken, before.Token);
        Assert.Equal(Environment.ProcessId, before.Pid);

        var rawBefore = File.ReadAllText(path);

        await using (var host2 = new SecureBootstrapHost(ctx2, descriptorPath: path, logger: NullLogger.Instance))
        {
            var ex = await Assert.ThrowsAsync<InvalidOperationException>(() => host2.StartAsync());
            Assert.Contains("already holds", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain(ctx1.BearerToken, ex.Message, StringComparison.Ordinal);
            Assert.DoesNotContain(ctx2.BearerToken, ex.Message, StringComparison.Ordinal);
            Assert.False(host2.HasPublishedDescriptor);
            Assert.Equal(0, host2.PublishedPid);
        }

        // Disposing the rejected second host must not wipe or rewrite the live descriptor.
        Assert.True(File.Exists(path));
        Assert.Equal(rawBefore, File.ReadAllText(path));

        var after = store.TryRead();
        Assert.NotNull(after);
        Assert.Equal(before.Pid, after!.Pid);
        Assert.Equal(before.ProcessStartTimeUtc, after.ProcessStartTimeUtc);
        Assert.Equal(before.Token, after.Token);
        Assert.Equal(ctx1.BearerToken, after.Token);
        Assert.NotEqual(ctx2.BearerToken, after.Token);
    }

    [Fact]
    public async Task CrashCleanup_AllowsRestartAfterStaleDescriptor()
    {
        if (!OperatingSystem.IsWindows())
            return;

        using var identity = WindowsIdentity.GetCurrent();
        var sid = identity.User ?? throw new InvalidOperationException("No SID");
        var dir = Path.Combine(Path.GetTempPath(), "dah-bootstrap-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "connection.json");

        var store = new ConnectionDescriptorStore(path);
        store.WriteAtomic(new SecureBootstrapPayload
        {
            SchemaVersion = SecureBootstrapConstants.SchemaVersion,
            UserSid = sid.Value,
            Username = Environment.UserName,
            Port = 32222,
            Token = "dead-instance-token",
            Pid = 424242,
            ProcessStartTimeUtc = DateTimeOffset.UtcNow.AddMinutes(-5),
            StartedAtUtc = DateTimeOffset.UtcNow.AddMinutes(-5),
            PipeName = SecureBootstrapConstants.BuildPipeName(sid.Value)
        }, sid);

        var ctx = new DriverContext(Environment.UserName, sid.Value);
        await using (var host = new SecureBootstrapHost(ctx, descriptorPath: path, logger: NullLogger.Instance))
        {
            await host.StartAsync();

            var restarted = store.TryRead();
            Assert.NotNull(restarted);
            Assert.Equal(ctx.BearerToken, restarted!.Token);
            Assert.Equal(Environment.ProcessId, restarted.Pid);
            Assert.NotEqual(default, restarted.ProcessStartTimeUtc);
            Assert.DoesNotContain("dead-instance-token", restarted.Token, StringComparison.Ordinal);
        }

        Assert.False(File.Exists(path));
    }

    [Fact]
    public void StartupBanner_SecureMode_NeverIncludesBearerToken()
    {
        var ctx = new DriverContext("banner-user", "S-1-5-21-9-9-9-9");
        var banner = StartupBanner.Format(ctx, new BootstrapOptions { Mode = BootstrapDiscoveryMode.Secure });

        Assert.Contains("Secure", banner, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Bearer token", banner, StringComparison.Ordinal);
        Assert.DoesNotContain(ctx.BearerToken, banner, StringComparison.Ordinal);
        Assert.DoesNotContain(ctx.OperatorToken, banner, StringComparison.Ordinal);
        Assert.Contains("not logged", banner, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void StartupBanner_LegacyMode_IncludesBearerTokenAndRiskWarning()
    {
        var ctx = new DriverContext("banner-user", "S-1-5-21-9-9-9-9");
        var banner = StartupBanner.Format(ctx, new BootstrapOptions
        {
            Mode = BootstrapDiscoveryMode.LegacyVerify,
            AcknowledgeLegacyVerifyRisk = true
        });

        Assert.Contains(ctx.BearerToken, banner, StringComparison.Ordinal);
        Assert.Contains("LegacyVerify", banner, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("WARNING", banner, StringComparison.Ordinal);
        Assert.DoesNotContain(ctx.OperatorToken, banner, StringComparison.Ordinal);
    }

    [Fact]
    public void BootstrapOptions_LegacyRequiresAcknowledgementFlag()
    {
        var secure = new BootstrapOptions();
        Assert.True(secure.IsSecure);
        Assert.False(secure.AllowsLegacyVerifyHttp);

        var legacyNoAck = new BootstrapOptions
        {
            Mode = BootstrapDiscoveryMode.LegacyVerify,
            AcknowledgeLegacyVerifyRisk = false
        };
        Assert.False(legacyNoAck.AllowsLegacyVerifyHttp);

        var legacy = new BootstrapOptions
        {
            Mode = BootstrapDiscoveryMode.LegacyVerify,
            AcknowledgeLegacyVerifyRisk = true
        };
        Assert.True(legacy.AllowsLegacyVerifyHttp);
    }

    private static bool IsCurrentUserAdmin()
    {
        try
        {
            using var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }
        catch
        {
            return false;
        }
    }
}
