using System.Net;
using System.Text;
using System.Text.Json;
using DesktopAutomationDriver.Middleware;
using DesktopAutomationDriver.Services;
using DesktopAutomationDriver.Services.Bootstrap;
using DesktopAutomationDriver.Services.Debug;
using DesktopAutomationDriver.Services.NativeUia;
using DesktopAutomationDriver.Services.Vision;
using WinForms = System.Windows.Forms.Application;

WinForms.SetHighDpiMode(System.Windows.Forms.HighDpiMode.PerMonitorV2);
WinForms.EnableVisualStyles();
WinForms.SetCompatibleTextRenderingDefault(false);

// -----------------------------------------------------------------------
// Create driver context first: determines the user-specific port and the
// one-time Bearer token before the web host is built.
// -----------------------------------------------------------------------
var driverContext = new DriverContext();

var builder = WebApplication.CreateBuilder(args);

// -----------------------------------------------------------------------
// Port: user-specific (derived from Windows login name) so that multiple
// users on the same shared Citrix machine never collide.
// Bind only on the IPv4 loopback (127.0.0.1) so that the driver is not
// reachable from outside the machine. Can still be overridden via
// ASPNETCORE_URLS or --urls when needed.
// -----------------------------------------------------------------------
builder.WebHost.UseUrls($"http://127.0.0.1:{driverContext.MainPort}");

// -----------------------------------------------------------------------
// Services
// -----------------------------------------------------------------------
builder.Services.AddControllers()
    .AddJsonOptions(options =>
    {
        options.JsonSerializerOptions.PropertyNamingPolicy =
            System.Text.Json.JsonNamingPolicy.CamelCase;
        options.JsonSerializerOptions.DefaultIgnoreCondition =
            System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull;
    });

builder.Services.Configure<BootstrapOptions>(builder.Configuration.GetSection(BootstrapOptions.SectionName));
builder.Services.AddSingleton<IDriverContext>(driverContext);
builder.Services.AddSingleton<ISessionManager, SessionManager>();
builder.Services.AddSingleton<IAutomationService, AutomationService>();
builder.Services.AddSingleton<IRecordingService, RecordingService>();
builder.Services.AddSingleton<IPlaybackService, PlaybackService>();
builder.Services.AddSingleton<IUiSessionContext, UiSessionContext>();
builder.Services.AddSingleton<INativeUiaComboBoxService, NativeUiaComboBoxService>();
builder.Services.AddSingleton<INativeUiaBasicOperationService, NativeUiaBasicOperationService>();
builder.Services.AddSingleton<INativeUiaGridService, NativeUiaGridService>();
builder.Services.AddSingleton<INativeUiaTreeDiagnosticService, NativeUiaTreeDiagnosticService>();
builder.Services.AddSingleton<IUiOperationCatalog, UiOperationCatalog>();
builder.Services.AddSingleton<IUiService, UiService>();

// Phase 8 — Debug Mode + vision fallback (isolated services; overlay stays UI/menu only).
// Vision is config-gated (Vision:Enabled, default false). Production HTTP: /vision/*.
builder.Services.Configure<VisionFallbackOptions>(builder.Configuration.GetSection("Vision"));
builder.Services.AddSingleton(sp =>
{
    var options = sp.GetRequiredService<Microsoft.Extensions.Options.IOptions<VisionFallbackOptions>>().Value
                  ?? new VisionFallbackOptions();
    // Fail closed: never enable vision unless the bound config explicitly sets Enabled=true.
    return options;
});
builder.Services.AddSingleton<VisionFallbackPolicy>();
builder.Services.AddSingleton<VisionScreenshotService>();
builder.Services.AddSingleton<IVisionProvider, FakeVisionProvider>();
builder.Services.AddSingleton(sp =>
    new VisionCandidateStore(sp.GetRequiredService<VisionFallbackOptions>()));
builder.Services.AddSingleton<ILiveWindowProbe, Win32LiveWindowProbe>();
builder.Services.AddSingleton<IWindowPixelCapture, ScreenWindowPixelCapture>();
builder.Services.AddSingleton<IVisionCaptureService, ServerVisionCaptureService>();
builder.Services.AddSingleton<IOperatorPresenceService>(sp =>
    new InteractiveOperatorPresenceService(sp.GetRequiredService<VisionFallbackOptions>()));
builder.Services.AddSingleton<CoordinateActionGuard>();
builder.Services.AddSingleton<VisionFallbackCoordinator>();
builder.Services.AddSingleton<IUiaElementProbe, FlaUIUiaElementProbe>();
builder.Services.AddSingleton<UiaTreeCaptureService>();
builder.Services.AddSingleton<LocatorCandidateGenerator>();
builder.Services.AddSingleton<CandidateVerificationService>();
builder.Services.AddSingleton<DebugArtifactWriter>();
builder.Services.AddSingleton<IHighlightSurface, WinFormsHighlightSurface>();
builder.Services.AddSingleton<NonInteractiveHighlightService>();
builder.Services.AddSingleton<IDebugSessionCoordinator, DebugSessionCoordinator>();

builder.Logging.AddConsole();

// -----------------------------------------------------------------------
// Build the application
// -----------------------------------------------------------------------
var app = builder.Build();
var logger = app.Services.GetRequiredService<ILogger<Program>>();
var bootstrapOptions = app.Services.GetRequiredService<Microsoft.Extensions.Options.IOptions<BootstrapOptions>>().Value
                       ?? new BootstrapOptions();

ValidateBootstrapOptions(bootstrapOptions, logger);

// Bearer-token authentication middleware.
// All routes except GET /verify require: Authorization: Bearer <token>
app.UseMiddleware<BearerTokenMiddleware>();

app.MapControllers();

// -----------------------------------------------------------------------
// Secure per-user bootstrap (SID ACL named pipe + connection descriptor).
// Always published so agents can discover without HTTP token return.
// -----------------------------------------------------------------------
await using var secureBootstrap = new SecureBootstrapHost(driverContext, logger: logger);
await secureBootstrap.StartAsync();

// -----------------------------------------------------------------------
// Probe server on port 9102 (best-effort).
// In LegacyVerify mode it may return the token. In Secure mode it refuses
// token return (same contract as VerifyController).
// -----------------------------------------------------------------------
using var probeCts = new CancellationTokenSource();
var probeTask = RunProbeServerAsync(driverContext, bootstrapOptions, logger, probeCts.Token);

// Wait briefly so the probe-port result is known before printing the banner.
await Task.Delay(200);

LogStartupBanner(driverContext, bootstrapOptions, logger);

try
{
    app.Run();            // Blocks until Ctrl+C / shutdown signal
}
finally
{
    probeCts.Cancel();
    await probeTask;
}

// =======================================================================
// Bootstrap option validation
// =======================================================================
static void ValidateBootstrapOptions(BootstrapOptions options, ILogger logger)
{
    if (options.Mode == BootstrapDiscoveryMode.LegacyVerify && !options.AcknowledgeLegacyVerifyRisk)
    {
        throw new InvalidOperationException(
            "Bootstrap:Mode=LegacyVerify also requires Bootstrap:AcknowledgeLegacyVerifyRisk=true. " +
            "Legacy HTTP /verify returns the bearer token without authentication and is unsafe on " +
            "shared Citrix/RDS/VDI hosts. Prefer Bootstrap:Mode=Secure (default).");
    }

    if (options.AllowsLegacyVerifyHttp)
    {
        logger.LogWarning(
            "LegacyVerify bootstrap is enabled: unauthenticated HTTP /verify returns the bearer token. " +
            "This is unsafe on shared hosts. See docs/shared-host-driver-discovery-security.md.");
    }
}

// =======================================================================
// Probe-server helper (HttpListener on port 9102)
// =======================================================================
static async Task RunProbeServerAsync(
    IDriverContext ctx,
    BootstrapOptions bootstrap,
    ILogger logger,
    CancellationToken ct)
{
    // Bind only on "localhost:9102" for the /verify probe endpoint.
    const string prefix = "http://localhost:9102/";
    using var listener = new HttpListener();
    listener.Prefixes.Add(prefix);

    try
    {
        listener.Start();
        ctx.ProbePortActive = true;
        logger.LogInformation(
            bootstrap.AllowsLegacyVerifyHttp
                ? "Probe server bound to http://localhost:9102/verify (legacy token return enabled)"
                : "Probe server bound to http://localhost:9102/verify (secure mode — token not returned)");
    }
    catch (HttpListenerException ex)
    {
        logger.LogWarning(
            "Could not bind probe port 9102 (likely held by another user's driver): {Msg}",
            ex.Message);
        return;
    }

    ct.Register(() => { try { listener.Stop(); } catch { /* best effort */ } });

    var jsonOpts = new JsonSerializerOptions
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition =
            System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
    };

    while (!ct.IsCancellationRequested)
    {
        HttpListenerContext? httpCtx;
        try { httpCtx = await listener.GetContextAsync(); }
        catch { break; }

        _ = Task.Run(async () =>
        {
            var req = httpCtx.Request;
            var resp = httpCtx.Response;

            try
            {
                // Only serve GET /verify; return 404 for everything else.
                var requestedPath = req.Url!.AbsolutePath.TrimEnd('/');
                if (!requestedPath.Equals("/verify", StringComparison.OrdinalIgnoreCase))
                {
                    resp.StatusCode = 404;
                    resp.Close();
                    return;
                }

                object payload;
                if (!bootstrap.AllowsLegacyVerifyHttp)
                {
                    resp.StatusCode = 403;
                    payload = new
                    {
                        status = 403,
                        value = new
                        {
                            error = "secure_bootstrap_required",
                            message =
                                "Unauthenticated HTTP /verify does not return the bearer token in secure mode. " +
                                "Use the per-user SID ACL-protected named pipe or connection descriptor."
                        }
                    };
                }
                else
                {
                    payload = new
                    {
                        status = 0,
                        value = new
                        {
                            running = true,
                            username = ctx.Username,
                            port = ctx.MainPort,
                            probePort = ctx.ProbePortActive ? (int?)ctx.ProbePort : null,
                            token = ctx.BearerToken,
                            authorizationHeader = $"Bearer {ctx.BearerToken}"
                        }
                    };
                }

                var json = JsonSerializer.Serialize(payload, jsonOpts);
                var bytes = Encoding.UTF8.GetBytes(json);
                resp.ContentType = "application/json; charset=utf-8";
                resp.ContentLength64 = bytes.Length;
                await resp.OutputStream.WriteAsync(bytes, ct);
            }
            finally
            {
                resp.Close();
            }
        }, ct);
    }
}

// =======================================================================
// Startup banner
// =======================================================================
static void LogStartupBanner(IDriverContext ctx, BootstrapOptions bootstrap, ILogger logger)
{
    // Secure mode: bearer token must never appear in stdout/logs.
    logger.LogInformation("{Banner}", StartupBanner.Format(ctx, bootstrap));
}
