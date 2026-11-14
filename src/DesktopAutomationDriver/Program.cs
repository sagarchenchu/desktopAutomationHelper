using System.Net;
using System.Text;
using System.Text.Json;
using DesktopAutomationDriver.Middleware;
using DesktopAutomationDriver.Services;
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
builder.Services.AddSingleton<IOperatorPresenceService, InteractiveOperatorPresenceService>();
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

// Bearer-token authentication middleware.
// All routes except GET /verify require: Authorization: Bearer <token>
app.UseMiddleware<BearerTokenMiddleware>();

app.MapControllers();

// -----------------------------------------------------------------------
// Probe server on port 9102 (best-effort).
// Serves only GET /verify so that callers can discover which port and
// token to use.  If 9102 is already held by another user's driver the
// attempt is silently skipped and the flag stays false.
// -----------------------------------------------------------------------
using var probeCts = new CancellationTokenSource();
var probeTask = RunProbeServerAsync(driverContext, logger, probeCts.Token);

// Wait briefly so the probe-port result is known before printing the banner.
await Task.Delay(200);

LogStartupBanner(driverContext, logger);

app.Run();            // Blocks until Ctrl+C / shutdown signal
probeCts.Cancel();
await probeTask;

// =======================================================================
// Probe-server helper (HttpListener on port 9102)
// =======================================================================
static async Task RunProbeServerAsync(
    IDriverContext ctx, ILogger logger, CancellationToken ct)
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
            "Probe server bound to http://localhost:9102/verify");
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

                var payload = new
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
static void LogStartupBanner(IDriverContext ctx, ILogger logger)
{
    // Operator token must never appear in stdout/logs. Approval requires interactive presence.
    logger.LogInformation("{Banner}", StartupBanner.Format(ctx));
}
