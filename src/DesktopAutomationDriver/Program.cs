using System.Net;
using System.Text;
using System.Text.Json;
using DesktopAutomationDriver.Middleware;
using DesktopAutomationDriver.Services;

// -----------------------------------------------------------------------
// Create driver context first: determines the user-specific port and the
// one-time Bearer token before the web host is built.
// -----------------------------------------------------------------------
var driverContext = new DriverContext();

var builder = WebApplication.CreateBuilder(args);

// -----------------------------------------------------------------------
// Port: user-specific (derived from Windows login name) so that multiple
// users on the same shared Citrix machine never collide.
// Can still be overridden by passing --urls on the command line or via the
// ASPNETCORE_URLS environment variable for special cases.
// -----------------------------------------------------------------------
// Bind on all IPv4 interfaces (0.0.0.0) AND the IPv6 loopback (::1) so that
// both http://localhost:{port} and http://127.0.0.1:{port} work regardless of
// whether the OS resolves "localhost" to IPv4 or IPv6.
builder.WebHost.UseUrls(
    $"http://0.0.0.0:{driverContext.MainPort}",
    $"http://[::1]:{driverContext.MainPort}");

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
    // Bind both "localhost" (hostname, covers IPv6 when OS resolves it that way)
    // and "127.0.0.1" (explicit IPv4 loopback) so callers can use either form.
    const string prefix = "http://localhost:9102/";
    const string prefix127 = "http://127.0.0.1:9102/";
    using var listener = new HttpListener();
    listener.Prefixes.Add(prefix);
    listener.Prefixes.Add(prefix127);

    try
    {
        listener.Start();
        ctx.ProbePortActive = true;
        logger.LogInformation(
            "Probe server bound to http://localhost:9102/verify and http://127.0.0.1:9102/verify");
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
    var probeInfo = ctx.ProbePortActive
        ? $"http://localhost:{ctx.ProbePort}/verify  (active)"
        : $"port {ctx.ProbePort} unavailable — use main port for /verify";
    var probeInfo2 = ctx.ProbePortActive
        ? $"http://127.0.0.1:{ctx.ProbePort}/verify  (active)"
        : string.Empty;

    var p = ctx.MainPort;
    var sb = new System.Text.StringBuilder();
    sb.AppendLine();
    sb.AppendLine("╔══════════════════════════════════════════════════════════════╗");
    sb.AppendLine("║          Desktop Automation Driver — Started                 ║");
    sb.AppendLine("╠══════════════════════════════════════════════════════════════╣");
    sb.AppendLine($"║  Windows user  : {ctx.Username,-44}║");
    sb.AppendLine($"║  Main port     : {ctx.MainPort,-44}║");
    sb.AppendLine($"║  Bearer token  : {ctx.BearerToken,-44}║");
    sb.AppendLine("╠══════════════════════════════════════════════════════════════╣");
    sb.AppendLine("║  ENDPOINTS  (require: Authorization: Bearer <token>)         ║");
    sb.AppendLine("║  Accessible via localhost OR 127.0.0.1                       ║");
    sb.AppendLine("║                                                              ║");
    sb.AppendLine($"║  GET  http://localhost:{p}/verify    <- no auth, returns token");
    sb.AppendLine($"║  GET  http://127.0.0.1:{p}/verify");
    sb.AppendLine($"║  GET  http://localhost:{p}/status");
    sb.AppendLine($"║  GET  http://127.0.0.1:{p}/status");
    sb.AppendLine($"║  POST http://localhost:{p}/session");
    sb.AppendLine($"║  GET  http://localhost:{p}/sessions");
    sb.AppendLine($"║  GET  http://localhost:{p}/session/{{id}}");
    sb.AppendLine($"║  DEL  http://localhost:{p}/session/{{id}}");
    sb.AppendLine($"║  POST http://localhost:{p}/session/{{id}}/element");
    sb.AppendLine($"║  POST http://localhost:{p}/session/{{id}}/elements");
    sb.AppendLine("║  ...and all element action / query routes");
    sb.AppendLine("╠══════════════════════════════════════════════════════════════╣");
    sb.AppendLine("║  RECORDING (require: Authorization: Bearer <token>)          ║");
    sb.AppendLine("║                                                              ║");
    sb.AppendLine($"║  POST http://localhost:{p}/record/start  <- open overlay");
    sb.AppendLine($"║  GET  http://localhost:{p}/record/status");
    sb.AppendLine($"║  GET  http://localhost:{p}/record/actions");
    sb.AppendLine($"║  POST http://localhost:{p}/record/stop   <- export JSON");
    sb.AppendLine("╠══════════════════════════════════════════════════════════════╣");
    sb.AppendLine("║  PROBE (no auth)                                             ║");
    sb.AppendLine($"║  {probeInfo,-62}║");
    if (!string.IsNullOrEmpty(probeInfo2))
        sb.AppendLine($"║  {probeInfo2,-62}║");
    sb.AppendLine("╚══════════════════════════════════════════════════════════════╝");
    sb.AppendLine($"  Authorization header to use:");
    sb.AppendLine($"    Authorization: Bearer {ctx.BearerToken}");

    logger.LogInformation("{Banner}", sb.ToString());
}
