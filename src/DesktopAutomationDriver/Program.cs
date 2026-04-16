using DesktopAutomationDriver.Services;

var builder = WebApplication.CreateBuilder(args);

// -------------------------------------------------------------------
// Configuration: port can be overridden via --urls, ASPNETCORE_URLS,
// or the "DriverPort" config key (defaults to 4723).
// -------------------------------------------------------------------
var port = builder.Configuration.GetValue<int>("DriverPort", 4723);
builder.WebHost.UseUrls($"http://0.0.0.0:{port}");

// -------------------------------------------------------------------
// Services
// -------------------------------------------------------------------
builder.Services.AddControllers()
    .AddJsonOptions(options =>
    {
        // Return JSON with camelCase property names (WebDriver convention)
        options.JsonSerializerOptions.PropertyNamingPolicy =
            System.Text.Json.JsonNamingPolicy.CamelCase;
        options.JsonSerializerOptions.DefaultIgnoreCondition =
            System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull;
    });

builder.Services.AddSingleton<ISessionManager, SessionManager>();
builder.Services.AddSingleton<IAutomationService, AutomationService>();

builder.Logging.AddConsole();

// -------------------------------------------------------------------
// Application pipeline
// -------------------------------------------------------------------
var app = builder.Build();

app.MapControllers();

Console.WriteLine($"Desktop Automation Driver listening on http://0.0.0.0:{port}");
Console.WriteLine("Press Ctrl+C to stop.");

app.Run();
