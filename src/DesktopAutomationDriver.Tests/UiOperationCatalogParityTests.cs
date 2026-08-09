using System.Text.Json;
using System.Text.RegularExpressions;
using DesktopAutomationDriver.Controllers;
using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Models.Response;
using DesktopAutomationDriver.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging.Abstractions;
using Moq;

namespace DesktopAutomationDriver.Tests;

public class UiOperationCatalogParityTests
{
    private readonly IUiOperationCatalog _catalog = new UiOperationCatalog();

    [Fact]
    public void Catalog_HasNoDuplicateCanonicalNames()
    {
        var names = _catalog.GetOperations().Select(o => o.Name.ToLowerInvariant()).ToList();
        Assert.Equal(names.Count, names.Distinct(StringComparer.Ordinal).Count());
    }

    [Fact]
    public void Catalog_AliasesAreUniqueAcrossOperations()
    {
        var owners = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var op in _catalog.GetOperations())
        {
            foreach (var alias in op.Aliases)
            {
                if (owners.TryGetValue(alias, out var existing))
                    Assert.Fail($"Alias '{alias}' claimed by both '{existing}' and '{op.Name}'.");
                owners[alias] = op.Name;
            }
        }
    }

    [Fact]
    public void Catalog_AliasDoesNotCollideWithCanonicalName()
    {
        var canonical = _catalog.GetOperations().Select(o => o.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var op in _catalog.GetOperations())
        {
            foreach (var alias in op.Aliases)
                Assert.False(canonical.Contains(alias), $"Alias '{alias}' collides with a canonical name.");
        }
    }

    [Fact]
    public void Catalog_MatchingIsCaseInsensitive()
    {
        Assert.True(_catalog.IsKnownOperation("CLICK"));
        Assert.True(_catalog.IsKnownOperation("SwitchWindow"));
        Assert.True(_catalog.IsKnownOperation("switchwinodw"));
        Assert.False(_catalog.IsKnownOperation("not-a-real-operation"));
    }

    [Fact]
    public void Catalog_DoesNotTrimOperationNames_MatchingUiService()
    {
        Assert.True(_catalog.IsKnownOperation("click"));
        Assert.False(_catalog.IsKnownOperation(" click "));
        Assert.False(_catalog.IsKnownOperation("click "));
        Assert.False(_catalog.IsKnownOperation(" click"));
    }

    [Fact]
    public void Catalog_CoversEveryUiServiceSwitchOperation()
    {
        var switchOps = ExtractUiServiceSwitchOperations();
        var recognized = _catalog.GetAllRecognizedNames();

        var missing = switchOps.Where(op => !recognized.Contains(op)).OrderBy(x => x).ToList();
        Assert.True(missing.Count == 0, "Catalog missing dispatcher ops: " + string.Join(", ", missing));
    }

    [Fact]
    public void Catalog_EveryEntryIsRecognizedByUiServiceDispatcher()
    {
        var switchOps = ExtractUiServiceSwitchOperations().ToHashSet(StringComparer.Ordinal);
        var unrecognized = _catalog.GetAllRecognizedNames()
            .Where(name => !switchOps.Contains(name))
            .OrderBy(x => x)
            .ToList();

        Assert.True(unrecognized.Count == 0,
            "Catalog contains names not in UiService switch: " + string.Join(", ", unrecognized));
    }

    [Fact]
    public void UiService_UnknownOperation_ThrowsArgumentException()
    {
        var ctx = new Mock<IUiSessionContext>();
        ctx.Setup(c => c.ActiveSession).Returns((AutomationSession?)null);
        var service = UiServiceTestFactory.Create(ctx.Object);

        var ex = Assert.Throws<ArgumentException>(() =>
            service.Execute(new UiRequest { Operation = "definitely-unknown-op-xyz" }));

        Assert.Contains("Unknown operation", ex.Message);
        Assert.Contains("/ui/operations", ex.Message);
    }

    [Fact]
    public void UiService_DeprecatedAlias_IsRecognized()
    {
        var ctx = new Mock<IUiSessionContext>();
        ctx.Setup(c => c.ActiveSession).Returns((AutomationSession?)null);
        var service = UiServiceTestFactory.Create(ctx.Object);

        var ex = Record.Exception(() =>
            service.Execute(new UiRequest { Operation = "switchwinodw", Value = "Title" }));

        Assert.False(ex is ArgumentException ae && ae.Message.Contains("Unknown operation"),
            "Deprecated alias switchwinodw must remain recognized.");
    }

    [Fact]
    public void GetOperationsEndpoint_ReturnsStableJsonShape()
    {
        var ui = new Mock<IUiService>(MockBehavior.Strict);
        var config = new Mock<IConfiguration>();
        config.Setup(c => c["FailureScreenshotDirectory"]).Returns((string?)null);

        var controller = new UiController(ui.Object, _catalog, NullLogger<UiController>.Instance, config.Object)
        {
            ControllerContext = new ControllerContext { HttpContext = new DefaultHttpContext() }
        };

        var result = controller.GetOperations();
        var ok = Assert.IsType<OkObjectResult>(result);
        var envelope = Assert.IsType<UiResponse>(ok.Value);
        Assert.True(envelope.Success);

        var json = JsonSerializer.Serialize(envelope, new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        });

        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;
        Assert.True(root.GetProperty("success").GetBoolean());
        var value = root.GetProperty("value");
        Assert.Equal(2, value.GetProperty("schemaVersion").GetInt32());
        Assert.True(value.TryGetProperty("driverVersion", out var version));
        Assert.False(string.IsNullOrWhiteSpace(version.GetString()));
        Assert.Equal(JsonValueKind.Array, value.GetProperty("operations").ValueKind);

        var first = value.GetProperty("operations").EnumerateArray().First();
        Assert.True(first.TryGetProperty("name", out _));
        Assert.True(first.TryGetProperty("aliases", out _));
        Assert.True(first.TryGetProperty("category", out _));
        Assert.True(first.TryGetProperty("operationType", out _));
        Assert.True(first.TryGetProperty("requiresSession", out _));
        Assert.True(first.TryGetProperty("requiredInputs", out _));
        Assert.True(first.TryGetProperty("requiredInputAlternatives", out _));
        Assert.True(first.TryGetProperty("deprecated", out _));

        var names = value.GetProperty("operations").EnumerateArray()
            .Select(o => o.GetProperty("name").GetString()!)
            .ToList();
        Assert.Equal(names.OrderBy(n => n, StringComparer.Ordinal).ToList(), names);
    }

    [Fact]
    public void Switchwinodw_IsMarkedDeprecatedAlias()
    {
        var switchWindow = _catalog.GetOperations()
            .Single(o => o.Name.Equals("switchwindow", StringComparison.OrdinalIgnoreCase));

        Assert.Contains(switchWindow.Aliases, a => a.Equals("switchwinodw", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(switchWindow.DeprecatedAliases, a => a.Equals("switchwinodw", StringComparison.OrdinalIgnoreCase));
        Assert.False(switchWindow.Deprecated);
        Assert.False(switchWindow.RequiresSession);
        Assert.Contains(switchWindow.RequiredInputAlternatives, alt => alt.SequenceEqual(new[] { "value" }));
        Assert.Contains(switchWindow.RequiredInputAlternatives, alt => alt.SequenceEqual(new[] { "hwnd" }));
        Assert.Contains(switchWindow.RequiredInputAlternatives, alt => alt.SequenceEqual(new[] { "className" }));
    }

    [Fact]
    public void Catalog_Select_HasValueOrIndexAlternatives()
    {
        var select = _catalog.GetOperations().Single(o => o.Name == "select");
        Assert.Equal(new[] { "locator" }, select.RequiredInputs);
        Assert.Contains(select.RequiredInputAlternatives, alt => alt.SequenceEqual(new[] { "locator", "value" }));
        Assert.Contains(select.RequiredInputAlternatives, alt => alt.SequenceEqual(new[] { "locator", "index" }));
    }

    [Fact]
    public void Catalog_Metadata_MatchesReviewedHandlers()
    {
        var launch = _catalog.GetOperations().Single(o => o.Name == "launch");
        Assert.Equal(new[] { "value" }, launch.RequiredInputs);

        var isEditable = _catalog.GetOperations().Single(o => o.Name == "iseditable");
        Assert.Equal("element-query", isEditable.Category);
        Assert.Equal("query", isEditable.OperationType);

        var inspect = _catalog.GetOperations().Single(o => o.Name == "inspectcombobox");
        Assert.Equal("diagnostic", inspect.OperationType);

        var findUia = _catalog.GetOperations().Single(o => o.Name == "finduia");
        Assert.False(findUia.RequiresSession);

        var sendKeysUia = _catalog.GetOperations().Single(o => o.Name == "sendkeysuia");
        Assert.True(sendKeysUia.RequiresSession);

        var getPosition = _catalog.GetOperations().Single(o => o.Name == "getposition");
        Assert.Equal(new[] { "locator", "locator2" }, getPosition.RequiredInputs);

        var contextMenu = _catalog.GetOperations().Single(o => o.Name == "contextmenupath");
        Assert.Equal(new[] { "locator", "value" }, contextMenu.RequiredInputs);
    }

    private static List<string> ExtractUiServiceSwitchOperations()
    {
        var path = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory, "..", "..", "..", "..",
            "DesktopAutomationDriver", "Services", "UiService.cs"));
        Assert.True(File.Exists(path), $"Expected UiService.cs at {path}");

        var content = File.ReadAllText(path).Replace("\r\n", "\n");
        var match = Regex.Match(
            content,
            @"request\.Operation\.ToLowerInvariant\(\)\s*switch\s*\{(.*?)\n\s*\};",
            RegexOptions.Singleline);
        Assert.True(match.Success, "Could not locate UiService.Execute switch expression.");

        return Regex.Matches(match.Groups[1].Value, "\"([^\"]+)\"\\s*=>")
            .Select(m => m.Groups[1].Value.ToLowerInvariant())
            .Distinct(StringComparer.Ordinal)
            .OrderBy(x => x, StringComparer.Ordinal)
            .ToList();
    }
}
