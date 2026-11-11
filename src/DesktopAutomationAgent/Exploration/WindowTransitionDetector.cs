using System.Text.Json;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Observation;
using DesktopAutomationAgent.Plans;

namespace DesktopAutomationAgent.Exploration;

public sealed class WindowInventory
{
    public List<WindowDescriptor> Windows { get; init; } = [];
}

public sealed class WindowDescriptor
{
    public long? Hwnd { get; init; }

    public string? Title { get; init; }

    public string? ClassName { get; init; }

    public int? ProcessId { get; init; }

    public string? ProcessName { get; init; }
}

public sealed class WindowTransitionResult
{
    public bool ScopeViolation { get; init; }

    public bool NewWindow { get; init; }

    public bool ClosedWindow { get; init; }

    public bool ModalDialog { get; init; }

    public string? Detail { get; init; }

    public WindowInventory Inventory { get; init; } = new();

    public List<WindowDescriptor> AddedWindows { get; init; } = [];
}

/// <summary>
/// Detects new/closed windows and scope violations against the exploration policy.
/// New windows without processName fail closed so they can be process-checked.
/// </summary>
public sealed class WindowTransitionDetector
{
    private readonly IDriverUiClient _uiClient;
    private readonly ProcessIdentityValidator _identityValidator;

    public WindowTransitionDetector(
        IDriverUiClient uiClient,
        ProcessIdentityValidator? identityValidator = null)
    {
        _uiClient = uiClient;
        _identityValidator = identityValidator ?? new ProcessIdentityValidator();
    }

    public async Task<WindowInventory> ListWindowsAsync(
        DriverConnection connection,
        CancellationToken cancellationToken = default)
    {
        var step = new PlanStep
        {
            Id = "explore-listwindows",
            Operation = "listwindows",
            CaptureResponse = true,
            // Include owned/child top-level windows (second window, dialogs) that are not
            // always direct desktop children.
            Arguments = new Dictionary<string, System.Text.Json.JsonElement>(StringComparer.Ordinal)
            {
                ["includeDesktopDescendants"] = System.Text.Json.JsonSerializer.SerializeToElement(true)
            }
        };

        var response = await _uiClient.ExecuteStepAsync(connection, step, cancellationToken).ConfigureAwait(false);
        return ParseInventory(response.Value);
    }

    public WindowTransitionResult Compare(
        WindowInventory before,
        WindowInventory after,
        SemanticSnapshot? afterSnapshot,
        ExplorationPolicyDocument policy)
    {
        ArgumentNullException.ThrowIfNull(before);
        ArgumentNullException.ThrowIfNull(after);
        ArgumentNullException.ThrowIfNull(policy);

        var beforeKeys = before.Windows.Select(WindowKey).ToHashSet(StringComparer.Ordinal);
        var afterKeys = after.Windows.Select(WindowKey).ToHashSet(StringComparer.Ordinal);

        var added = after.Windows.Where(w => !beforeKeys.Contains(WindowKey(w))).ToList();
        var removed = before.Windows.Where(w => !afterKeys.Contains(WindowKey(w))).ToList();

        // Only newly appeared windows are scope-checked. Pre-existing desktop windows
        // (browser, terminals, etc.) are ignored so listwindows noise cannot fail the crawl.
        foreach (var window in added)
        {
            var identity = _identityValidator.ValidateNewWindow(window, policy);
            if (!identity.IsValid || !IsWindowInScope(window, policy))
            {
                return new WindowTransitionResult
                {
                    ScopeViolation = true,
                    NewWindow = true,
                    ClosedWindow = removed.Count > 0,
                    Detail = identity.Detail
                             ?? $"Unexpected external window '{window.Title}' / process '{window.ProcessName}'.",
                    Inventory = after
                };
            }
        }

        if (afterSnapshot?.Application.ProcessName is string processName)
        {
            if (string.IsNullOrWhiteSpace(processName) || !IsProcessAllowed(processName, policy))
            {
                return new WindowTransitionResult
                {
                    ScopeViolation = true,
                    Detail = $"Process '{processName}' is outside the allowed process tree.",
                    Inventory = after
                };
            }
        }
        else if (afterSnapshot is not null)
        {
            return new WindowTransitionResult
            {
                ScopeViolation = true,
                Detail = "Observed application processName is missing.",
                Inventory = after
            };
        }

        if (policy.Scope.WindowsSessionId is int expectedSession)
        {
            if (afterSnapshot?.WindowsSession.SessionId is not int actualSession)
            {
                return new WindowTransitionResult
                {
                    ScopeViolation = true,
                    Detail = "Windows session id is missing; fail-closed.",
                    Inventory = after
                };
            }

            if (expectedSession != actualSession)
            {
                return new WindowTransitionResult
                {
                    ScopeViolation = true,
                    Detail = $"Windows session {actualSession} differs from allowed session {expectedSession}.",
                    Inventory = after
                };
            }
        }

        if (afterSnapshot?.Window is { } active)
        {
            if (!IsTitleAllowed(active.Title, policy) || !IsClassAllowed(active.ClassName, policy))
            {
                return new WindowTransitionResult
                {
                    ScopeViolation = true,
                    Detail = $"Active window '{active.Title}' ({active.ClassName}) is out of scope.",
                    Inventory = after
                };
            }
        }

        return new WindowTransitionResult
        {
            NewWindow = added.Count > 0,
            ClosedWindow = removed.Count > 0,
            ModalDialog = afterSnapshot?.Window?.IsModal == true,
            Detail = added.Count > 0
                ? $"New window(s): {string.Join(", ", added.Select(w => w.Title))}"
                : removed.Count > 0
                    ? $"Closed window(s): {string.Join(", ", removed.Select(w => w.Title))}"
                    : null,
            Inventory = after,
            AddedWindows = added
        };
    }

    public static bool IsSnapshotInScope(SemanticSnapshot snapshot, ExplorationPolicyDocument policy)
    {
        if (string.IsNullOrWhiteSpace(snapshot.Application.ProcessName))
            return false;

        if (!IsProcessAllowed(snapshot.Application.ProcessName, policy))
            return false;

        if (policy.Scope.WindowsSessionId is int expectedSession)
        {
            if (snapshot.WindowsSession.SessionId is not int actualSession)
                return false;
            if (expectedSession != actualSession)
                return false;
        }

        if (snapshot.Window is { } window)
        {
            if (!IsTitleAllowed(window.Title, policy) || !IsClassAllowed(window.ClassName, policy))
                return false;
        }

        return true;
    }

    private static bool IsProcessAllowed(string processName, ExplorationPolicyDocument policy)
    {
        if (string.Equals(processName, policy.Application.ProcessName, StringComparison.OrdinalIgnoreCase))
            return true;

        return policy.Application.AllowedChildProcessNames
            .Any(child => string.Equals(child, processName, StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsWindowInScope(WindowDescriptor window, ExplorationPolicyDocument policy)
    {
        // Fail closed: new windows must carry processName so they can be process-checked.
        if (string.IsNullOrWhiteSpace(window.ProcessName))
            return false;

        if (!IsProcessAllowed(window.ProcessName, policy))
            return false;

        if (!IsTitleAllowed(window.Title, policy) || !IsClassAllowed(window.ClassName, policy))
            return false;

        return true;
    }

    private static bool IsTitleAllowed(string? title, ExplorationPolicyDocument policy)
    {
        var denied = policy.Scope.DeniedWindowTitles;
        if (denied.Count > 0
            && denied.Any(d => !string.IsNullOrWhiteSpace(title)
                               && title!.IndexOf(d, StringComparison.OrdinalIgnoreCase) >= 0))
            return false;

        var allowed = policy.Scope.AllowedWindowTitles;
        if (allowed.Count == 0)
            return true;

        return !string.IsNullOrWhiteSpace(title)
               && allowed.Any(a => string.Equals(a, title, StringComparison.OrdinalIgnoreCase)
                                   || title!.IndexOf(a, StringComparison.OrdinalIgnoreCase) >= 0);
    }

    private static bool IsClassAllowed(string? className, ExplorationPolicyDocument policy)
    {
        var denied = policy.Scope.DeniedWindowClasses;
        if (denied.Count > 0
            && denied.Any(d => string.Equals(d, className, StringComparison.OrdinalIgnoreCase)))
            return false;

        var allowed = policy.Scope.AllowedWindowClasses;
        if (allowed.Count == 0)
            return true;

        return allowed.Any(a => string.Equals(a, className, StringComparison.OrdinalIgnoreCase));
    }

    private static string WindowKey(WindowDescriptor window) =>
        $"{window.Hwnd?.ToString() ?? "-"}|{window.Title ?? "-"}|{window.ClassName ?? "-"}|{window.ProcessId?.ToString() ?? "-"}";

    private static WindowInventory ParseInventory(JsonElement? value)
    {
        var inventory = new WindowInventory();
        if (value is null || value.Value.ValueKind == JsonValueKind.Undefined || value.Value.ValueKind == JsonValueKind.Null)
            return inventory;

        var root = value.Value;
        if (root.ValueKind == JsonValueKind.Object && root.TryGetProperty("windows", out var windows))
            root = windows;
        else if (root.ValueKind == JsonValueKind.Object && root.TryGetProperty("Windows", out var windowsPascal))
            root = windowsPascal;

        if (root.ValueKind != JsonValueKind.Array)
            return inventory;

        foreach (var item in root.EnumerateArray())
        {
            inventory.Windows.Add(new WindowDescriptor
            {
                Hwnd = TryGetInt64(item, "hwnd") ?? TryGetInt64(item, "Hwnd"),
                Title = TryGetString(item, "title") ?? TryGetString(item, "Title") ?? TryGetString(item, "name"),
                ClassName = TryGetString(item, "className") ?? TryGetString(item, "ClassName"),
                ProcessId = TryGetInt32(item, "processId") ?? TryGetInt32(item, "ProcessId") ?? TryGetInt32(item, "pid"),
                ProcessName = TryGetString(item, "processName") ?? TryGetString(item, "ProcessName")
            });
        }

        return inventory;
    }

    private static string? TryGetString(JsonElement element, string name) =>
        element.ValueKind == JsonValueKind.Object
        && element.TryGetProperty(name, out var property)
        && property.ValueKind == JsonValueKind.String
            ? property.GetString()
            : null;

    private static int? TryGetInt32(JsonElement element, string name)
    {
        if (element.ValueKind != JsonValueKind.Object || !element.TryGetProperty(name, out var property))
            return null;
        if (property.ValueKind == JsonValueKind.Number && property.TryGetInt32(out var value))
            return value;
        if (property.ValueKind == JsonValueKind.String && int.TryParse(property.GetString(), out value))
            return value;
        return null;
    }

    private static long? TryGetInt64(JsonElement element, string name)
    {
        if (element.ValueKind != JsonValueKind.Object || !element.TryGetProperty(name, out var property))
            return null;
        if (property.ValueKind == JsonValueKind.Number && property.TryGetInt64(out var value))
            return value;
        if (property.ValueKind == JsonValueKind.String && long.TryParse(property.GetString(), out value))
            return value;
        return null;
    }
}
