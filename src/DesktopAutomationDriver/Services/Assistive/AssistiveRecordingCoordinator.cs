using DesktopAutomationDriver.Models.Recording;

namespace DesktopAutomationDriver.Services.Assistive;

internal enum BddArmState
{
    None,
    ArmedNextAction,
    ActiveUntilFinished
}

internal sealed class BddPendingContext
{
    public required string GroupId { get; init; }
    public required string Statement { get; init; }
    public required BddScope Scope { get; init; }
    public DateTimeOffset CreatedAt { get; init; } = DateTimeOffset.UtcNow;
    public int? FirstAssociatedSequence { get; set; }
    public bool Completed { get; set; }
    public int AssociatedActionCount { get; set; }
}

/// <summary>
/// Thread-safe Assistive Jira/BDD and event enrichment coordinator.
/// Owned by <see cref="RecordingService"/> and protected by the recording lock.
/// </summary>
public sealed class AssistiveRecordingCoordinator
{
    private string? _jiraKey;
    private int? _jiraScopeStartSequence;
    private int _nextSequence;
    private int _nextBddGroupNumber;
    private BddArmState _bddState = BddArmState.None;
    private BddPendingContext? _pendingBdd;
    private readonly Dictionary<string, string> _normalizedTitleToPageId = new(StringComparer.Ordinal);
    private readonly Dictionary<string, HashSet<string>> _pageElementIds = new(StringComparer.Ordinal);
    private readonly Dictionary<string, string> _locatorKeyToObjectRef = new(StringComparer.Ordinal);
    private AssistiveActionCaptureContext? _pendingCaptureContext;

    public string? RecordingId { get; private set; }

    public string? JiraKey => _jiraKey;

    public int? JiraScopeStartSequence => _jiraScopeStartSequence;

    public BddScope? ActiveBddScope => _pendingBdd?.Scope;

    public string? ActiveBddGroupId => _pendingBdd is { Completed: false } ? _pendingBdd.GroupId : null;

    public int ActiveBddStatementLength => _pendingBdd?.Statement.Length ?? 0;

    public string OverlayStatusSuffix
    {
        get
        {
            if (string.IsNullOrWhiteSpace(_jiraKey))
                return string.Empty;

            if (_pendingBdd is { Completed: false } pending)
            {
                var arm = pending.Scope == BddScope.NextAction
                    ? "armed for next action"
                    : "active for multiple actions";
                return $"Jira {_jiraKey} | BDD {pending.GroupId} {arm}";
            }

            return $"Jira {_jiraKey}";
        }
    }

    public void Reset(string recordingId)
    {
        RecordingId = recordingId;
        _jiraKey = null;
        _jiraScopeStartSequence = null;
        _nextSequence = 0;
        _nextBddGroupNumber = 0;
        _bddState = BddArmState.None;
        _pendingBdd = null;
        _normalizedTitleToPageId.Clear();
        _pageElementIds.Clear();
        _locatorKeyToObjectRef.Clear();
        _pendingCaptureContext = null;
    }

    public void SetPendingCaptureContext(AssistiveActionCaptureContext? context) =>
        _pendingCaptureContext = context;

    public AssistiveActionCaptureContext? TakePendingCaptureContext()
    {
        var ctx = _pendingCaptureContext;
        _pendingCaptureContext = null;
        return ctx;
    }

    public bool TryStartJiraRecording(string? rawKey, out string canonical, out string error)
    {
        canonical = string.Empty;
        error = string.Empty;

        if (_jiraScopeStartSequence is not null)
        {
            error = "Jira key is locked for this recording because a Jira-scoped action was already recorded. Stop and start a new recording to use another key.";
            return false;
        }

        if (!string.IsNullOrWhiteSpace(_jiraKey))
        {
            // Allow replacing the key only before any Jira-scoped action exists.
            if (!JiraKeyRules.TryCanonicalize(rawKey, out canonical, out error))
                return false;

            _jiraKey = canonical;
            return true;
        }

        if (!JiraKeyRules.TryCanonicalize(rawKey, out canonical, out error))
            return false;

        _jiraKey = canonical;
        return true;
    }

    public bool TryArmBdd(string? rawStatement, BddScope scope, out string groupId, out string error)
    {
        groupId = string.Empty;
        error = string.Empty;

        if (string.IsNullOrWhiteSpace(_jiraKey))
        {
            error = "Start Jira recording before entering a BDD statement.";
            return false;
        }

        if (!BddStatementRules.TryNormalize(rawStatement, out var statement, out error))
            return false;

        _nextBddGroupNumber++;
        groupId = $"bdd-{_nextBddGroupNumber:D4}";
        _pendingBdd = new BddPendingContext
        {
            GroupId = groupId,
            Statement = statement,
            Scope = scope
        };
        _bddState = scope == BddScope.NextAction
            ? BddArmState.ArmedNextAction
            : BddArmState.ActiveUntilFinished;
        return true;
    }

    public bool TryFinishBdd(out string message)
    {
        if (_pendingBdd is null || _bddState == BddArmState.None)
        {
            message = "No active BDD statement to finish.";
            return false;
        }

        _pendingBdd.Completed = true;
        _bddState = BddArmState.None;
        message = $"Finished BDD {_pendingBdd.GroupId}.";
        _pendingBdd = null;
        return true;
    }

    public bool TryCancelBdd(out string message)
    {
        if (_pendingBdd is null || _bddState == BddArmState.None)
        {
            message = "No pending BDD statement to cancel.";
            return false;
        }

        if (_pendingBdd.AssociatedActionCount > 0)
        {
            message = $"BDD {_pendingBdd.GroupId} already has recorded actions and cannot be cancelled. Use Finish instead.";
            return false;
        }

        message = $"Cancelled BDD {_pendingBdd.GroupId}.";
        _pendingBdd = null;
        _bddState = BddArmState.None;
        return true;
    }

    public void EnrichAssistiveAction(RecordedAction action, AssistiveActionCaptureContext? context)
    {
        _nextSequence++;
        action.Sequence = _nextSequence;
        action.EventId = $"evt-{_nextSequence:D6}";
        action.Mode = RecordingMode.Assistive;

        var window = context?.Window;
        var pageId = context?.PageId;
        if (string.IsNullOrWhiteSpace(pageId))
        {
            var normalized = DeterministicPageIdGenerator.NormalizeTitle(window?.Title);
            pageId = DeterministicPageIdGenerator.FromWindowTitle(window?.Title, _normalizedTitleToPageId);
            _normalizedTitleToPageId[normalized] = pageId;
            if (window != null)
                window.NormalizedTitle = string.IsNullOrWhiteSpace(window.NormalizedTitle)
                    ? normalized
                    : window.NormalizedTitle;
        }
        else if (window != null)
        {
            var normalized = DeterministicPageIdGenerator.NormalizeTitle(window.Title);
            _normalizedTitleToPageId[normalized] = pageId!;
            window.NormalizedTitle ??= normalized;
        }

        action.Window = window;
        action.PageId = pageId;

        if (!string.IsNullOrWhiteSpace(_jiraKey))
        {
            action.JiraKey = _jiraKey;
            _jiraScopeStartSequence ??= action.Sequence;
        }

        if (!string.IsNullOrWhiteSpace(pageId))
        {
            action.ObjectRef = ResolveObjectRef(pageId!, action.Element);
            if (action.TargetElement != null)
                action.TargetObjectRef = ResolveObjectRef(pageId!, action.TargetElement);
        }

        var deferBdd = context?.DeferBddConsumption == true;
        if (!deferBdd)
            AttachAndMaybeConsumeBdd(action);
    }

    private void AttachAndMaybeConsumeBdd(RecordedAction action)
    {
        if (_pendingBdd is null || _bddState == BddArmState.None || _pendingBdd.Completed)
            return;

        action.Bdd = new RecordedBddAssociation
        {
            GroupId = _pendingBdd.GroupId,
            Statement = _pendingBdd.Statement
        };

        _pendingBdd.FirstAssociatedSequence ??= action.Sequence;
        _pendingBdd.AssociatedActionCount++;

        if (_bddState == BddArmState.ArmedNextAction)
        {
            _pendingBdd.Completed = true;
            _pendingBdd = null;
            _bddState = BddArmState.None;
        }
    }

    private string? ResolveObjectRef(string pageId, ElementInfo? element)
    {
        if (element == null)
            return null;

        if (!HasUsableLocator(element))
            return null;

        var key = LocatorKey(pageId, element);
        if (_locatorKeyToObjectRef.TryGetValue(key, out var existing))
            return existing;

        if (!_pageElementIds.TryGetValue(pageId, out var used))
        {
            used = new HashSet<string>(StringComparer.Ordinal);
            _pageElementIds[pageId] = used;
        }

        var elementId = DeterministicElementIdGenerator.Resolve(element, used);
        var objectRef = $"{pageId}.{elementId}";
        _locatorKeyToObjectRef[key] = objectRef;
        return objectRef;
    }

    private static bool HasUsableLocator(ElementInfo element) =>
        !string.IsNullOrWhiteSpace(element.AutomationId)
        || (!string.IsNullOrWhiteSpace(element.Name) && !string.IsNullOrWhiteSpace(element.ControlType))
        || (!string.IsNullOrWhiteSpace(element.ClassName) && !string.IsNullOrWhiteSpace(element.ControlType));

    private static string LocatorKey(string pageId, ElementInfo element) =>
        string.Join(
            '\u001f',
            pageId,
            Normalize(element.AutomationId),
            Normalize(element.Name),
            Normalize(element.ClassName),
            Normalize(element.ControlType));

    private static string Normalize(string? value) =>
        string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim().ToLowerInvariant();
}
