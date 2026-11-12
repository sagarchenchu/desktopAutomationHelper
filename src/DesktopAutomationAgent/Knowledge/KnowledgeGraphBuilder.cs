using System.Text.RegularExpressions;
using DesktopAutomationAgent.Exploration;

namespace DesktopAutomationAgent.Knowledge;

/// <summary>
/// Converts exploration runs into a deterministic application knowledge document.
/// Merge is order-independent for identical semantic content once runs are sorted.
/// </summary>
public sealed class KnowledgeGraphBuilder
{
    private static readonly Regex ElementKeyPattern = new(
        @"^(?<control>[^|]+)\|(?<automationId>[^|]*)\|(?<name>[^|]*)",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private readonly ApplicationFingerprintBuilder _fingerprintBuilder = new();

    public ApplicationKnowledgeDocument Build(
        string applicationId,
        IReadOnlyList<LoadedExplorationRun> runs,
        DateTimeOffset? createdAtUtc = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(applicationId);
        ArgumentNullException.ThrowIfNull(runs);

        if (runs.Count == 0)
            throw new ArgumentException("At least one exploration run is required.", nameof(runs));
        if (runs.Count > KnowledgeConstants.MaxExplorationRuns)
            throw new InvalidOperationException($"At most {KnowledgeConstants.MaxExplorationRuns} exploration runs may be merged.");

        RejectDuplicateRuns(runs);

        var ordered = runs
            .OrderBy(r => r.ExplorationId, StringComparer.Ordinal)
            .ThenBy(r => r.ArtifactHash, StringComparer.Ordinal)
            .ToList();

        ValidateCompatibleFingerprints(applicationId, ordered, out var fingerprint, out var buildFingerprint);

        var ids = new KnowledgeStableIdFactory();
        var windows = new Dictionary<string, KnowledgeWindowNode>(StringComparer.Ordinal);
        var states = new Dictionary<string, KnowledgeStateNode>(StringComparer.Ordinal);
        var elements = new Dictionary<string, KnowledgeElementNode>(StringComparer.Ordinal);
        var transitions = new Dictionary<string, KnowledgeTransitionEdge>(StringComparer.Ordinal);
        var unresolved = new List<KnowledgeUnresolvedElement>();
        var ambiguous = new List<KnowledgeAmbiguousElement>();
        var dialogs = new Dictionary<string, KnowledgeDialogRelation>(StringComparer.Ordinal);
        var children = new Dictionary<string, KnowledgeChildWindowRelation>(StringComparer.Ordinal);
        var truncation = new SortedSet<string>(StringComparer.Ordinal);
        var deniedCount = 0;
        var warnings = new List<string>();

        foreach (var run in ordered)
        {
            deniedCount += run.Summary.DeniedActionSamples?.Count ?? 0;
            if (!string.IsNullOrWhiteSpace(run.Summary.StopReason))
                truncation.Add(run.Summary.StopReason!);
            foreach (var warning in run.Summary.Warnings ?? [])
            {
                if (!string.IsNullOrWhiteSpace(warning))
                    truncation.Add(warning);
            }

            var stateIdByFingerprint = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var state in run.States.OrderBy(s => s.StateId, StringComparer.Ordinal))
            {
                var windowId = ids.WindowId(state.WindowTitle, state.WindowClassName, state.IsModal == true);
                MergeWindow(windows, windowId, state, run.ExplorationId);

                var stateId = ids.StateId(state.Fingerprint);
                stateIdByFingerprint[state.Fingerprint] = stateId;
                MergeState(states, stateId, windowId, state, run.ExplorationId);
            }

            // Entry = first recorded state in the run (depth 0 preferred).
            var entry = run.States
                .OrderBy(s => s.Depth)
                .ThenBy(s => s.CapturedAtUtc)
                .ThenBy(s => s.StateId, StringComparer.Ordinal)
                .FirstOrDefault();
            if (entry is not null && stateIdByFingerprint.TryGetValue(entry.Fingerprint, out var entryId))
                states[entryId].IsEntry = true;

            foreach (var transition in run.Transitions.OrderBy(t => t.TransitionId, StringComparer.Ordinal))
            {
                if (!stateIdByFingerprint.TryGetValue(transition.FromFingerprint, out var sourceStateId))
                {
                    warnings.Add($"{run.ExplorationId}: transition {transition.TransitionId} missing source state fingerprint.");
                    continue;
                }

                string? targetStateId = null;
                if (!string.IsNullOrWhiteSpace(transition.ToFingerprint))
                {
                    if (!stateIdByFingerprint.TryGetValue(transition.ToFingerprint, out targetStateId))
                    {
                        // Target may be a newly observed fingerprint not present if write truncated;
                        // synthesize a minimal state so the edge remains explicit.
                        targetStateId = ids.StateId(transition.ToFingerprint);
                        if (!states.ContainsKey(targetStateId))
                        {
                            states[targetStateId] = new KnowledgeStateNode
                            {
                                StateId = targetStateId,
                                SemanticFingerprint = transition.ToFingerprint!,
                                WindowId = states[sourceStateId].WindowId,
                                FirstObservedAtUtc = transition.CompletedAtUtc,
                                LastObservedAtUtc = transition.CompletedAtUtc,
                                ObservationCount = 1,
                                Provenance = [run.ExplorationId]
                            };
                        }
                    }
                }

                string? elementId = null;
                if (!IsBacktrack(transition))
                {
                    var identity = ParseElementIdentity(transition);
                    if (identity is null)
                    {
                        unresolved.Add(new KnowledgeUnresolvedElement
                        {
                            Reason = "unparseable-element-key",
                            RawKey = transition.ElementKey,
                            Name = transition.Name
                        });
                    }
                    else if (IsNameOnlyWeakIdentity(identity))
                    {
                        ambiguous.Add(new KnowledgeAmbiguousElement
                        {
                            Reason = "name-only-identity-when-stronger-attributes-preferred",
                            CandidateKeys = [transition.ElementKey]
                        });
                        elementId = ids.ElementId(identity);
                        MergeElement(elements, elementId, identity, transition, run.ExplorationId);
                    }
                    else
                    {
                        elementId = ids.ElementId(identity);
                        MergeElement(elements, elementId, identity, transition, run.ExplorationId);
                    }
                }

                var normalizedAction = KnowledgeStableIdFactory.NormalizeActionKey(transition.ActionKey);
                var outcome = CamelOutcome(transition.Outcome);
                var semanticTransitionId = ids.TransitionId(sourceStateId, normalizedAction, targetStateId, outcome);

                MergeTransition(
                    transitions,
                    ids,
                    semanticTransitionId,
                    sourceStateId,
                    targetStateId,
                    elementId,
                    normalizedAction,
                    transition,
                    outcome,
                    run.ExplorationId);

                if (string.Equals(outcome, "newWindow", StringComparison.Ordinal)
                    && targetStateId is not null
                    && states.TryGetValue(sourceStateId, out var src)
                    && states.TryGetValue(targetStateId, out var dst)
                    && !string.Equals(src.WindowId, dst.WindowId, StringComparison.Ordinal))
                {
                    var key = src.WindowId + "->" + dst.WindowId;
                    if (!children.TryGetValue(key, out var rel))
                    {
                        children[key] = new KnowledgeChildWindowRelation
                        {
                            ParentWindowId = src.WindowId,
                            ChildWindowId = dst.WindowId,
                            ObservationCount = 1
                        };
                    }
                    else
                    {
                        rel.ObservationCount++;
                    }
                }

                if (string.Equals(outcome, "modalDialog", StringComparison.Ordinal)
                    && targetStateId is not null
                    && states.TryGetValue(sourceStateId, out var owner)
                    && states.TryGetValue(targetStateId, out var dialog))
                {
                    var key = owner.WindowId + "=>" + dialog.WindowId;
                    if (!dialogs.TryGetValue(key, out var rel))
                    {
                        dialogs[key] = new KnowledgeDialogRelation
                        {
                            OwnerWindowId = owner.WindowId,
                            DialogWindowId = dialog.WindowId,
                            ObservationCount = 1
                        };
                    }
                    else
                    {
                        rel.ObservationCount++;
                    }

                    if (windows.TryGetValue(dialog.WindowId, out var dialogWindow))
                        dialogWindow.Role = "dialog";
                }
            }
        }

        MarkTerminalStates(states, transitions.Values);

        // Detect conflicting transitions: same source+action, different target/outcome.
        AnnotateConflicts(transitions, ids);

        EnforceLimits(windows, states, elements, transitions, unresolved, ambiguous);

        var graph = new ApplicationStateGraph
        {
            GraphSchemaVersion = KnowledgeConstants.GraphSchemaVersion,
            Windows = windows.Values.OrderBy(w => w.WindowId, StringComparer.Ordinal).ToList(),
            States = states.Values.OrderBy(s => s.StateId, StringComparer.Ordinal).ToList(),
            Elements = elements.Values.OrderBy(e => e.ElementId, StringComparer.Ordinal).ToList(),
            Transitions = transitions.Values.OrderBy(t => t.TransitionId, StringComparer.Ordinal).ToList(),
            EntryStateIds = states.Values.Where(s => s.IsEntry).Select(s => s.StateId).OrderBy(s => s, StringComparer.Ordinal).ToList(),
            TerminalStateIds = states.Values.Where(s => s.IsTerminal).Select(s => s.StateId).OrderBy(s => s, StringComparer.Ordinal).ToList(),
            Dialogs = dialogs.Values.OrderBy(d => d.DialogWindowId, StringComparer.Ordinal).ThenBy(d => d.OwnerWindowId, StringComparer.Ordinal).ToList(),
            ChildWindows = children.Values.OrderBy(c => c.ChildWindowId, StringComparer.Ordinal).ThenBy(c => c.ParentWindowId, StringComparer.Ordinal).ToList(),
            UnresolvedElements = unresolved
                .OrderBy(u => u.RawKey, StringComparer.Ordinal)
                .ThenBy(u => u.Reason, StringComparer.Ordinal)
                .Take(5000)
                .ToList(),
            AmbiguousElements = ambiguous
                .OrderBy(a => a.Reason, StringComparer.Ordinal)
                .ThenBy(a => string.Join('|', a.CandidateKeys), StringComparer.Ordinal)
                .Take(5000)
                .ToList()
        };

        var workflows = BuildWorkflows(graph, ids, ordered);
        var coverage = KnowledgeCoverageBuilder.Build(graph, Enumerable.Range(0, deniedCount).Select(_ => "denied"), truncation);
        coverage.DeniedSamples = deniedCount;

        var created = createdAtUtc ?? DateTimeOffset.UnixEpoch;
        var knowledgeId = "kn_" + KnowledgeHash.StableIdPrefix(
            KnowledgeHash.Sha256Hex(string.Join('|', ordered.Select(r => r.ExplorationId).Append(buildFingerprint))),
            16);

        return new ApplicationKnowledgeDocument
        {
            KnowledgeSchemaVersion = KnowledgeConstants.KnowledgeSchemaVersion,
            KnowledgeId = knowledgeId,
            ApplicationId = applicationId.Trim().ToLowerInvariant(),
            BuildFingerprint = buildFingerprint,
            Fingerprint = fingerprint,
            Status = "candidate",
            CreatedAtUtc = created,
            UpdatedAtUtc = created,
            SourceExplorationRuns = ordered.Select(r => new SourceExplorationRun
            {
                ExplorationId = r.ExplorationId,
                ArtifactHash = r.ArtifactHash,
                PolicySha256 = r.PolicySha256,
                // Workspace-relative only — never machine-specific absolute paths.
                ManifestPath = ToWorkspaceRelativeManifest(r.RunDirectory)
            }).ToList(),
            Graph = graph,
            Workflows = workflows,
            Coverage = coverage,
            IdCollisions = ids.Collisions.ToList(),
            Warnings = warnings.Distinct(StringComparer.Ordinal).OrderBy(w => w, StringComparer.Ordinal).ToList()
        };
    }

    private static void RejectDuplicateRuns(IReadOnlyList<LoadedExplorationRun> runs)
    {
        var byId = new HashSet<string>(StringComparer.Ordinal);
        var byHash = new HashSet<string>(StringComparer.Ordinal);
        foreach (var run in runs)
        {
            if (!byId.Add(run.ExplorationId))
            {
                throw new InvalidOperationException(
                    $"Duplicate exploration run id '{run.ExplorationId}' cannot be merged "
                    + "(would inflate observation counts).");
            }

            if (!byHash.Add(run.ArtifactHash))
            {
                throw new InvalidOperationException(
                    $"Duplicate exploration artifact hash '{run.ArtifactHash}' "
                    + $"(run '{run.ExplorationId}') cannot be merged.");
            }
        }
    }

    private void ValidateCompatibleFingerprints(
        string applicationId,
        IReadOnlyList<LoadedExplorationRun> ordered,
        out ApplicationFingerprint fingerprint,
        out string buildFingerprint)
    {
        ApplicationFingerprint? baseline = null;
        string? baselineBuild = null;

        foreach (var run in ordered)
        {
            var fp = _fingerprintBuilder.BuildFromIdentity(applicationId, run.BuildIdentity);

            var build = _fingerprintBuilder.ComputeBuildFingerprint(fp);
            if (baseline is null)
            {
                baseline = fp;
                baselineBuild = build;
                continue;
            }

            if (!_fingerprintBuilder.AreCompatible(baseline, fp, out var mismatches))
            {
                throw new InvalidOperationException(
                    "Incompatible application fingerprints cannot be merged: "
                    + string.Join(", ", mismatches)
                    + $". Runs '{ordered[0].ExplorationId}' and '{run.ExplorationId}' differ. "
                    + "A build-fingerprint change starts a new knowledge version.");
            }
        }

        fingerprint = baseline!;
        buildFingerprint = baselineBuild!;
    }

    private static string ToWorkspaceRelativeManifest(string runDirectory)
    {
        var normalized = runDirectory.Replace('\\', '/').TrimEnd('/');
        var marker = "/exploration/runs/";
        var idx = normalized.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (idx >= 0)
            return "exploration/runs/" + normalized[(idx + marker.Length)..] + "/manifest.json";
        return "exploration/runs/" + Path.GetFileName(normalized) + "/manifest.json";
    }

    private static void MergeWindow(
        Dictionary<string, KnowledgeWindowNode> windows,
        string windowId,
        ExplorationStateRecord state,
        string explorationId)
    {
        if (!windows.TryGetValue(windowId, out var existing))
        {
            windows[windowId] = new KnowledgeWindowNode
            {
                WindowId = windowId,
                Title = state.WindowTitle ?? string.Empty,
                ClassName = NormalizeVolatileClassName(state.WindowClassName),
                IsModal = state.IsModal == true,
                Role = state.IsModal == true ? "dialog" : "unknown",
                FirstObservedAtUtc = state.CapturedAtUtc,
                LastObservedAtUtc = state.CapturedAtUtc,
                ObservationCount = 1
            };
            return;
        }

        existing.ObservationCount++;
        existing.FirstObservedAtUtc = Min(existing.FirstObservedAtUtc, state.CapturedAtUtc);
        existing.LastObservedAtUtc = Max(existing.LastObservedAtUtc, state.CapturedAtUtc);
        if (existing.IsModal || state.IsModal == true)
        {
            existing.IsModal = true;
            existing.Role = "dialog";
        }

        _ = explorationId;
    }

    private static void MergeState(
        Dictionary<string, KnowledgeStateNode> states,
        string stateId,
        string windowId,
        ExplorationStateRecord state,
        string explorationId)
    {
        if (!states.TryGetValue(stateId, out var existing))
        {
            states[stateId] = new KnowledgeStateNode
            {
                StateId = stateId,
                SemanticFingerprint = state.Fingerprint,
                WindowId = windowId,
                Label = state.WindowTitle,
                FirstObservedAtUtc = state.CapturedAtUtc,
                LastObservedAtUtc = state.CapturedAtUtc,
                ObservationCount = 1,
                Provenance = [explorationId],
                ObservedValues =
                {
                    ["nodeCount"] = state.NodeCount.ToString()
                }
            };
            return;
        }

        existing.ObservationCount++;
        existing.FirstObservedAtUtc = Min(existing.FirstObservedAtUtc, state.CapturedAtUtc);
        existing.LastObservedAtUtc = Max(existing.LastObservedAtUtc, state.CapturedAtUtc);
        if (!existing.Provenance.Contains(explorationId, StringComparer.Ordinal))
            existing.Provenance.Add(explorationId);
        existing.Provenance.Sort(StringComparer.Ordinal);
    }

    private static void MergeElement(
        Dictionary<string, KnowledgeElementNode> elements,
        string elementId,
        KnowledgeElementIdentity identity,
        ExplorationTransitionRecord transition,
        string explorationId)
    {
        if (!elements.TryGetValue(elementId, out var existing))
        {
            var locator = new Dictionary<string, object?>(StringComparer.Ordinal);
            if (!string.IsNullOrWhiteSpace(identity.AutomationId))
                locator["automationId"] = identity.AutomationId;
            if (!string.IsNullOrWhiteSpace(identity.Name))
                locator["name"] = identity.Name;
            locator["controlType"] = identity.ControlType;
            if (!string.IsNullOrWhiteSpace(identity.ClassName))
                locator["className"] = identity.ClassName;

            elements[elementId] = new KnowledgeElementNode
            {
                ElementId = elementId,
                Identity = identity,
                Capabilities = string.IsNullOrWhiteSpace(transition.ActionKind)
                    ? []
                    : [transition.ActionKind],
                LocatorEvidence =
                [
                    new KnowledgeLocatorEvidence
                    {
                        Kind = "candidate",
                        Status = "suggested",
                        Locator = locator
                    }
                ],
                FirstObservedAtUtc = transition.StartedAtUtc,
                LastObservedAtUtc = transition.CompletedAtUtc,
                ObservationCount = 1,
                Provenance = [explorationId],
                Confidence = string.IsNullOrWhiteSpace(identity.AutomationId) ? 0.4 : 0.8
            };
            return;
        }

        existing.ObservationCount++;
        existing.FirstObservedAtUtc = Min(existing.FirstObservedAtUtc, transition.StartedAtUtc);
        existing.LastObservedAtUtc = Max(existing.LastObservedAtUtc, transition.CompletedAtUtc);
        if (!string.IsNullOrWhiteSpace(transition.ActionKind)
            && !existing.Capabilities.Contains(transition.ActionKind, StringComparer.Ordinal))
        {
            existing.Capabilities.Add(transition.ActionKind);
            existing.Capabilities.Sort(StringComparer.Ordinal);
        }

        if (!existing.Provenance.Contains(explorationId, StringComparer.Ordinal))
            existing.Provenance.Add(explorationId);
        existing.Provenance.Sort(StringComparer.Ordinal);
    }

    private static void MergeTransition(
        Dictionary<string, KnowledgeTransitionEdge> transitions,
        KnowledgeStableIdFactory ids,
        string semanticTransitionId,
        string sourceStateId,
        string? targetStateId,
        string? elementId,
        string normalizedAction,
        ExplorationTransitionRecord transition,
        string outcome,
        string explorationId)
    {
        if (transitions.TryGetValue(semanticTransitionId, out var existing))
        {
            existing.ObservationCount++;
            existing.FirstObservedAtUtc = Min(existing.FirstObservedAtUtc, transition.StartedAtUtc);
            existing.LastObservedAtUtc = Max(existing.LastObservedAtUtc, transition.CompletedAtUtc);
            existing.Evidence.Add(new KnowledgeTransitionEvidence
            {
                ExplorationId = explorationId,
                TransitionId = transition.TransitionId,
                Detail = transition.Detail
            });
            existing.Evidence = existing.Evidence
                .OrderBy(e => e.ExplorationId, StringComparer.Ordinal)
                .ThenBy(e => e.TransitionId, StringComparer.Ordinal)
                .ToList();
            return;
        }

        // Conflicting edge may share source+action with a different outcome/target —
        // keep it as an explicit separate edge (conflict annotation happens later).
        _ = ids;
        transitions[semanticTransitionId] = new KnowledgeTransitionEdge
        {
            TransitionId = semanticTransitionId,
            Action = normalizedAction,
            ActionKind = transition.ActionKind,
            ActionCategory = transition.Category,
            ActionParams = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["driverOperation"] = transition.DriverOperation
            },
            ElementId = elementId,
            SourceStateId = sourceStateId,
            TargetStateId = targetStateId,
            Outcome = outcome,
            ObservationCount = 1,
            Evidence =
            [
                new KnowledgeTransitionEvidence
                {
                    ExplorationId = explorationId,
                    TransitionId = transition.TransitionId,
                    Detail = transition.Detail
                }
            ],
            FirstObservedAtUtc = transition.StartedAtUtc,
            LastObservedAtUtc = transition.CompletedAtUtc,
            Confidence = 1.0
        };
    }

    private static void AnnotateConflicts(
        Dictionary<string, KnowledgeTransitionEdge> transitions,
        KnowledgeStableIdFactory ids)
    {
        var groups = transitions.Values
            .GroupBy(
                t => t.SourceStateId + "\u001f" + KnowledgeStableIdFactory.NormalizeActionKey(t.Action),
                StringComparer.Ordinal);

        foreach (var group in groups)
        {
            var distinct = group
                .Select(t => (t.TargetStateId ?? string.Empty) + "\u001f" + t.Outcome)
                .Distinct(StringComparer.Ordinal)
                .ToList();
            if (distinct.Count <= 1)
                continue;

            var conflictId = ids.ConflictGroupId(group.First().SourceStateId, group.First().Action);
            foreach (var edge in group)
                edge.ConflictGroupId = conflictId;
        }
    }

    private static void MarkTerminalStates(
        Dictionary<string, KnowledgeStateNode> states,
        IEnumerable<KnowledgeTransitionEdge> transitions)
    {
        var sources = new HashSet<string>(
            transitions.Select(t => t.SourceStateId),
            StringComparer.Ordinal);
        var targets = new HashSet<string>(
            transitions.Where(t => !string.IsNullOrWhiteSpace(t.TargetStateId)).Select(t => t.TargetStateId!),
            StringComparer.Ordinal);

        foreach (var state in states.Values)
        {
            // Terminal: observed as a target (or entry) but never used as a successful navigation source.
            var hasOutgoingSuccess = transitions.Any(t =>
                string.Equals(t.SourceStateId, state.StateId, StringComparison.Ordinal)
                && t.Outcome is "success" or "newWindow" or "modalDialog");
            state.IsTerminal = !hasOutgoingSuccess && (targets.Contains(state.StateId) || state.IsEntry);
            _ = sources;
        }
    }

    private static List<ApplicationWorkflow> BuildWorkflows(
        ApplicationStateGraph graph,
        KnowledgeStableIdFactory ids,
        IReadOnlyList<LoadedExplorationRun> runs)
    {
        var workflows = new Dictionary<string, ApplicationWorkflow>(StringComparer.Ordinal);
        var segmentIndex = 0;

        foreach (var run in runs)
        {
            var successEdges = run.Transitions
                .Where(t => t.Outcome is ExplorationOutcome.Success
                    or ExplorationOutcome.NewWindow
                    or ExplorationOutcome.ModalDialog)
                .OrderBy(t => t.StartedAtUtc)
                .ThenBy(t => t.TransitionId, StringComparer.Ordinal)
                .ToList();

            if (successEdges.Count == 0)
                continue;

            // Split into continuous segments: next.source must equal previous.target.
            var segments = new List<List<KnowledgeWorkflowStep>>();
            List<KnowledgeWorkflowStep>? current = null;
            string? previousTarget = null;

            foreach (var edge in successEdges)
            {
                var source = graph.States.FirstOrDefault(s =>
                    string.Equals(s.SemanticFingerprint, edge.FromFingerprint, StringComparison.Ordinal));
                var target = edge.ToFingerprint is null
                    ? null
                    : graph.States.FirstOrDefault(s =>
                        string.Equals(s.SemanticFingerprint, edge.ToFingerprint, StringComparison.Ordinal));
                if (source is null)
                    continue;

                if (current is not null
                    && previousTarget is not null
                    && !string.Equals(previousTarget, source.StateId, StringComparison.Ordinal))
                {
                    segments.Add(current);
                    current = null;
                }

                current ??= [];
                var normalizedAction = KnowledgeStableIdFactory.NormalizeActionKey(edge.ActionKey);
                var outcome = CamelOutcome(edge.Outcome);
                var transitionId = ids.TransitionId(
                    source.StateId,
                    normalizedAction,
                    target?.StateId,
                    outcome);

                current.Add(new KnowledgeWorkflowStep
                {
                    Ordinal = current.Count,
                    TransitionId = transitionId,
                    Action = normalizedAction,
                    SourceStateId = source.StateId,
                    TargetStateId = target?.StateId,
                    Outcome = outcome
                });
                previousTarget = target?.StateId ?? source.StateId;
            }

            if (current is { Count: > 0 })
                segments.Add(current);

            foreach (var steps in segments)
            {
                var entry = steps[0].SourceStateId;
                var terminal = steps[^1].TargetStateId ?? steps[^1].SourceStateId;
                // Re-number ordinals contiguously after any prior split.
                for (var i = 0; i < steps.Count; i++)
                    steps[i].Ordinal = i;

                var workflowId = ids.WorkflowId(entry, steps.Select(s => s.TransitionId).ToList());
                if (workflows.TryGetValue(workflowId, out var existing))
                {
                    existing.ObservationCount++;
                    if (!existing.Provenance.Contains(run.ExplorationId, StringComparer.Ordinal))
                        existing.Provenance.Add(run.ExplorationId);
                    existing.Provenance.Sort(StringComparer.Ordinal);
                    continue;
                }

                segmentIndex++;
                workflows[workflowId] = new ApplicationWorkflow
                {
                    WorkflowSchemaVersion = KnowledgeConstants.WorkflowSchemaVersion,
                    WorkflowId = workflowId,
                    Name = $"run:{run.ExplorationId}:seg{segmentIndex}",
                    EntryStateId = entry,
                    TerminalStateId = terminal,
                    Steps = steps,
                    ObservationCount = 1,
                    Confidence = 1.0,
                    Provenance = [run.ExplorationId]
                };
            }
        }

        return workflows.Values
            .OrderBy(w => w.WorkflowId, StringComparer.Ordinal)
            .Take(KnowledgeConstants.MaxWorkflows)
            .ToList();
    }

    private static KnowledgeElementIdentity? ParseElementIdentity(ExplorationTransitionRecord transition)
    {
        var key = transition.ElementKey ?? string.Empty;
        var match = ElementKeyPattern.Match(key);
        if (match.Success)
        {
            var automationId = EmptyToNull(match.Groups["automationId"].Value);
            // Guard against HWND-ish automation ids leaking into stable identity.
            if (automationId is not null && automationId.Contains("hwnd:", StringComparison.OrdinalIgnoreCase))
                automationId = null;

            return new KnowledgeElementIdentity
            {
                ControlType = EmptyToNull(match.Groups["control"].Value) ?? transition.AutomationId ?? "Unknown",
                AutomationId = automationId ?? EmptyToNull(transition.AutomationId),
                Name = EmptyToNull(match.Groups["name"].Value) ?? EmptyToNull(transition.Name)
            };
        }

        if (!string.IsNullOrWhiteSpace(transition.AutomationId) || !string.IsNullOrWhiteSpace(transition.Name))
        {
            return new KnowledgeElementIdentity
            {
                ControlType = "Unknown",
                AutomationId = EmptyToNull(transition.AutomationId),
                Name = EmptyToNull(transition.Name)
            };
        }

        return null;
    }

    private static bool IsNameOnlyWeakIdentity(KnowledgeElementIdentity identity) =>
        string.IsNullOrWhiteSpace(identity.AutomationId)
        && string.IsNullOrWhiteSpace(identity.ClassName)
        && !string.IsNullOrWhiteSpace(identity.Name);

    private static bool IsBacktrack(ExplorationTransitionRecord transition) =>
        transition.ActionKey.StartsWith("backtrack:", StringComparison.Ordinal)
        || string.Equals(transition.Category, "backtrack", StringComparison.OrdinalIgnoreCase);

    private static string CamelOutcome(ExplorationOutcome outcome) =>
        outcome switch
        {
            ExplorationOutcome.Success => "success",
            ExplorationOutcome.NoOp => "noOp",
            ExplorationOutcome.Failure => "failure",
            ExplorationOutcome.ModalDialog => "modalDialog",
            ExplorationOutcome.NewWindow => "newWindow",
            ExplorationOutcome.ClosedWindow => "closedWindow",
            ExplorationOutcome.ScopeViolation => "scopeViolation",
            ExplorationOutcome.BudgetExhausted => "budgetExhausted",
            ExplorationOutcome.Cancelled => "cancelled",
            ExplorationOutcome.StabilizationTimeout => "stabilizationTimeout",
            ExplorationOutcome.StaleRef => "staleRef",
            ExplorationOutcome.DryRun => "dryRun",
            _ => outcome.ToString()
        };

    /// <summary>
    /// WinForms class names embed runtime app-domain suffixes; keep a stable prefix for identity.
    /// </summary>
    private static string? NormalizeVolatileClassName(string? className)
    {
        if (string.IsNullOrWhiteSpace(className))
            return null;

        var value = className.Trim();
        const string marker = ".app.";
        var idx = value.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (idx > 0)
            return value[..idx];
        return value;
    }

    private static void EnforceLimits(
        Dictionary<string, KnowledgeWindowNode> windows,
        Dictionary<string, KnowledgeStateNode> states,
        Dictionary<string, KnowledgeElementNode> elements,
        Dictionary<string, KnowledgeTransitionEdge> transitions,
        List<KnowledgeUnresolvedElement> unresolved,
        List<KnowledgeAmbiguousElement> ambiguous)
    {
        if (windows.Count > KnowledgeConstants.MaxWindows)
            throw new InvalidOperationException($"Knowledge window count exceeds {KnowledgeConstants.MaxWindows}.");
        if (states.Count > KnowledgeConstants.MaxStates)
            throw new InvalidOperationException($"Knowledge state count exceeds {KnowledgeConstants.MaxStates}.");
        if (elements.Count > KnowledgeConstants.MaxElements)
            throw new InvalidOperationException($"Knowledge element count exceeds {KnowledgeConstants.MaxElements}.");
        if (transitions.Count > KnowledgeConstants.MaxTransitions)
            throw new InvalidOperationException($"Knowledge transition count exceeds {KnowledgeConstants.MaxTransitions}.");
        if (unresolved.Count > 5000)
            unresolved.RemoveRange(5000, unresolved.Count - 5000);
        if (ambiguous.Count > 5000)
            ambiguous.RemoveRange(5000, ambiguous.Count - 5000);
    }

    private static string? EmptyToNull(string? value) =>
        string.IsNullOrWhiteSpace(value) || value == "-" ? null : value.Trim();

    private static DateTimeOffset? Min(DateTimeOffset? left, DateTimeOffset right) =>
        left is null || right < left ? right : left;

    private static DateTimeOffset? Max(DateTimeOffset? left, DateTimeOffset right) =>
        left is null || right > left ? right : left;
}
