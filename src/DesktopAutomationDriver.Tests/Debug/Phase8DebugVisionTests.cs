using System.Drawing;
using System.Drawing.Imaging;
using System.Reflection;
using System.Text;
using System.Text.Json;
using DesktopAutomationDriver.Services;
using DesktopAutomationDriver.Services.Debug;
using DesktopAutomationDriver.Services.Vision;

namespace DesktopAutomationDriver.Tests.Debug;

public class Phase8DebugVisionTests
{
    [Fact]
    public void ServiceSeparation_DebugTypesLiveOutsideOverlayAndUiService()
    {
        var overlay = typeof(RecordingOverlayWindow);
        var ui = typeof(UiService);
        Assert.Equal("DesktopAutomationDriver.Services", overlay.Namespace);
        Assert.Equal("DesktopAutomationDriver.Services", ui.Namespace);

        Assert.Equal("DesktopAutomationDriver.Services.Debug", typeof(DebugSessionCoordinator).Namespace);
        Assert.Equal("DesktopAutomationDriver.Services.Debug", typeof(UiaTreeCaptureService).Namespace);
        Assert.Equal("DesktopAutomationDriver.Services.Debug", typeof(LocatorCandidateGenerator).Namespace);
        Assert.Equal("DesktopAutomationDriver.Services.Debug", typeof(NonInteractiveHighlightService).Namespace);
        Assert.Equal("DesktopAutomationDriver.Services.Debug", typeof(CandidateVerificationService).Namespace);
        Assert.Equal("DesktopAutomationDriver.Services.Debug", typeof(DebugArtifactWriter).Namespace);
        Assert.Equal("DesktopAutomationDriver.Services.Vision", typeof(VisionFallbackCoordinator).Namespace);

        // Overlay may reference the coordinator type for thin Ctrl+D delegation only.
        var overlayMethods = overlay.GetMethods(BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
        Assert.DoesNotContain(overlayMethods, m => m.Name.Contains("GenerateCandidate", StringComparison.Ordinal));
        Assert.DoesNotContain(overlayMethods, m => m.Name.Contains("WriteCandidate", StringComparison.Ordinal));
        Assert.DoesNotContain(
            ui.GetMethods(BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public),
            m => m.Name.Contains("DebugSession", StringComparison.Ordinal));
    }

    [Fact]
    public void LocatorCandidateGenerator_PrefersUniqueAutomationId()
    {
        var generator = new LocatorCandidateGenerator();
        var target = new DebugUiaNode
        {
            AutomationId = "SaveButton",
            Name = "Save",
            ControlType = "Button",
            ClassName = "Button"
        };
        var siblings = new[]
        {
            new DebugUiaNode { AutomationId = "CancelButton", Name = "Cancel", ControlType = "Button" }
        };

        var candidates = generator.Generate(target, siblings);
        Assert.Single(candidates);
        Assert.Equal("strong", candidates[0].Grade);
        Assert.Equal("automationId+controlType", candidates[0].Strategy);
        Assert.False(candidates[0].AutoPromoted);
        Assert.Equal("SaveButton", candidates[0].Locator["automationId"]);
    }

    [Fact]
    public void LocatorCandidateGenerator_RejectsAmbiguousNameWithoutUniqueStrategy()
    {
        var generator = new LocatorCandidateGenerator();
        var target = new DebugUiaNode { Name = "Item", ControlType = "ListItem" };
        var siblings = new[]
        {
            new DebugUiaNode { Name = "Item", ControlType = "ListItem" },
            target
        };

        var candidates = generator.Generate(target, siblings);
        Assert.Empty(candidates);
    }

    [Fact]
    public void CandidateVerification_UniqueAndAmbiguous()
    {
        var svc = new CandidateVerificationService();
        var candidate = new LocatorCandidate
        {
            ElementId = "save",
            Locator = new Dictionary<string, object?> { ["automationId"] = "SaveButton" },
            Grade = "strong",
            AutoPromoted = false
        };

        var unique = svc.Verify(candidate, 1);
        Assert.True(unique.Unique);
        Assert.False(unique.Ambiguous);

        var ambiguous = svc.Verify(candidate, 3);
        Assert.False(ambiguous.Unique);
        Assert.True(ambiguous.Ambiguous);

        var missing = svc.Verify(candidate, 0);
        Assert.True(missing.NotFound);

        var promotedCandidate = new LocatorCandidate
        {
            ElementId = candidate.ElementId,
            Locator = candidate.Locator,
            Grade = candidate.Grade,
            AutoPromoted = true
        };
        var promoted = svc.Verify(promotedCandidate, 1);
        Assert.False(promoted.Unique);
        Assert.Contains("never promotes", promoted.Message!, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void HighlightLifecycle_ClearsAndDoesNotInterceptClicks()
    {
        var surface = new FakeHighlightSurface();
        using var highlight = new NonInteractiveHighlightService(surface);

        highlight.Highlight(new Rectangle(10, 20, 100, 40));
        Assert.True(highlight.IsVisible);
        Assert.False(highlight.InterceptsClicks);
        Assert.Equal(new Rectangle(10, 20, 100, 40), highlight.CurrentBounds);

        highlight.Clear();
        Assert.False(highlight.IsVisible);
        Assert.Null(highlight.CurrentBounds);
        Assert.Equal(1, surface.HideCount);

        highlight.Dispose();
        Assert.True(surface.Disposed);
    }

    [Fact]
    public void DebugSession_StartStopInspectWritesCandidateOnly_NoAutoPromotion()
    {
        var probe = new ScriptedUiaElementProbe();
        var surface = new FakeHighlightSurface();
        using var coordinator = DebugSessionCoordinator.ForTests(probe, surface);
        var dir = Path.Combine(Path.GetTempPath(), "phase8-debug-" + Guid.NewGuid().ToN());

        try
        {
            var started = coordinator.Start(dir);
            Assert.True(started.IsActive);

            var inspected = coordinator.InspectUnderCursor(new Point(50, 60), writeArtifact: true);
            Assert.NotNull(inspected.LastCapture);
            Assert.NotEmpty(inspected.LastCandidates);
            Assert.All(inspected.LastCandidates, c => Assert.False(c.AutoPromoted));
            Assert.True(inspected.LastVerification!.Unique);
            Assert.True(File.Exists(inspected.ArtifactPath));
            Assert.Contains("debug-candidates", inspected.ArtifactPath!, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain($"{Path.DirectorySeparatorChar}pages{Path.DirectorySeparatorChar}", inspected.ArtifactPath!);

            var json = File.ReadAllText(inspected.ArtifactPath!);
            Assert.Contains("\"autoPromoted\": false", json, StringComparison.Ordinal);
            Assert.Contains("\"requiresHumanApproval\": false", json, StringComparison.Ordinal);
            Assert.Contains("\"requiresOperatorConfirmationUx\": true", json, StringComparison.Ordinal);

            var stopped = coordinator.Stop();
            Assert.False(stopped.IsActive);
            Assert.False(surface.IsVisible);
        }
        finally
        {
            if (Directory.Exists(dir))
                Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void VisionPolicy_DisabledByDefault_AndDeniedWhenUiaAccessible()
    {
        var disabled = new VisionFallbackPolicy();
        Assert.False(disabled.IsEnabled);
        Assert.Equal(
            VisionEligibility.DisabledByDefault,
            disabled.Evaluate(new DebugUiaNode { ControlType = "Custom", IsSemanticAccessible = false }));

        var enabled = new VisionFallbackPolicy(new VisionFallbackOptions
        {
            Enabled = true,
            AllowWhenUiaInaccessible = true
        });
        Assert.Equal(
            VisionEligibility.DeniedUiaAccessible,
            enabled.Evaluate(new DebugUiaNode
            {
                AutomationId = "Ok",
                ControlType = "Button",
                IsSemanticAccessible = true
            }));
        Assert.Equal(
            VisionEligibility.AllowedUiaInaccessible,
            enabled.Evaluate(new DebugUiaNode
            {
                ControlType = "Custom",
                IsSemanticAccessible = false,
                InaccessibilityReason = "custom-drawn"
            }));
    }

    [Fact]
    public void ScreenshotScope_IsWindowOnly_AndSensitiveNotPersistedByDefault()
    {
        var options = new VisionFallbackOptions { PersistSensitiveScreenshots = false };
        var svc = new VisionScreenshotService(options);
        var window = new DebugWindowIdentity
        {
            WindowTitle = "Login — password entry",
            Bounds = new Rectangle(0, 0, 400, 300),
            DpiScaleX = 1,
            DpiScaleY = 1
        };
        var dir = Path.Combine(Path.GetTempPath(), "phase8-vision-" + Guid.NewGuid().ToN());
        Directory.CreateDirectory(dir);
        try
        {
            var shot = svc.CreateWindowScoped(
                Encoding.UTF8.GetBytes("fake-png"),
                window,
                ocrOrUiTextHint: "password field",
                persistDirectory: dir);
            Assert.Equal("window", shot.Scope);
            Assert.True(shot.ContainsSensitiveContent);
            Assert.False(shot.Persisted);
            Assert.Null(shot.PersistedPath);
            Assert.Empty(Directory.GetFiles(dir));

            Assert.Throws<InvalidOperationException>(() =>
                svc.CreateWindowScoped(
                    Encoding.UTF8.GetBytes("x"),
                    window,
                    allowUnrestrictedDesktop: true));
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void CoordinateGuard_RejectsStaleDpiLayoutBoundsOutside_AndDestructiveLabels()
    {
        var options = new VisionFallbackOptions { MaxScreenshotAgeMilliseconds = 500, MinimumConfidence = 0.5 };
        var guard = new CoordinateActionGuard(options);
        var window = new DebugWindowIdentity
        {
            WindowHandle = new IntPtr(42),
            ProcessId = 100,
            ProcessName = "app",
            Bounds = new Rectangle(0, 0, 800, 600),
            DpiScaleX = 1.0,
            DpiScaleY = 1.0
        };
        var candidate = new VisionCandidateTarget
        {
            Label = "OK",
            Bounds = new Rectangle(40, 40, 80, 24),
            Confidence = 0.9,
            Window = window,
            ProducedAtUtc = DateTimeOffset.UtcNow.AddSeconds(-5),
            AutoPromoted = false
        };

        var stale = guard.Evaluate(new CoordinateActionRequest
        {
            Candidate = candidate,
            CurrentWindow = window,
            NowUtc = DateTimeOffset.UtcNow,
            Action = "click"
        });
        Assert.False(stale.Allowed);
        Assert.Contains("Stale", stale.Reason, StringComparison.OrdinalIgnoreCase);

        var fresh = candidate with { ProducedAtUtc = DateTimeOffset.UtcNow };
        var dpiChanged = guard.Evaluate(new CoordinateActionRequest
        {
            Candidate = fresh,
            CurrentWindow = window with { DpiScaleX = 1.5, DpiScaleY = 1.5 },
            NowUtc = DateTimeOffset.UtcNow,
            Action = "click"
        });
        Assert.False(dpiChanged.Allowed);
        Assert.Contains("DPI", dpiChanged.Reason, StringComparison.OrdinalIgnoreCase);

        var layoutChanged = guard.Evaluate(new CoordinateActionRequest
        {
            Candidate = fresh,
            CurrentWindow = window with { Bounds = new Rectangle(0, 0, 1000, 700) },
            NowUtc = DateTimeOffset.UtcNow,
            Action = "click"
        });
        Assert.False(layoutChanged.Allowed);

        // Intersecting but not fully contained — center would be outside a tight window.
        var smallWindow = window with { Bounds = new Rectangle(0, 0, 100, 100) };
        var overflowing = fresh with
        {
            Bounds = new Rectangle(90, 90, 100, 100),
            Window = smallWindow
        };
        var outside = guard.Evaluate(new CoordinateActionRequest
        {
            Candidate = overflowing,
            CurrentWindow = smallWindow,
            NowUtc = DateTimeOffset.UtcNow,
            Action = "click"
        });
        Assert.False(outside.Allowed);
        Assert.Contains("inside", outside.Reason, StringComparison.OrdinalIgnoreCase);

        var missingHwnd = guard.Evaluate(new CoordinateActionRequest
        {
            Candidate = fresh with { Window = window with { WindowHandle = IntPtr.Zero } },
            CurrentWindow = window with { WindowHandle = IntPtr.Zero },
            NowUtc = DateTimeOffset.UtcNow,
            Action = "click"
        });
        Assert.False(missingHwnd.Allowed);

        var missingPid = guard.Evaluate(new CoordinateActionRequest
        {
            Candidate = fresh with { Window = window with { ProcessId = null } },
            CurrentWindow = window with { ProcessId = null },
            NowUtc = DateTimeOffset.UtcNow,
            Action = "click"
        });
        Assert.False(missingPid.Allowed);

        var destructiveFlag = guard.Evaluate(new CoordinateActionRequest
        {
            Candidate = fresh,
            CurrentWindow = window,
            NowUtc = DateTimeOffset.UtcNow,
            Action = "click",
            IsDestructive = true
        });
        Assert.False(destructiveFlag.Allowed);

        foreach (var label in new[] { "Delete", "Pay now", "Purchase", "Submit order", "Confirm", "Send", "Approve", "Publish", "Install" })
        {
            var denied = guard.Evaluate(new CoordinateActionRequest
            {
                Candidate = fresh with { Label = label },
                CurrentWindow = window,
                NowUtc = DateTimeOffset.UtcNow,
                Action = "click"
            });
            Assert.False(denied.Allowed);
            Assert.Contains("Destructive", denied.Reason, StringComparison.OrdinalIgnoreCase);
        }

        var unknownAction = guard.Evaluate(new CoordinateActionRequest
        {
            Candidate = fresh,
            CurrentWindow = window,
            NowUtc = DateTimeOffset.UtcNow,
            Action = "warp"
        });
        Assert.False(unknownAction.Allowed);

        var ok = guard.Evaluate(new CoordinateActionRequest
        {
            Candidate = fresh,
            CurrentWindow = window,
            NowUtc = DateTimeOffset.UtcNow,
            Action = "click"
        });
        Assert.True(ok.Allowed);
        Assert.NotNull(ok.ScreenPoint);
        Assert.True(window.Bounds.Contains(ok.ScreenPoint!.Value));
    }

    [Fact]
    public void VisionCoordinator_DoesNotProduceCandidatesWhenDisabled()
    {
        var coordinator = VisionFallbackCoordinator.ForTests();
        var window = new DebugWindowIdentity { Bounds = new Rectangle(0, 0, 200, 100) };
        var inaccessible = new DebugUiaNode
        {
            ControlType = "Custom",
            IsSemanticAccessible = false,
            InaccessibilityReason = "custom-drawn"
        };

        var (eligibility, shot, candidates) = coordinator.TryProduceCandidates(
            inaccessible,
            window,
            Encoding.UTF8.GetBytes("png"),
            query: "Custom Drawn Target");

        Assert.Equal(VisionEligibility.DisabledByDefault, eligibility);
        Assert.Null(shot);
        Assert.Empty(candidates);
    }

    [Fact]
    public void VisionCoordinator_ProducesCandidateForCustomDrawnWhenEnabled()
    {
        var options = new VisionFallbackOptions
        {
            Enabled = true,
            AllowWhenUiaInaccessible = true,
            Provider = "fake",
            MinimumConfidence = 0.5
        };
        var coordinator = VisionFallbackCoordinator.ForTests(options);
        var window = new DebugWindowIdentity
        {
            Bounds = new Rectangle(10, 10, 400, 200),
            DpiScaleX = 1,
            DpiScaleY = 1,
            WindowHandle = new IntPtr(7)
        };
        var inaccessible = new DebugUiaNode
        {
            ControlType = "Custom",
            Name = null,
            AutomationId = null,
            IsSemanticAccessible = false,
            InaccessibilityReason = "custom-drawn"
        };

        var (eligibility, shot, candidates) = coordinator.TryProduceCandidates(
            inaccessible,
            window,
            Encoding.UTF8.GetBytes("png-bytes"),
            query: "Custom Drawn Target");

        Assert.Equal(VisionEligibility.AllowedUiaInaccessible, eligibility);
        Assert.NotNull(shot);
        Assert.Equal("window", shot!.Scope);
        Assert.NotEmpty(candidates);
        Assert.All(candidates, c => Assert.False(c.AutoPromoted));
        Assert.True(candidates[0].Confidence >= options.MinimumConfidence);
        Assert.True(candidates[0].Bounds.Width > 0);
    }

    [Fact]
    public void DoubleClickPolicy_RequiresPolicyOrKnownBehavior()
    {
        Assert.False(DoubleClickPolicy.IsAllowed("ListItem", policyAllowsDoubleClick: false));
        Assert.False(DoubleClickPolicy.IsAllowed("Button", policyAllowsDoubleClick: false, knownBehavior: false));
        Assert.True(DoubleClickPolicy.IsAllowed("ListItem", policyAllowsDoubleClick: true));
        Assert.True(DoubleClickPolicy.IsAllowed("DataItem", policyAllowsDoubleClick: false, knownBehavior: true));
    }

    [Fact]
    public void UiaTreeCapture_DetectsCustomDrawnInaccessible()
    {
        Assert.True(UiaTreeCaptureService.IsSemanticallyInaccessible(new DebugUiaNode
        {
            ControlType = "Custom"
        }));
        // WinForms UserPaint often reports Pane with empty AutomationId+Name.
        Assert.True(UiaTreeCaptureService.IsSemanticallyInaccessible(new DebugUiaNode
        {
            ControlType = "Pane",
            AutomationId = null,
            Name = null,
            IsSemanticAccessible = true
        }));
        // Named canvas would break the inaccessibility rule (regression guard).
        Assert.False(UiaTreeCaptureService.IsSemanticallyInaccessible(new DebugUiaNode
        {
            ControlType = "Pane",
            AutomationId = "CustomDrawnCanvas",
            Name = "CustomDrawnCanvas",
            IsSemanticAccessible = true
        }));
        Assert.False(UiaTreeCaptureService.IsSemanticallyInaccessible(new DebugUiaNode
        {
            AutomationId = "Ok",
            ControlType = "Button",
            IsSemanticAccessible = true
        }));
    }

    [Fact]
    public void CustomDrawnCanvas_HasEmptyWinFormsName_ForInaccessibilityRule()
    {
        // Fixture check: Name must stay empty so UIA AutomationId stays empty.
        // Strip comments first — the source intentionally documents "Do NOT set Name = UiIds.CustomDrawnCanvas".
        var source = File.ReadAllText(Path.Combine(
            FindRepoRoot(),
            "src", "DesktopAutomationTestApp", "CustomDrawnCanvas.cs"));
        var withoutBlockComments = System.Text.RegularExpressions.Regex.Replace(
            source, @"/\*.*?\*/", " ", System.Text.RegularExpressions.RegexOptions.Singleline);
        var codeOnly = string.Join(
            '\n',
            withoutBlockComments
                .Split('\n')
                .Select(line =>
                {
                    var idx = line.IndexOf("//", StringComparison.Ordinal);
                    return idx < 0 ? line : line[..idx];
                }));

        Assert.Contains("Name = string.Empty", codeOnly, StringComparison.Ordinal);
        Assert.DoesNotContain("Name = UiIds.CustomDrawnCanvas", codeOnly, StringComparison.Ordinal);
        // Tag may still reference the id as an app-only marker (not UIA AutomationId).
        Assert.Contains("Tag = UiIds.CustomDrawnCanvas", codeOnly, StringComparison.Ordinal);
    }

    [Fact]
    public void DebugSession_ProbeUnavailable_DoesNotBecomeUniqueViaStaleCapture()
    {
        var probe = new ScriptedUiaElementProbe { MatchResult = LocatorMatchProbeResult.Unavailable("boom") };
        using var coordinator = DebugSessionCoordinator.ForTests(probe);
        coordinator.Start();
        var inspected = coordinator.InspectUnderCursor(new Point(50, 60));
        Assert.NotNull(inspected.LastVerification);
        Assert.True(inspected.LastVerification!.ProbeUnavailable);
        Assert.False(inspected.LastVerification.Unique);

        var verified = coordinator.VerifyLastCandidate();
        Assert.True(verified.LastVerification!.ProbeUnavailable);
        Assert.False(verified.LastVerification.Unique);
    }

    [Fact]
    public void DebugSession_LiveZeroMatch_IsNotFound_NotUnique()
    {
        var probe = new ScriptedUiaElementProbe { MatchResult = LocatorMatchProbeResult.Ok(0) };
        using var coordinator = DebugSessionCoordinator.ForTests(probe);
        coordinator.Start();
        var inspected = coordinator.InspectUnderCursor(new Point(50, 60));
        Assert.True(inspected.LastVerification!.NotFound);
        Assert.False(inspected.LastVerification.Unique);
        Assert.False(inspected.LastVerification.ProbeUnavailable);
    }

    [Fact]
    public void DebugSession_VisionFallback_ProducesCandidatesWhenEnabledAndInaccessible()
    {
        var options = new VisionFallbackOptions
        {
            Enabled = true,
            AllowWhenUiaInaccessible = true,
            Provider = "fake",
            MinimumConfidence = 0.5
        };
        var vision = VisionFallbackCoordinator.ForTests(options);
        var probe = new InaccessibleUiaElementProbe();
        using var coordinator = DebugSessionCoordinator.ForTests(
            probe,
            vision: vision,
            screenshots: new VisionScreenshotService(options));

        coordinator.Start();
        var inspected = coordinator.InspectUnderCursor(new Point(120, 140));
        Assert.Equal(nameof(VisionEligibility.AllowedUiaInaccessible), inspected.VisionEligibility);
        Assert.NotEmpty(inspected.VisionCandidates);
        Assert.All(inspected.VisionCandidates, c => Assert.False(c.AutoPromoted));
    }

    [Fact]
    public void FlaUIProbe_ScaleFromWindowDpi_UsesRealDpiNotHardcodedOne()
    {
        var (x96, y96) = FlaUIUiaElementProbe.ScaleFromWindowDpi(96);
        Assert.Equal(1.0, x96);
        Assert.Equal(1.0, y96);

        var (x144, y144) = FlaUIUiaElementProbe.ScaleFromWindowDpi(144);
        Assert.Equal(1.5, x144);
        Assert.Equal(1.5, y144);

        var (x120, _) = FlaUIUiaElementProbe.ScaleFromWindowDpi(120);
        Assert.Equal(1.25, x120);
    }

    [Fact]
    public void VisionOptions_DefaultDisabled_AndDestructiveHeuristicsCoverLabels()
    {
        Assert.False(new VisionFallbackOptions().Enabled);
        Assert.True(DestructiveActionHeuristics.IsDestructiveLabel("Delete item"));
        Assert.True(DestructiveActionHeuristics.IsDestructiveLabel("Pay"));
        Assert.True(DestructiveActionHeuristics.IsDestructive("click", "Purchase", false));
        Assert.False(DestructiveActionHeuristics.IsDestructiveLabel("Save"));

        // P2 fail-closed: Send/Approve/Publish/Install labels + unknown actions denied.
        Assert.True(DestructiveActionHeuristics.IsDestructiveLabel("Send"));
        Assert.True(DestructiveActionHeuristics.IsDestructiveLabel("Approve request"));
        Assert.True(DestructiveActionHeuristics.IsDestructiveLabel("Publish"));
        Assert.True(DestructiveActionHeuristics.IsDestructiveLabel("Install now"));
        Assert.True(DestructiveActionHeuristics.IsUnknownOrDisallowedAction("explode"));
        Assert.True(DestructiveActionHeuristics.IsUnknownOrDisallowedAction("delete"));
        Assert.True(DestructiveActionHeuristics.IsUnknownOrDisallowedAction(null));
        Assert.False(DestructiveActionHeuristics.IsUnknownOrDisallowedAction("click"));
        Assert.False(DestructiveActionHeuristics.IsUnknownOrDisallowedAction("hover"));
        Assert.True(DestructiveActionHeuristics.IsDestructive("send", "OK", false));
    }

    [Fact]
    public void Authorize_DeniesWhenVisionDisabled_EvenWithFreshCandidate()
    {
        var window = new DebugWindowIdentity
        {
            WindowHandle = new IntPtr(42),
            ProcessId = 100,
            Bounds = new Rectangle(0, 0, 800, 600),
            DpiScaleX = 1,
            DpiScaleY = 1
        };
        var options = new VisionFallbackOptions { Enabled = false, Provider = "fake", MinimumConfidence = 0.5 };
        var store = new VisionCandidateStore(options);
        var probe = new ScriptedLiveWindowProbe(window);
        var coordinator = VisionFallbackCoordinator.ForTests(options, store, probe);

        var issued = store.Issue(new VisionCandidateTarget
        {
            Label = "OK",
            Bounds = new Rectangle(40, 40, 80, 24),
            Confidence = 0.9,
            Window = window,
            ProducedAtUtc = DateTimeOffset.UtcNow
        }, DateTimeOffset.UtcNow);
        store.TryBindApproval(issued.Id, "click", "vtx-test", DateTimeOffset.UtcNow, out _);

        var decision = coordinator.AuthorizeCoordinateAction("vtx-test", "click");
        Assert.False(decision.Allowed);
        Assert.Contains("Vision:Enabled=false", decision.Reason, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Authorize_RequiresOperatorApprovalToken_AndConsumesOnSuccess()
    {
        var window = new DebugWindowIdentity
        {
            WindowHandle = new IntPtr(99),
            ProcessId = 200,
            Bounds = new Rectangle(0, 0, 400, 300),
            DpiScaleX = 1,
            DpiScaleY = 1
        };
        var options = new VisionFallbackOptions
        {
            Enabled = true,
            AllowWhenUiaInaccessible = true,
            Provider = "fake",
            MinimumConfidence = 0.5,
            MaxScreenshotAgeMilliseconds = 5000,
            CandidateTtlMilliseconds = 60_000,
            ExecutionTokenTtlMilliseconds = 30_000
        };
        var store = new VisionCandidateStore(options);
        var probe = new ScriptedLiveWindowProbe(window);
        var capture = new ScriptedVisionCaptureService(new VisionServerCapture
        {
            Window = window,
            SemanticTarget = new DebugUiaNode
            {
                ControlType = "Custom",
                IsSemanticAccessible = false,
                InaccessibilityReason = "custom-drawn"
            },
            PngBytes = Encoding.UTF8.GetBytes("server-png")
        });
        var coordinator = VisionFallbackCoordinator.ForTests(options, store, probe, capture);

        var (_, _, candidates) = coordinator.TryProduceCandidatesFromHandle(
            window.WindowHandle,
            cursorScreenPoint: new Point(50, 60),
            query: "OK");
        Assert.NotEmpty(candidates);
        Assert.False(string.IsNullOrWhiteSpace(candidates[0].Id));

        var missingToken = coordinator.AuthorizeCoordinateAction("vtx-missing", "click");
        Assert.False(missingToken.Allowed);

        var approval = coordinator.ApproveCandidate(candidates[0].Id!, "click");
        Assert.True(approval.Allowed);
        Assert.False(string.IsNullOrWhiteSpace(approval.ExecutionToken));

        var wrongAction = coordinator.AuthorizeCoordinateAction(approval.ExecutionToken!, "hover");
        Assert.False(wrongAction.Allowed);

        var ok = coordinator.AuthorizeCoordinateAction(approval.ExecutionToken!, "click");
        Assert.True(ok.Allowed);
        Assert.NotNull(ok.ScreenPoint);

        var replay = coordinator.AuthorizeCoordinateAction(approval.ExecutionToken!, "click");
        Assert.False(replay.Allowed);
    }

    [Fact]
    public void Destructive_AfterApprove_StillDenied_WithVisionEnabled()
    {
        var window = new DebugWindowIdentity
        {
            WindowHandle = new IntPtr(77),
            ProcessId = 300,
            Bounds = new Rectangle(0, 0, 400, 300),
            DpiScaleX = 1,
            DpiScaleY = 1
        };
        var options = new VisionFallbackOptions
        {
            Enabled = true,
            AllowWhenUiaInaccessible = true,
            Provider = "fake",
            MinimumConfidence = 0.5,
            MaxScreenshotAgeMilliseconds = 5000
        };
        var store = new VisionCandidateStore(options);
        var probe = new ScriptedLiveWindowProbe(window);
        var coordinator = VisionFallbackCoordinator.ForTests(options, store, probe);

        var issued = store.Issue(new VisionCandidateTarget
        {
            Label = "Delete",
            Bounds = new Rectangle(40, 40, 80, 24),
            Confidence = 0.9,
            Window = window,
            ProducedAtUtc = DateTimeOffset.UtcNow
        }, DateTimeOffset.UtcNow);

        var approval = coordinator.ApproveCandidate(issued.Id, "click");
        Assert.False(approval.Allowed);
        Assert.DoesNotContain("Vision:Enabled=false", approval.Reason, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Destructive", approval.Reason, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void CandidateStore_EvictsByCapacity_AndExpires()
    {
        var options = new VisionFallbackOptions { MaxStoredCandidates = 2, CandidateTtlMilliseconds = 50 };
        var store = new VisionCandidateStore(options);
        var window = new DebugWindowIdentity
        {
            WindowHandle = new IntPtr(1),
            ProcessId = 1,
            Bounds = new Rectangle(0, 0, 10, 10),
            DpiScaleX = 1,
            DpiScaleY = 1
        };

        VisionCandidateTarget Make(string label) => new()
        {
            Label = label,
            Bounds = new Rectangle(1, 1, 2, 2),
            Confidence = 0.9,
            Window = window
        };

        var a = store.Issue(Make("a"), DateTimeOffset.UtcNow);
        var b = store.Issue(Make("b"), DateTimeOffset.UtcNow);
        var c = store.Issue(Make("c"), DateTimeOffset.UtcNow);
        Assert.True(store.Count <= 2);
        Assert.False(store.TryGet(a.Id, out _));
        Assert.True(store.TryGet(c.Id, out _));

        Thread.Sleep(60);
        store.EvictExpired(DateTimeOffset.UtcNow);
        Assert.Equal(0, store.Count);
        _ = b;
    }

    [Fact]
    public void ReApproval_InvalidatesOldExecutionToken_NoOrphanGrant()
    {
        var window = MakeWindow(new IntPtr(55), 550);
        var options = EnabledVisionOptions();
        var store = new VisionCandidateStore(options);
        var probe = new ScriptedLiveWindowProbe(window);
        var png = Encoding.UTF8.GetBytes("stable-png");
        var capture = new ScriptedVisionCaptureService(new VisionServerCapture
        {
            Window = window,
            SemanticTarget = InaccessibleTarget(),
            PngBytes = png
        });
        var coordinator = VisionFallbackCoordinator.ForTests(options, store, probe, capture);

        var (_, _, candidates) = coordinator.TryProduceCandidatesFromHandle(
            window.WindowHandle, new Point(40, 50), query: "OK");
        Assert.NotEmpty(candidates);

        var first = coordinator.ApproveCandidate(candidates[0].Id!, "click");
        Assert.True(first.Allowed);
        Assert.Equal(1, store.TokenGrantCount);

        var second = coordinator.ApproveCandidate(candidates[0].Id!, "click");
        Assert.True(second.Allowed);
        Assert.NotEqual(first.ExecutionToken, second.ExecutionToken);
        Assert.Equal(1, store.TokenGrantCount);

        var oldDenied = coordinator.AuthorizeCoordinateAction(first.ExecutionToken!, "click");
        Assert.False(oldDenied.Allowed);

        var newOk = coordinator.AuthorizeCoordinateAction(second.ExecutionToken!, "click");
        Assert.True(newOk.Allowed);
        Assert.Equal(0, store.TokenGrantCount);
    }

    [Fact]
    public void ConcurrentAuthorize_OnlyOneConsumesToken()
    {
        var window = MakeWindow(new IntPtr(66), 660);
        var options = EnabledVisionOptions();
        var store = new VisionCandidateStore(options);
        var probe = new ScriptedLiveWindowProbe(window);
        var capture = new ScriptedVisionCaptureService(new VisionServerCapture
        {
            Window = window,
            SemanticTarget = InaccessibleTarget(),
            PngBytes = Encoding.UTF8.GetBytes("concurrent-png")
        });
        var coordinator = VisionFallbackCoordinator.ForTests(options, store, probe, capture);

        var (_, _, candidates) = coordinator.TryProduceCandidatesFromHandle(
            window.WindowHandle, new Point(40, 50), query: "OK");
        var approval = coordinator.ApproveCandidate(candidates[0].Id!, "click");
        Assert.True(approval.Allowed);

        var barrier = new Barrier(2);
        var results = new CoordinateActionDecision[2];
        void Worker(int index)
        {
            barrier.SignalAndWait();
            results[index] = coordinator.AuthorizeCoordinateAction(approval.ExecutionToken!, "click");
        }

        var t0 = new Thread(() => Worker(0));
        var t1 = new Thread(() => Worker(1));
        t0.Start();
        t1.Start();
        t0.Join();
        t1.Join();

        Assert.Equal(1, results.Count(r => r.Allowed));
        Assert.Equal(1, results.Count(r => !r.Allowed));
        Assert.Equal(0, store.TokenGrantCount);
    }

    [Fact]
    public void Approve_DeniesWhenContentFingerprintChanges()
    {
        var window = MakeWindow(new IntPtr(77), 770);
        var options = EnabledVisionOptions();
        var store = new VisionCandidateStore(options);
        var probe = new ScriptedLiveWindowProbe(window);
        var holder = new MutableCapture { Window = window, PngBytes = Encoding.UTF8.GetBytes("content-v1") };
        var capture = new ScriptedVisionCaptureService((hwnd, _) =>
            hwnd == holder.Window.WindowHandle
                ? new VisionServerCapture
                {
                    Window = holder.Window,
                    SemanticTarget = InaccessibleTarget(),
                    PngBytes = holder.PngBytes
                }
                : null);
        var coordinator = VisionFallbackCoordinator.ForTests(options, store, probe, capture);

        var (_, _, candidates) = coordinator.TryProduceCandidatesFromHandle(
            window.WindowHandle, new Point(40, 50), query: "OK");
        Assert.NotEmpty(candidates);

        holder.PngBytes = Encoding.UTF8.GetBytes("content-v2-changed");
        var approval = coordinator.ApproveCandidate(candidates[0].Id!, "click");
        Assert.False(approval.Allowed);
        Assert.Contains("fingerprint", approval.Reason, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(0, store.TokenGrantCount);
    }

    [Fact]
    public void ServerCapture_RejectsCursorOnOtherWindow_AndOutsideBounds()
    {
        var requested = MakeWindow(new IntPtr(0x100), 1000, new Rectangle(0, 0, 400, 300));
        var other = MakeWindow(new IntPtr(0x200), 2000, new Rectangle(500, 500, 400, 300));
        var windowProbe = new ScriptedLiveWindowProbe(requested);
        var pixels = new ScriptedWindowPixelCapture(MakeSolidPng(Color.Silver, 16, 16));
        var uia = new ScriptedWindowBoundUiaProbe(other);
        var capture = new ServerVisionCaptureService(windowProbe, uia, pixels);

        Assert.Null(capture.Capture(requested.WindowHandle, new Point(50, 60)));
        Assert.Null(capture.Capture(requested.WindowHandle, new Point(900, 900)));

        var matchingUia = new ScriptedWindowBoundUiaProbe(requested);
        var okCapture = new ServerVisionCaptureService(windowProbe, matchingUia, pixels);
        Assert.NotNull(okCapture.Capture(requested.WindowHandle, new Point(50, 60)));
    }

    [Fact]
    public void StartupBanner_NeverIncludesOperatorToken()
    {
        var ctx = new DriverContext("banner-test-user");
        var secureBanner = StartupBanner.Format(ctx); // Secure mode is the default
        Assert.False(string.IsNullOrWhiteSpace(ctx.OperatorToken));
        Assert.Contains("Bearer token", secureBanner, StringComparison.Ordinal);
        Assert.DoesNotContain(ctx.BearerToken, secureBanner, StringComparison.Ordinal);
        Assert.DoesNotContain(ctx.OperatorToken, secureBanner, StringComparison.Ordinal);
        Assert.Contains("confirmation UX", secureBanner, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("UIA-automatable", secureBanner, StringComparison.OrdinalIgnoreCase);
        var bannerFlat = string.Join(' ',
            secureBanner.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
        Assert.Contains("not OS-enforced human presence", bannerFlat, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("ACL secret path", secureBanner, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("X-Operator-Token:", secureBanner, StringComparison.Ordinal);

        var legacyBanner = StartupBanner.Format(ctx, new DesktopAutomationDriver.Services.Bootstrap.BootstrapOptions
        {
            Mode = DesktopAutomationDriver.Services.Bootstrap.BootstrapDiscoveryMode.LegacyVerify,
            AcknowledgeLegacyVerifyRisk = true
        });
        Assert.Contains(ctx.BearerToken, legacyBanner, StringComparison.Ordinal);
        Assert.DoesNotContain(ctx.OperatorToken, legacyBanner, StringComparison.Ordinal);
    }

    [Fact]
    public void OperatorTokenAclSecret_IsDisabled_NotApprovalChannel()
    {
#pragma warning disable CS0618
        Assert.Null(OperatorTokenAclSecret.TryRead());
        Assert.Throws<InvalidOperationException>(() =>
            OperatorTokenAclSecret.WriteCurrentUserOnly("should-not-write"));
#pragma warning restore CS0618
    }

    [Fact]
    public void Approve_RequiresOperatorConfirmationUx_AclSelfReadIsNotEnough()
    {
        var window = MakeWindow(new IntPtr(0xA1), 101);
        var options = EnabledVisionOptions();
        var store = new VisionCandidateStore(options);
        var probe = new ScriptedLiveWindowProbe(window);
        var png = MakeSolidPng(Color.SteelBlue, 40, 30);
        var capture = new ScriptedVisionCaptureService(new VisionServerCapture
        {
            Window = window,
            SemanticTarget = InaccessibleTarget(),
            PngBytes = png
        });
        var presence = new ScriptedOperatorPresenceService { Allow = false };
        var coordinator = VisionFallbackCoordinator.ForTests(options, store, probe, capture, presence);

        var (_, _, candidates) = coordinator.TryProduceCandidatesFromHandle(
            window.WindowHandle, new Point(40, 50), query: "OK");
        Assert.NotEmpty(candidates);

#pragma warning disable CS0618
        Assert.Null(OperatorTokenAclSecret.TryRead());
#pragma warning restore CS0618
        var denied = coordinator.ApproveCandidate(candidates[0].Id!, "click");
        Assert.False(denied.Allowed);
        Assert.Contains("confirmation", denied.Reason, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(1, presence.ConfirmCalls);
    }

    [Fact]
    public void OperatorConfirmation_SingleFlight_DeniesSecondCaller()
    {
        var request = new OperatorPresenceRequest("vc-1", "click", "OK");
        var holding = new HoldingPresenceService(TimeSpan.FromMilliseconds(500));
        var gate = new ManualResetEventSlim(false);
        var firstStarted = new ManualResetEventSlim(false);
        var firstOk = false;

        var a = new Thread(() =>
        {
            firstStarted.Set();
            firstOk = holding.TryConfirmVisionApproval(request, out _);
            gate.Set();
        });
        a.Start();
        Assert.True(firstStarted.Wait(1000));
        Thread.Sleep(30);
        var concurrentOk = holding.TryConfirmVisionApproval(request, out var concurrentReason);
        Assert.False(concurrentOk);
        Assert.Contains("single-flight", concurrentReason, StringComparison.OrdinalIgnoreCase);
        Assert.True(gate.Wait(2000));
        Assert.True(firstOk);
    }

    [Fact]
    public void OperatorConfirmation_Cancel_Denies()
    {
        var svc = new ScriptedOperatorPresenceService
        {
            Allow = true,
            ArtificialDelay = TimeSpan.FromSeconds(5)
        };
        using var cts = new CancellationTokenSource(50);
        var ok = svc.TryConfirmVisionApproval(
            new OperatorPresenceRequest("vc-cancel", "click", "OK"),
            out var reason,
            cts.Token);
        Assert.False(ok);
        Assert.Contains("cancelled", reason, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void InteractiveOperatorConfirmation_Cancel_DeniesAndReleasesSingleFlight()
    {
        // Exercises production InteractiveOperatorPresenceService cancel path (not scripted).
        var svc = new InteractiveOperatorPresenceService(TimeSpan.FromSeconds(30));
        using var cts = new CancellationTokenSource();
        var started = new ManualResetEventSlim(false);
        var done = new ManualResetEventSlim(false);
        var ok = true;
        string? reason = null;

        var worker = new Thread(() =>
        {
            started.Set();
            ok = svc.TryConfirmVisionApproval(
                new OperatorPresenceRequest("vc-interactive-cancel", "click", "OK"),
                out reason,
                cts.Token);
            done.Set();
        })
        {
            IsBackground = true,
            Name = "InteractiveConfirmCancelTest"
        };
        worker.Start();

        Assert.True(started.Wait(TimeSpan.FromSeconds(2)));
        // Poll until single-flight is acquired — fixed sleep flakes on slow CI before dialog starts.
        var inFlightDeadline = DateTime.UtcNow + TimeSpan.FromSeconds(5);
        while (!svc.IsInFlight && DateTime.UtcNow < inFlightDeadline)
            Thread.Sleep(20);
        Assert.True(svc.IsInFlight, "Confirmation never acquired single-flight before cancel window elapsed.");
        cts.Cancel();

        Assert.True(done.Wait(TimeSpan.FromSeconds(15)), "Confirmation did not return after cancel.");
        Assert.False(ok);
        Assert.Contains("cancelled", reason ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        Assert.True(
            WaitUntil(() => !svc.IsInFlight, TimeSpan.FromSeconds(15)),
            "Single-flight was not released after cancel cleanup.");

        // Single-flight must be released — a follow-up call must not hit single-flight deny.
        using var cts2 = new CancellationTokenSource(TimeSpan.FromMilliseconds(400));
        var second = svc.TryConfirmVisionApproval(
            new OperatorPresenceRequest("vc-interactive-cancel-2", "click", "OK"),
            out var reason2,
            cts2.Token);
        Assert.False(second);
        Assert.DoesNotContain("single-flight", reason2, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void InteractiveOperatorConfirmation_CancelBeforeFormReady_DoesNotShowDialog()
    {
        // Deterministic race: delay form creation past cancel so ready.Set has not fired.
        // Production must not ShowDialog after cancel, must end the UI thread, and must
        // keep single-flight correct (held until thread ends, then released).
        var svc = new InteractiveOperatorPresenceService(TimeSpan.FromSeconds(30))
        {
            // Longer than CleanupJoinMilliseconds * 2 so ReleaseSingleFlight may hand off.
            FormCreationDelay = TimeSpan.FromMilliseconds(
                InteractiveOperatorPresenceService.CleanupJoinMilliseconds * 2 + 1_500)
        };
        using var cts = new CancellationTokenSource();
        var started = new ManualResetEventSlim(false);
        var done = new ManualResetEventSlim(false);
        var ok = true;
        string? reason = null;
        var httpElapsed = TimeSpan.Zero;

        var worker = new Thread(() =>
        {
            started.Set();
            var sw = System.Diagnostics.Stopwatch.StartNew();
            ok = svc.TryConfirmVisionApproval(
                new OperatorPresenceRequest("vc-cancel-before-ready", "click", "OK"),
                out reason,
                cts.Token);
            sw.Stop();
            httpElapsed = sw.Elapsed;
            done.Set();
        })
        {
            IsBackground = true,
            Name = "InteractiveConfirmCancelBeforeReadyTest"
        };
        worker.Start();

        Assert.True(started.Wait(TimeSpan.FromSeconds(2)));
        var inFlightDeadline = DateTime.UtcNow + TimeSpan.FromSeconds(5);
        while (!svc.IsInFlight && DateTime.UtcNow < inFlightDeadline)
            Thread.Sleep(20);
        Assert.True(svc.IsInFlight, "Confirmation never acquired single-flight before cancel.");
        cts.Cancel();

        Assert.True(done.Wait(TimeSpan.FromSeconds(20)), "Confirmation did not return after cancel.");
        Assert.False(ok);
        Assert.Contains("cancelled", reason ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        // HTTP must not wait for the full FormCreationDelay / unbounded Join.
        Assert.True(
            httpElapsed < svc.FormCreationDelay,
            $"HTTP blocked for full form-creation delay: {httpElapsed} >= {svc.FormCreationDelay}");
        Assert.Equal(0, svc.ShowDialogEntryCount);

        Assert.True(
            WaitUntil(() => !svc.IsInFlight, TimeSpan.FromSeconds(20)),
            "Single-flight was not released after delayed UI thread ended.");
        var ui = svc.LastUiThread;
        Assert.True(ui is null || !ui.IsAlive, "Stale UI thread still alive after cancel cleanup.");

        using var cts2 = new CancellationTokenSource(TimeSpan.FromMilliseconds(400));
        var second = svc.TryConfirmVisionApproval(
            new OperatorPresenceRequest("vc-cancel-before-ready-2", "click", "OK"),
            out var reason2,
            cts2.Token);
        Assert.False(second);
        Assert.DoesNotContain("single-flight", reason2, StringComparison.OrdinalIgnoreCase);
    }

    private static bool WaitUntil(Func<bool> condition, TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < deadline)
        {
            if (condition())
                return true;
            Thread.Sleep(25);
        }

        return condition();
    }

    [Fact]
    public void OperatorConfirmation_Timeout_Denies_WithoutUnlimitedJoin()
    {
        // Short timeout; form never receives Yes — must deny without unbounded Join.
        // On CI, WinForms may surface timeout, cancel-after-close, or UI-failed; all are denies.
        var svc = new InteractiveOperatorPresenceService(TimeSpan.FromMilliseconds(600));
        var sw = System.Diagnostics.Stopwatch.StartNew();
        var ok = svc.TryConfirmVisionApproval(
            new OperatorPresenceRequest("vc-timeout", "click", "OK"),
            out var reason);
        sw.Stop();
        Assert.False(ok);
        Assert.False(string.IsNullOrWhiteSpace(reason));
        Assert.True(
            reason.Contains("timed out", StringComparison.OrdinalIgnoreCase)
            || reason.Contains("cancelled", StringComparison.OrdinalIgnoreCase)
            || reason.Contains("denied", StringComparison.OrdinalIgnoreCase)
            || reason.Contains("UI failed", StringComparison.OrdinalIgnoreCase),
            $"Expected a confirmation deny reason, got: {reason}");
        Assert.True(sw.Elapsed < TimeSpan.FromSeconds(5), $"Join appeared unbounded: {sw.Elapsed}");
    }

    /// <summary>Scripted single-flight for unit tests (mirrors Interactive lock semantics).</summary>
    private sealed class HoldingPresenceService : IOperatorPresenceService
    {
        private readonly TimeSpan _hold;
        private int _inFlight;

        public HoldingPresenceService(TimeSpan hold) => _hold = hold;

        public bool TryConfirmVisionApproval(
            OperatorPresenceRequest request,
            out string denyReason,
            CancellationToken cancellationToken = default)
        {
            if (Interlocked.CompareExchange(ref _inFlight, 1, 0) != 0)
            {
                denyReason = "Operator confirmation already in progress (single-flight).";
                return false;
            }

            try
            {
                if (cancellationToken.WaitHandle.WaitOne(_hold))
                {
                    denyReason = "Operator confirmation cancelled.";
                    return false;
                }

                denyReason = string.Empty;
                return true;
            }
            finally
            {
                Interlocked.Exchange(ref _inFlight, 0);
            }
        }
    }

    [Fact]
    public void Approve_FailsWhenVisibleContentChangesWithoutResize()
    {
        var window = MakeWindow(new IntPtr(0xB2), 202);
        var options = EnabledVisionOptions();
        var store = new VisionCandidateStore(options);
        var probe = new ScriptedLiveWindowProbe(window);
        var holder = new MutableCapture
        {
            Window = window,
            PngBytes = MakeSolidPng(Color.Red, 80, 60)
        };
        var capture = new ScriptedVisionCaptureService((hwnd, _) =>
            hwnd == holder.Window.WindowHandle
                ? new VisionServerCapture
                {
                    Window = holder.Window,
                    SemanticTarget = InaccessibleTarget(),
                    PngBytes = holder.PngBytes
                }
                : null);
        var coordinator = VisionFallbackCoordinator.ForTests(options, store, probe, capture);

        var (_, _, candidates) = coordinator.TryProduceCandidatesFromHandle(
            window.WindowHandle, new Point(40, 50), query: "OK");
        Assert.NotEmpty(candidates);

        // Same bounds / window identity; only pixels change.
        holder.PngBytes = MakeSolidPng(Color.Lime, 80, 60);
        var approval = coordinator.ApproveCandidate(candidates[0].Id!, "click");
        Assert.False(approval.Allowed);
        Assert.Contains("fingerprint", approval.Reason, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Approve_ComparesAgainstOriginalBaseline_NotLiveVsLive()
    {
        var baseline = MakeWindow(new IntPtr(0xC3), 303, new Rectangle(10, 10, 400, 300));
        var drifted = baseline with { DpiScaleX = 1.5, DpiScaleY = 1.5 };
        var options = EnabledVisionOptions();
        var store = new VisionCandidateStore(options);
        var probe = new ScriptedLiveWindowProbe(hwnd =>
            hwnd == baseline.WindowHandle ? drifted : null);
        var png = MakeSolidPng(Color.Navy, 50, 40);
        var capture = new ScriptedVisionCaptureService((hwnd, _) =>
            hwnd == baseline.WindowHandle
                ? new VisionServerCapture
                {
                    Window = drifted,
                    SemanticTarget = InaccessibleTarget(),
                    PngBytes = png
                }
                : null);
        var coordinator = VisionFallbackCoordinator.ForTests(options, store, probe, capture);

        // Issue against original baseline window identity.
        var issued = store.Issue(new VisionCandidateTarget
        {
            Label = "OK",
            Bounds = new Rectangle(40, 40, 80, 24),
            Confidence = 0.9,
            Window = baseline,
            ProducedAtUtc = DateTimeOffset.UtcNow,
            ScreenshotFingerprint = VisionFingerprint.Compute(png, "OK", new Rectangle(40, 40, 80, 24)),
            AutoPromoted = false
        }, DateTimeOffset.UtcNow, new Point(50, 60));

        var approval = coordinator.ApproveCandidate(issued.Id, "click");
        Assert.False(approval.Allowed);
        Assert.Contains("DPI", approval.Reason, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ServerCapture_RejectsMissingOrMismatchedPid_FailClosed()
    {
        var requested = MakeWindow(new IntPtr(0x100), 1000, new Rectangle(0, 0, 400, 300));
        var windowProbe = new ScriptedLiveWindowProbe(requested);
        var pixels = new ScriptedWindowPixelCapture(MakeSolidPng(Color.Gray, 20, 20));

        var uiaMissing = new ScriptedWindowBoundUiaProbe(
            new DebugWindowIdentity
            {
                WindowHandle = requested.WindowHandle,
                ProcessId = null,
                Bounds = requested.Bounds,
                DpiScaleX = 1,
                DpiScaleY = 1
            });
        Assert.Null(new ServerVisionCaptureService(windowProbe, uiaMissing, pixels)
            .Capture(requested.WindowHandle, new Point(50, 60)));

        var zeroPid = new ScriptedWindowBoundUiaProbe(
            requested with { ProcessId = 0 });
        Assert.Null(new ServerVisionCaptureService(windowProbe, zeroPid, pixels)
            .Capture(requested.WindowHandle, new Point(50, 60)));

        var otherPid = new ScriptedWindowBoundUiaProbe(
            requested with { ProcessId = 9999 });
        Assert.Null(new ServerVisionCaptureService(windowProbe, otherPid, pixels)
            .Capture(requested.WindowHandle, new Point(50, 60)));

        var ok = new ScriptedWindowBoundUiaProbe(requested);
        Assert.NotNull(new ServerVisionCaptureService(windowProbe, ok, pixels)
            .Capture(requested.WindowHandle, new Point(50, 60)));
    }

    [Fact]
    public void ServerCapture_UsesRealPixelBytes_NotHwndTextPlaceholder()
    {
        var window = MakeWindow(new IntPtr(0xD4), 404);
        var png = MakeSolidPng(Color.Orange, 32, 24);
        var pixels = new ScriptedWindowPixelCapture(png);
        var capture = new ServerVisionCaptureService(
            new ScriptedLiveWindowProbe(window),
            new ScriptedWindowBoundUiaProbe(window),
            pixels);

        var result = capture.Capture(window.WindowHandle, new Point(20, 20));
        Assert.NotNull(result);
        Assert.Equal(png, result!.PngBytes);
        Assert.False(Encoding.UTF8.GetString(result.PngBytes).Contains("server-vision-capture", StringComparison.Ordinal));
        Assert.True(result.PngBytes.Length >= 32);
        // PNG magic
        Assert.Equal(0x89, result.PngBytes[0]);
        Assert.Equal((byte)'P', result.PngBytes[1]);
        Assert.Equal((byte)'N', result.PngBytes[2]);
        Assert.Equal((byte)'G', result.PngBytes[3]);
    }

    private static VisionFallbackOptions EnabledVisionOptions() => new()
    {
        Enabled = true,
        AllowWhenUiaInaccessible = true,
        Provider = "fake",
        MinimumConfidence = 0.5,
        MaxScreenshotAgeMilliseconds = 5000,
        CandidateTtlMilliseconds = 60_000,
        ExecutionTokenTtlMilliseconds = 30_000
    };

    private static DebugWindowIdentity MakeWindow(
        IntPtr hwnd,
        int pid,
        Rectangle? bounds = null) => new()
    {
        WindowHandle = hwnd,
        ProcessId = pid,
        Bounds = bounds ?? new Rectangle(0, 0, 400, 300),
        DpiScaleX = 1,
        DpiScaleY = 1
    };

    private static DebugUiaNode InaccessibleTarget() => new()
    {
        ControlType = "Custom",
        IsSemanticAccessible = false,
        InaccessibilityReason = "custom-drawn"
    };

    private static byte[] MakeSolidPng(Color color, int width, int height)
    {
        using var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb);
        using (var g = Graphics.FromImage(bitmap))
            g.Clear(color);
        using var stream = new MemoryStream();
        bitmap.Save(stream, ImageFormat.Png);
        return stream.ToArray();
    }

    private sealed class MutableCapture
    {
        public required DebugWindowIdentity Window { get; init; }
        public required byte[] PngBytes { get; set; }
    }

    private sealed class ScriptedWindowBoundUiaProbe : IUiaElementProbe
    {
        private readonly DebugWindowIdentity _window;

        public ScriptedWindowBoundUiaProbe(DebugWindowIdentity window) => _window = window;

        public DebugTreeCapture? CaptureAtPoint(Point screenPoint, int maxAncestry = 8, int maxChildren = 25) =>
            UiaTreeCaptureService.FromNodes(
                new DebugUiaNode
                {
                    ControlType = "Custom",
                    IsSemanticAccessible = false,
                    InaccessibilityReason = "custom-drawn",
                    Bounds = new Rectangle(screenPoint.X - 5, screenPoint.Y - 5, 20, 20)
                },
                window: _window,
                cursor: screenPoint);

        public LocatorMatchProbeResult CountLocatorMatches(
            IReadOnlyDictionary<string, object?> locator,
            DebugWindowIdentity? scope) =>
            LocatorMatchProbeResult.Ok(0);

        public HoverProbeResult ProbeHover(Point screenPoint) =>
            new() { ScreenPoint = screenPoint, NoOp = true };
    }

    [Fact]
    public void DebugSessionSnapshot_JsonSerializesWindowHandleAsHex_NotIntPtr()
    {
        // Repro for Windows acceptance: /debug/inspect returned 500 because STJ cannot serialize IntPtr.
        var options = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
        };

        // Use long ctor so the handle is unambiguous across runtimes (not a truncated int cast).
        var handle = new IntPtr(0x1234ABCDL);
        Assert.Equal("0x1234ABCD", IntPtrHexJsonConverter.ToHex(handle));

        var snapshot = new DebugSessionSnapshot
        {
            IsActive = true,
            SessionId = "dbg-test",
            StatusText = "Inspect: Save",
            LastCapture = UiaTreeCaptureService.FromNodes(
                new DebugUiaNode
                {
                    AutomationId = "SaveButton",
                    ControlType = "Button",
                    Bounds = new Rectangle(1, 2, 3, 4)
                },
                window: new DebugWindowIdentity
                {
                    WindowTitle = "App",
                    WindowHandle = handle,
                    Bounds = new Rectangle(0, 0, 100, 80),
                    DpiScaleX = 1,
                    DpiScaleY = 1
                },
                cursor: new Point(10, 20))
        };

        var json = JsonSerializer.Serialize(snapshot, options);
        Assert.Contains("\"windowHandle\":\"0x1234ABCD\"", json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("IntPtr", json, StringComparison.Ordinal);

        // Round-trip must not throw and must restore the handle (JsonConverter, not JsonIgnore).
        var roundTrip = JsonSerializer.Deserialize<DebugSessionSnapshot>(json, options);
        Assert.NotNull(roundTrip);
        Assert.Equal(handle, roundTrip!.LastCapture!.Window!.WindowHandle);
        Assert.Equal("0x1234ABCD", roundTrip.LastCapture.Window.WindowHandleHex);
    }

    private sealed class ScriptedUiaElementProbe : IUiaElementProbe
    {
        public LocatorMatchProbeResult MatchResult { get; set; } = LocatorMatchProbeResult.Ok(1);

        public DebugTreeCapture? CaptureAtPoint(Point screenPoint, int maxAncestry = 8, int maxChildren = 25)
        {
            var target = new DebugUiaNode
            {
                AutomationId = "SaveButton",
                Name = "Save",
                ControlType = "Button",
                ClassName = "Button",
                Bounds = new Rectangle(screenPoint.X - 10, screenPoint.Y - 5, 80, 24),
                AvailablePatterns = ["Invoke"],
                Path = "target",
                IsSemanticAccessible = true
            };
            return UiaTreeCaptureService.FromNodes(
                target,
                ancestry:
                [
                    new DebugUiaNode { Name = "Main", ControlType = "Window", Depth = 1, Path = "ancestor/1" }
                ],
                children: [],
                window: new DebugWindowIdentity
                {
                    WindowTitle = "App",
                    WindowHandle = new IntPtr(42),
                    ProcessId = 100,
                    Bounds = new Rectangle(0, 0, 800, 600),
                    DpiScaleX = 1,
                    DpiScaleY = 1
                },
                cursor: screenPoint);
        }

        public LocatorMatchProbeResult CountLocatorMatches(
            IReadOnlyDictionary<string, object?> locator,
            DebugWindowIdentity? scope) =>
            MatchResult;

        public HoverProbeResult ProbeHover(Point screenPoint) =>
            new() { ScreenPoint = screenPoint, NoOp = true };
    }

    private sealed class InaccessibleUiaElementProbe : IUiaElementProbe
    {
        public DebugTreeCapture? CaptureAtPoint(Point screenPoint, int maxAncestry = 8, int maxChildren = 25)
        {
            var target = new DebugUiaNode
            {
                ControlType = "Pane",
                AutomationId = null,
                Name = null,
                Bounds = new Rectangle(screenPoint.X - 20, screenPoint.Y - 10, 280, 80),
                Path = "target",
                IsSemanticAccessible = false,
                InaccessibilityReason = "custom-drawn-or-empty-identity"
            };
            return UiaTreeCaptureService.FromNodes(
                target,
                window: new DebugWindowIdentity
                {
                    WindowTitle = "App",
                    WindowHandle = new IntPtr(99),
                    ProcessId = 200,
                    Bounds = new Rectangle(0, 0, 800, 600),
                    DpiScaleX = 1,
                    DpiScaleY = 1
                },
                cursor: screenPoint);
        }

        public LocatorMatchProbeResult CountLocatorMatches(
            IReadOnlyDictionary<string, object?> locator,
            DebugWindowIdentity? scope) =>
            LocatorMatchProbeResult.Ok(0);

        public HoverProbeResult ProbeHover(Point screenPoint) =>
            new() { ScreenPoint = screenPoint, NoOp = true };
    }

    private static string FindRepoRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "DesktopAutomationHelper.sln"))
                || File.Exists(Path.Combine(dir.FullName, "src", "DesktopAutomationTestApp", "CustomDrawnCanvas.cs")))
            {
                return dir.FullName;
            }

            dir = dir.Parent;
        }

        throw new InvalidOperationException("Could not locate repo root from test base directory.");
    }
}

file static class GuidExtensions
{
    public static string ToN(this Guid guid) => guid.ToString("N");
}
