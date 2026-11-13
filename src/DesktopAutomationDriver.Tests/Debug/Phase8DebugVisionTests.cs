using System.Reflection;
using System.Text;
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
            Assert.Contains("requiresHumanApproval", json, StringComparison.Ordinal);

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
    public void CoordinateGuard_RejectsStaleDpiAndLayoutChanges_AndDestructive()
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

        var destructive = guard.Evaluate(new CoordinateActionRequest
        {
            Candidate = fresh,
            CurrentWindow = window,
            NowUtc = DateTimeOffset.UtcNow,
            Action = "click",
            IsDestructive = true
        });
        Assert.False(destructive.Allowed);

        var ok = guard.Evaluate(new CoordinateActionRequest
        {
            Candidate = fresh,
            CurrentWindow = window,
            NowUtc = DateTimeOffset.UtcNow,
            Action = "click"
        });
        Assert.True(ok.Allowed);
        Assert.NotNull(ok.ScreenPoint);
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
        Assert.False(UiaTreeCaptureService.IsSemanticallyInaccessible(new DebugUiaNode
        {
            AutomationId = "Ok",
            ControlType = "Button",
            IsSemanticAccessible = true
        }));
    }

    private sealed class ScriptedUiaElementProbe : IUiaElementProbe
    {
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
                    Bounds = new Rectangle(0, 0, 800, 600),
                    DpiScaleX = 1,
                    DpiScaleY = 1
                },
                cursor: screenPoint);
        }

        public int CountLocatorMatches(IReadOnlyDictionary<string, object?> locator, DebugWindowIdentity? scope) =>
            1;

        public HoverProbeResult ProbeHover(Point screenPoint) =>
            new() { ScreenPoint = screenPoint, NoOp = true };
    }
}

file static class GuidExtensions
{
    public static string ToN(this Guid guid) => guid.ToString("N");
}
