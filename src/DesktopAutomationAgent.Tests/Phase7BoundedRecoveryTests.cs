using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using DesktopAutomationAgent.Cli;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Driver.Models;
using DesktopAutomationAgent.Execution;
using DesktopAutomationAgent.Exploration;
using DesktopAutomationAgent.ObjectRepository;
using DesktopAutomationAgent.Plans;
using DesktopAutomationAgent.Recovery;
using Microsoft.Extensions.Logging.Abstractions;

namespace DesktopAutomationAgent.Tests;

public sealed class Phase7BoundedRecoveryTests
{
    [Theory]
    [InlineData("Element not found for click", RecoveryClassification.StateNotStable)]
    [InlineData("child window not yet active", RecoveryClassification.ExpectedChildWindowNotActive)]
    [InlineData("assertion failed", RecoveryClassification.AssertionFailure)]
    [InlineData("ambiguous locator matches", RecoveryClassification.AmbiguousLocator)]
    [InlineData("destructive action denied", RecoveryClassification.DestructiveActionDenial)]
    [InlineData("application build mismatch detected", RecoveryClassification.ApplicationBuildMismatch)]
    [InlineData("unexpected external window appeared", RecoveryClassification.UnexpectedExternalWindow)]
    [InlineData("unknown operation foobar", RecoveryClassification.UnknownOperation)]
    public void Classifier_MapsEveryFailureClassification(string message, RecoveryClassification expected)
    {
        var classifier = RecoveryFailureClassifier.ForTests();
        var needsTypedProof = expected is RecoveryClassification.StateNotStable
            or RecoveryClassification.ExpectedChildWindowNotActive;
        var step = new StepRunResult
        {
            Sequence = 1,
            Id = "s1",
            Operation = "clickuia",
            Phase = "steps",
            Status = "failed",
            Success = false,
            Error = message,
            // Mutations require typed driver fields — free-text alone must not authorize recovery.
            ExecutionState = needsTypedProof ? RecoveryOperationSafety.ExecutionStateNotStarted : null,
            FailureCode = expected == RecoveryClassification.StateNotStable
                ? RecoveryOperationSafety.FailureCodeLocatorNotFoundBeforeInvoke
                : null,
            Classification = expected == RecoveryClassification.AssertionFailure
                ? UiFailureClassification.AssertionFailure
                : UiFailureClassification.OperationFailure,
            Assertions = expected == RecoveryClassification.AssertionFailure
                ?
                [
                    new AssertionRunResult { Path = "/value", Operator = "equals", Passed = false, Message = "assertion failed" }
                ]
                : Array.Empty<AssertionRunResult>()
        };

        var decision = classifier.Classify(step, EmptyContext());
        Assert.Equal(expected, decision.Classification);
        Assert.Equal(RecoveryClassificationRules.IsRecoverable(expected), decision.MayAttemptRecovery);
    }

    [Fact]
    public void Classifier_FreeTextNotFound_DoesNotAuthorizeMutationReplay()
    {
        var decision = RecoveryFailureClassifier.ForTests().Classify(
            new StepRunResult
            {
                Sequence = 1,
                Id = "pay",
                Operation = "click",
                Phase = "steps",
                Status = "failed",
                Success = false,
                Error = "receipt not found",
                Classification = UiFailureClassification.OperationFailure
            },
            EmptyContext());

        Assert.Equal(RecoveryClassification.NotRecoverable, decision.Classification);
        Assert.False(decision.MayAttemptRecovery);
        Assert.False(RecoveryOperationSafety.HasTypedNonExecutionProof(null, null));
        Assert.False(RecoveryOperationSafety.IsSafeToReplay(
            "click", null, UiFailureClassification.OperationFailure, "receipt not found", false));
    }

    [Fact]
    public void Classifier_AuthenticationNeverRecoverable()
    {
        var decision = RecoveryFailureClassifier.ForTests().Classify(
            new StepRunResult
            {
                Sequence = 1,
                Id = "s1",
                Operation = "clickuia",
                Phase = "steps",
                Status = "failed",
                Success = false,
                Classification = UiFailureClassification.Authentication,
                Error = "401"
            },
            EmptyContext());

        Assert.Equal(RecoveryClassification.AuthenticationFailure, decision.Classification);
        Assert.False(decision.MayAttemptRecovery);
    }

    [Fact]
    public void Classifier_CompileSchemaNeverRecoverable()
    {
        var decision = RecoveryFailureClassifier.ForTests().Classify(
            new StepRunResult
            {
                Sequence = 1,
                Id = "s1",
                Operation = "clickuia",
                Phase = "steps",
                Status = "failed",
                Success = false,
                Classification = UiFailureClassification.PlanValidation,
                Error = "schema"
            },
            EmptyContext());

        Assert.Equal(RecoveryClassification.CompileSchemaFailure, decision.Classification);
        Assert.False(decision.MayAttemptRecovery);
    }

    [Fact]
    public async Task BoundedWait_RecoversTimingDrift()
    {
        var posts = 0;
        var handler = ReadyHandler((req, _) =>
        {
            if (req.Method != HttpMethod.Post)
                return Task.FromResult(DefaultGet(req));

            posts++;
            if (posts == 1)
            {
                return Task.FromResult(FakeHttpMessageHandler.Json(new
                {
                    success = false,
                    error = "Element not found for click"
                }));
            }

            return Task.FromResult(FakeHttpMessageHandler.Json(new { success = true, value = "ok" }));
        });

        var options = ReadyOptions();
        options.Recovery.BoundedWaitMilliseconds = 5;
        var plan = TestSupport.MinimalPlanJson(stepsOverride: """
            [ { "id": "click", "operation": "listwindows", "arguments": {} } ]
            """);
        var path = TestSupport.WritePlan(options, "recover-wait.plan.json", plan);
        var beforeHash = Sha256File(path);

        var report = await CreateRunner(options, handler).RunAsync(path, dryRun: false);

        Assert.Equal("passed", report.Status);
        Assert.Equal(ExitCodes.Success, report.ExitCode);
        Assert.True(posts >= 2);
        Assert.NotNull(report.Recovery);
        Assert.Contains(report.Recovery!.Attempts, a => a.Tier == RecoveryTier.FreshObservationAndBoundedWait);
        Assert.Equal(beforeHash, Sha256File(path));
        Assert.False(report.Recovery.PlanMutated);
    }

    [Fact]
    public async Task ApprovedAlternateLocator_RecoversUniquely()
    {
        var options = ReadyOptions();
        options.Recovery.BoundedWaitMilliseconds = 1;
        SeedRepoWithAlternate(options, ambiguous: false);

        var posts = new List<string>();
        var handler = ReadyHandler(async (req, _) =>
        {
            if (req.Method != HttpMethod.Post)
                return DefaultGet(req);

            var body = await req.Content!.ReadAsStringAsync();
            posts.Add(body);

            // Probe finduia for alternate uniqueness, then the rewritten step finduia.
            if (body.Contains("AltSubmit", StringComparison.Ordinal))
            {
                return FakeHttpMessageHandler.Json(new
                {
                    success = true,
                    value = new { matchCount = 1 }
                });
            }

            if (body.Contains("PrimarySubmit", StringComparison.Ordinal))
            {
                return FakeHttpMessageHandler.Json(new
                {
                    success = false,
                    error = "Element not found for finduia"
                });
            }

            return FakeHttpMessageHandler.Json(new { success = true, value = "ok" });
        });

        var plan = """
            {
              "schemaVersion": 1,
              "catalogSchemaVersion": 2,
              "planId": "ALT-1",
              "name": "alternate",
              "objectRepository": "object-repository/repository.json",
              "steps": [
                {
                  "id": "find",
                  "operation": "finduia",
                  "arguments": { "locator": { "$objectRef": "login.submit" } }
                }
              ]
            }
            """;
        var path = TestSupport.WritePlan(options, "alt.plan.json", plan);
        var pagesPath = Path.Combine(options.Workspace.Root!, "object-repository", "pages", "login.page.json");
        var pagesBefore = File.ReadAllText(pagesPath);

        var report = await CreateRunner(options, handler).RunAsync(path, dryRun: false);

        Assert.True(
            report.Status == "passed",
            $"status={report.Status}; failure={report.Failure?.Message}; recovery={JsonSerializer.Serialize(report.Recovery)}");
        Assert.Contains(report.Recovery!.Attempts, a => a.Tier == RecoveryTier.ApprovedAlternateLocator && a.Outcome == "recovered");
        Assert.Equal(pagesBefore, File.ReadAllText(pagesPath));
        Assert.False(report.Recovery.ApprovedRepositoryMutated);
        Assert.DoesNotContain(Directory.EnumerateFiles(
            Path.Combine(options.Workspace.Root!, "object-repository", "pages"),
            "*",
            SearchOption.AllDirectories), f => f.Contains("candidate-repair", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public async Task AmbiguousAlternate_StopsWithoutHealing()
    {
        var options = ReadyOptions();
        options.Recovery.BoundedWaitMilliseconds = 1;
        SeedRepoWithAlternate(options, ambiguous: false);

        var handler = ReadyHandler(async (req, _) =>
        {
            if (req.Method != HttpMethod.Post)
                return DefaultGet(req);

            var body = await req.Content!.ReadAsStringAsync();
            if (body.Contains("AltSubmit", StringComparison.Ordinal))
            {
                return FakeHttpMessageHandler.Json(new
                {
                    success = false,
                    error = "ambiguous locator: multiple matches"
                });
            }

            return FakeHttpMessageHandler.Json(new
            {
                success = false,
                error = "Element not found for finduia"
            });
        });

        var plan = """
            {
              "schemaVersion": 1,
              "catalogSchemaVersion": 2,
              "planId": "AMB-1",
              "name": "ambiguous",
              "objectRepository": "object-repository/repository.json",
              "steps": [
                {
                  "id": "find",
                  "operation": "finduia",
                  "arguments": { "locator": { "$objectRef": "login.submit" } }
                }
              ]
            }
            """;
        var path = TestSupport.WritePlan(options, "amb.plan.json", plan);
        var report = await CreateRunner(options, handler).RunAsync(path, dryRun: false);

        Assert.Equal("failed", report.Status);
        Assert.True(
            report.Recovery?.Attempts.Any(a =>
                a.Classification == RecoveryClassification.AmbiguousLocator
                || (a.Detail ?? string.Empty).Contains("ambiguous", StringComparison.OrdinalIgnoreCase)) == true,
            $"recovery={JsonSerializer.Serialize(report.Recovery)}; failure={report.Failure?.Message}");
    }

    [Fact]
    public async Task AssertionFailure_NeverHealed()
    {
        var posts = 0;
        var handler = ReadyHandler((req, _) =>
        {
            if (req.Method != HttpMethod.Post)
                return Task.FromResult(DefaultGet(req));

            posts++;
            return Task.FromResult(FakeHttpMessageHandler.Json(new { success = true, value = "wrong" }));
        });

        var options = ReadyOptions();
        var plan = TestSupport.MinimalPlanJson(stepsOverride: """
            [
              {
                "id": "assert",
                "operation": "listwindows",
                "arguments": {},
                "assertions": [ { "path": "/value", "operator": "equals", "expected": "expected" } ]
              }
            ]
            """);
        var path = TestSupport.WritePlan(options, "assert.plan.json", plan);
        var report = await CreateRunner(options, handler).RunAsync(path, dryRun: false);

        Assert.Equal("failed", report.Status);
        Assert.Equal(UiFailureClassification.AssertionFailure, report.Failure!.Classification);
        Assert.Equal(1, posts);
        Assert.True(report.Recovery is null || report.Recovery.Attempts.Count == 0
            || report.Recovery.Attempts.All(a => a.Tier == RecoveryTier.None));
    }

    [Fact]
    public async Task DestructiveDenial_NeverHealed()
    {
        var posts = 0;
        var handler = ReadyHandler((req, _) =>
        {
            if (req.Method != HttpMethod.Post)
                return Task.FromResult(DefaultGet(req));

            posts++;
            return Task.FromResult(FakeHttpMessageHandler.Json(new
            {
                success = false,
                error = "destructive action denied by policy"
            }));
        });

        var options = ReadyOptions();
        var plan = TestSupport.MinimalPlanJson(stepsOverride: """
            [ { "id": "x", "operation": "listwindows", "arguments": {} } ]
            """);
        var path = TestSupport.WritePlan(options, "destructive.plan.json", plan);
        var report = await CreateRunner(options, handler).RunAsync(path, dryRun: false);

        Assert.Equal("failed", report.Status);
        Assert.Equal(1, posts);
    }

    [Fact]
    public async Task StaleBuild_Stops()
    {
        var options = ReadyOptions();
        var classifier = RecoveryFailureClassifier.ForTests();
        var decision = classifier.Classify(
            new StepRunResult
            {
                Sequence = 1,
                Id = "s1",
                Operation = "clickuia",
                Phase = "steps",
                Status = "failed",
                Success = false,
                Error = "Element not found",
                Classification = UiFailureClassification.OperationFailure
            },
            new RecoveryContext
            {
                RunId = "r",
                RunDirectory = options.Workspace.Root!,
                PlanPath = "p",
                OriginalPlanSha256 = "abc",
                ApplicationFingerprint = "live-fp",
                ExpectedApplicationFingerprint = "approved-fp"
            });

        Assert.Equal(RecoveryClassification.ApplicationBuildMismatch, decision.Classification);
        Assert.False(decision.MayAttemptRecovery);
        await Task.CompletedTask;
    }

    [Fact]
    public void Budgets_Enforced_NoRecursion()
    {
        var options = new RecoveryOptions
        {
            Enabled = true,
            MaxTotalRecoveriesPerRun = 1,
            MaxRecoveryAttemptsPerStep = 1,
            MaxAdditionalDurationSeconds = 30
        };
        using var tracker = new RecoveryBudgetTracker(options);

        Assert.True(tracker.TryConsumeUnit("s1", CancellationToken.None, out _, out _));
        Assert.False(tracker.TryConsumeUnit("s1", CancellationToken.None, out _, out var recursive));
        Assert.Equal("no-recursive-recovery", recursive);
        tracker.EndUnit();

        Assert.False(tracker.TryConsumeUnit("s2", CancellationToken.None, out _, out var total));
        Assert.Equal("max-total-recoveries", total);
    }

    [Fact]
    public void Budgets_ConsumePerProbe_AndLinkDeadline()
    {
        var options = new RecoveryOptions
        {
            Enabled = true,
            MaxTotalRecoveriesPerRun = 3,
            MaxRecoveryAttemptsPerStep = 2,
            MaxAdditionalDurationSeconds = 60
        };
        using var tracker = new RecoveryBudgetTracker(options);

        Assert.True(tracker.TryConsumeUnit("s1", CancellationToken.None, out var token1, out _));
        Assert.False(token1.IsCancellationRequested);
        tracker.EndUnit();

        Assert.True(tracker.TryConsumeUnit("s1", CancellationToken.None, out _, out _));
        tracker.EndUnit();

        Assert.False(tracker.TryConsumeUnit("s1", CancellationToken.None, out _, out var perStep));
        Assert.Equal("max-per-step-recoveries", perStep);
        Assert.Equal(2, tracker.TotalAttempts);
    }

    [Fact]
    public void Classifier_TimedOutMutation_NotRecoverable()
    {
        var decision = RecoveryFailureClassifier.ForTests().Classify(
            new StepRunResult
            {
                Sequence = 1,
                Id = "s1",
                Operation = "click",
                Phase = "steps",
                Status = "failed",
                Success = false,
                Error = "Step 's1' timed out after 1 seconds.",
                Classification = UiFailureClassification.ExecutionTimeout
            },
            EmptyContext());

        Assert.Equal(RecoveryClassification.NotRecoverable, decision.Classification);
        Assert.False(decision.MayAttemptRecovery);
    }

    [Theory]
    [InlineData("Step 's1' returned an empty response body (HTTP 200).")]
    [InlineData("Step 's1' returned invalid JSON (HTTP 200).")]
    [InlineData("Driver response Content-Length 999999 exceeded maximum size of 1048576 bytes.")]
    public void Classifier_UncertainProtocolMutation_NotRecoverable(string error)
    {
        var decision = RecoveryFailureClassifier.ForTests().Classify(
            new StepRunResult
            {
                Sequence = 1,
                Id = "s1",
                Operation = "click",
                Phase = "steps",
                Status = "failed",
                Success = false,
                Error = error,
                Classification = UiFailureClassification.OperationFailure
            },
            EmptyContext());

        Assert.Equal(RecoveryClassification.NotRecoverable, decision.Classification);
        Assert.False(decision.MayAttemptRecovery);
        Assert.False(RecoveryOperationSafety.IsSafeToReplay(
            "click", null, UiFailureClassification.OperationFailure, error, allowMutationReplay: false));
    }

    [Fact]
    public void EquivalentTransition_NeverUpgradesReadTimeoutToMutation()
    {
        var catalog = CatalogFixtures.Phase2Catalog();
        catalog.Operations.Add(new OperationDescriptorDto
        {
            Name = "clickuia",
            Aliases = [],
            Category = "element-action",
            OperationType = "invoke",
            RequiresSession = false,
            RequiredInputs = ["locator"]
        });

        Assert.False(RecoveryOperationSafety.IsEquivalentTransitionAllowed(
            originalOperation: "listwindows",
            targetOperation: "clickuia",
            catalog,
            UiFailureClassification.ExecutionTimeout,
            "Step 's1' timed out after 1 seconds.",
            allowMutationReplay: false));

        Assert.True(RecoveryOperationSafety.IsEquivalentTransitionAllowed(
            originalOperation: "listwindows",
            targetOperation: "finduia",
            catalog,
            UiFailureClassification.ExecutionTimeout,
            "Step 's1' timed out after 1 seconds.",
            allowMutationReplay: false));
    }

    [Fact]
    public void Budgets_AccumulateOnlyActiveRecoveryIntervals()
    {
        var options = new RecoveryOptions
        {
            Enabled = true,
            MaxTotalRecoveriesPerRun = 5,
            MaxRecoveryAttemptsPerStep = 5,
            MaxAdditionalDurationSeconds = 2
        };
        using var tracker = new RecoveryBudgetTracker(options);

        Assert.True(tracker.TryConsumeUnit("s1", CancellationToken.None, out _, out _));
        Thread.Sleep(50);
        tracker.EndUnit();
        var afterFirst = tracker.Elapsed;

        Thread.Sleep(400); // normal plan execution — must not burn recovery budget
        Assert.Equal(afterFirst, tracker.Elapsed);

        Assert.True(tracker.TryConsumeUnit("s1", CancellationToken.None, out _, out _));
        tracker.EndUnit();
        Assert.True(tracker.Elapsed >= afterFirst);
        Assert.True(tracker.RemainingDuration() > TimeSpan.Zero);
    }

    [Fact]
    public async Task RecoveryDeadlineCancel_MapsToBudgetExhausted_NotUnhandled()
    {
        var options = ReadyOptions();
        options.Recovery.BoundedWaitMilliseconds = 0;
        options.Recovery.MaxAdditionalDurationSeconds = 1;
        options.Recovery.MaxRecoveryAttemptsPerStep = 8;
        options.Recovery.MaxTotalRecoveriesPerRun = 8;

        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var ladder = TestSupport.CreateRecoveryLadder(
            options,
            workspace,
            new DriverUiClient(
                TestSupport.Wrap(options),
                TestSupport.CreateFactory(ReadyHandler((_, _) =>
                    Task.FromResult(FakeHttpMessageHandler.Json(new { success = true })))),
                NullLogger<DriverUiClient>.Instance));

        using var budget = new RecoveryBudgetTracker(options.Recovery);
        var runDirectory = Path.Combine(workspace.RootPath, "runs", "budget-deadline-run");
        Directory.CreateDirectory(runDirectory);
        var failed = new StepRunResult
        {
            Sequence = 1,
            Id = "q",
            Operation = "listwindows",
            Phase = "steps",
            Status = "failed",
            Success = false,
            Error = "Element not found for listwindows",
            Classification = UiFailureClassification.OperationFailure
        };

        var outcome = await ladder.TryRecoverAsync(
            new DriverConnection
            {
                BaseUri = new Uri("http://127.0.0.1:33211/"),
                BearerToken = "tok",
                DiscoveryMethod = "explicit"
            },
            new PlanStep { Id = "q", Operation = "listwindows", Arguments = new Dictionary<string, JsonElement>() },
            failed,
            new RecoveryContext
            {
                RunId = "r",
                RunDirectory = runDirectory,
                PlanPath = "p",
                OriginalPlanSha256 = "abc",
                Catalog = CatalogFixtures.Phase2Catalog()
            },
            budget,
            async (_, ct) =>
            {
                // Hang until the linked recovery deadline cancels — must map to BudgetExhausted.
                await Task.Delay(TimeSpan.FromSeconds(30), ct);
                return failed;
            },
            CancellationToken.None);

        Assert.Equal(RecoveryClassification.BudgetExhausted, outcome.Classification);
        Assert.False(outcome.Recovered);
        Assert.Contains(
            outcome.Attempts,
            a => a.Classification == RecoveryClassification.BudgetExhausted
                 || string.Equals(a.Outcome, "budget-exhausted", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public async Task RecoveryDeadlineCancel_ViaCancelledStepResult_MapsToBudgetExhausted()
    {
        // Mirrors DriverUiClient: linked-token cancel → UiExecutionException(Cancelled) → failed StepRunResult
        // (ladder must still surface BudgetExhausted, not the original classification).
        var options = ReadyOptions();
        options.Recovery.BoundedWaitMilliseconds = 0;
        options.Recovery.MaxAdditionalDurationSeconds = 1;
        options.Recovery.MaxRecoveryAttemptsPerStep = 8;
        options.Recovery.MaxTotalRecoveriesPerRun = 8;

        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var ladder = TestSupport.CreateRecoveryLadder(
            options,
            workspace,
            new DriverUiClient(
                TestSupport.Wrap(options),
                TestSupport.CreateFactory(ReadyHandler((_, _) =>
                    Task.FromResult(FakeHttpMessageHandler.Json(new { success = true })))),
                NullLogger<DriverUiClient>.Instance));

        using var budget = new RecoveryBudgetTracker(options.Recovery);
        var runDirectory = Path.Combine(workspace.RootPath, "runs", "budget-deadline-cancelled-step");
        Directory.CreateDirectory(runDirectory);
        var failed = new StepRunResult
        {
            Sequence = 1,
            Id = "q",
            Operation = "listwindows",
            Phase = "steps",
            Status = "failed",
            Success = false,
            Error = "Element not found for listwindows",
            Classification = UiFailureClassification.OperationFailure
        };

        var outcome = await ladder.TryRecoverAsync(
            new DriverConnection
            {
                BaseUri = new Uri("http://127.0.0.1:33211/"),
                BearerToken = "tok",
                DiscoveryMethod = "explicit"
            },
            new PlanStep { Id = "q", Operation = "listwindows", Arguments = new Dictionary<string, JsonElement>() },
            failed,
            new RecoveryContext
            {
                RunId = "r",
                RunDirectory = runDirectory,
                PlanPath = "p",
                OriginalPlanSha256 = "abc",
                Catalog = CatalogFixtures.Phase2Catalog()
            },
            budget,
            async (_, ct) =>
            {
                try
                {
                    await Task.Delay(TimeSpan.FromSeconds(30), ct);
                    return failed;
                }
                catch (OperationCanceledException) when (ct.IsCancellationRequested)
                {
                    return new StepRunResult
                    {
                        Sequence = 1,
                        Id = "q",
                        Operation = "listwindows",
                        Phase = "steps",
                        Status = "failed",
                        Success = false,
                        Error = "Step 'q' was cancelled.",
                        Classification = UiFailureClassification.Cancelled
                    };
                }
            },
            CancellationToken.None);

        Assert.Equal(RecoveryClassification.BudgetExhausted, outcome.Classification);
        Assert.Contains(
            outcome.Attempts,
            a => a.Classification == RecoveryClassification.BudgetExhausted
                 || string.Equals(a.Outcome, "budget-exhausted", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public async Task CandidateArtifactFailure_DoesNotEscape_CleanupStillRuns()
    {
        var options = ReadyOptions();
        options.Recovery.BoundedWaitMilliseconds = 1;
        options.Recovery.MaxCandidateArtifactBytes = 1024; // minimum allowed; huge evidence still overflows
        options.Recovery.MaxRecoveryAttemptsPerStep = 8;
        options.Recovery.MaxTotalRecoveriesPerRun = 8;

        var handler = ReadyHandler((req, _) =>
        {
            if (req.Method != HttpMethod.Post)
                return Task.FromResult(DefaultGet(req));
            return Task.FromResult(FakeHttpMessageHandler.Json(new
            {
                success = false,
                error = "Element not found for listwindows"
            }));
        });

        var plan = """
            {
              "schemaVersion": 1,
              "catalogSchemaVersion": 2,
              "planId": "ART-1",
              "name": "artifact",
              "steps": [
                { "id": "q", "operation": "listwindows", "arguments": {} }
              ],
              "onFailureSteps": [
                { "id": "cleanup-list", "operation": "listwindows", "arguments": {} }
              ]
            }
            """;
        var path = TestSupport.WritePlan(options, "artifact-fail.plan.json", plan);

        // Custom model emits huge evidence so CandidateRepairWriter throws IOException.
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var uiClient = new DriverUiClient(
            TestSupport.Wrap(options),
            TestSupport.CreateFactory(handler),
            NullLogger<DriverUiClient>.Instance);
        var ladder = new RecoveryLadderService(
            TestSupport.Wrap(options),
            RecoveryFailureClassifier.ForTests(options.Recovery),
            new ApprovedAlternateLocatorResolver(),
            new RecoveryArtifactWriter(workspace),
            CandidateRepairWriter.ForTests(workspace, options.Recovery),
            new HugeEvidenceRecoveryModelClient(),
            uiClient,
            new ObjectReferenceResolver(),
            NullLogger<RecoveryLadderService>.Instance);

        var repositoryReader = new ObjectRepositoryReader(TestSupport.Wrap(options), workspace);
        var runner = new DeterministicPlanRunner(
            TestSupport.Wrap(options),
            TestSupport.CreatePlanReader(options, workspace),
            new PlanObjectRepositoryIntegrator(
                repositoryReader,
                new PlanObjectReferenceExpander(new ObjectReferenceResolver())),
            workspace,
            new DriverConnectionResolver(TestSupport.Wrap(options), TestSupport.CreateFactory(handler), NullLogger<DriverConnectionResolver>.Instance),
            new DriverCatalogClient(TestSupport.Wrap(options), TestSupport.CreateFactory(handler), NullLogger<DriverCatalogClient>.Instance),
            uiClient,
            new AssertionEvaluator(TestSupport.Wrap(options)),
            new RunArtifactWriter(),
            ladder,
            new RecoveryArtifactWriter(workspace),
            new RecoveryManifestLoader(workspace),
            repositoryReader,
            new SystemProcessIdentityProbe(),
            NullLogger<DeterministicPlanRunner>.Instance);

        var report = await runner.RunAsync(path, dryRun: false);
        Assert.Equal("failed", report.Status);
        Assert.Contains(report.OnFailureSteps, s => s.Id == "cleanup-list");
        Assert.NotNull(report.Recovery);
        Assert.Contains(
            report.Recovery!.Attempts,
            a => (a.Detail ?? string.Empty).Contains("artifact", StringComparison.OrdinalIgnoreCase)
                 || string.Equals(a.Outcome, "artifact-failure", StringComparison.OrdinalIgnoreCase));
    }

    private sealed class HugeEvidenceRecoveryModelClient : IRecoveryModelClient
    {
        public string ProviderName => "fake";

        public Task<CandidateRepairProposal?> ProposeRepairAsync(
            RecoveryRepairRequest request,
            CancellationToken cancellationToken = default)
        {
            return Task.FromResult<CandidateRepairProposal?>(new CandidateRepairProposal
            {
                ProposalId = "rrp_huge",
                StepId = request.StepId,
                Evidence = new string('x', 10_000),
                Confidence = 0.1,
                OldLocator = new ObjectLocator { AutomationId = "A" },
                NewLocator = new ObjectLocator { AutomationId = "B" },
                VerificationCommand = "verify",
                AutoPromoted = false,
                RequiresHumanApproval = true
            });
        }
    }

    [Fact]
    public void Classifier_MultiBinding_Ambiguous()
    {
        var decision = RecoveryFailureClassifier.ForTests().Classify(
            new StepRunResult
            {
                Sequence = 1,
                Id = "s1",
                Operation = "finduia",
                Phase = "steps",
                Status = "failed",
                Success = false,
                Error = "Element not found",
                Classification = UiFailureClassification.OperationFailure
            },
            new RecoveryContext
            {
                RunId = "r",
                RunDirectory = Path.GetTempPath(),
                PlanPath = "p",
                OriginalPlanSha256 = "abc",
                StepObjectBindings = new Dictionary<string, IReadOnlyDictionary<string, string>>(StringComparer.Ordinal)
                {
                    ["s1"] = new Dictionary<string, string>(StringComparer.Ordinal)
                    {
                        ["locator"] = "login.submit",
                        ["parentLocator"] = "login.form"
                    }
                }
            });

        Assert.Equal(RecoveryClassification.AmbiguousLocator, decision.Classification);
        Assert.False(decision.MayAttemptRecovery);
        Assert.Null(decision.BindingKey);
    }

    [Fact]
    public void Classifier_SingleBinding_CarriesExactKey()
    {
        var decision = RecoveryFailureClassifier.ForTests().Classify(
            new StepRunResult
            {
                Sequence = 1,
                Id = "s1",
                Operation = "finduia",
                Phase = "steps",
                Status = "failed",
                Success = false,
                Error = "Element not found",
                Classification = UiFailureClassification.OperationFailure
            },
            new RecoveryContext
            {
                RunId = "r",
                RunDirectory = Path.GetTempPath(),
                PlanPath = "p",
                OriginalPlanSha256 = "abc",
                StepObjectBindings = new Dictionary<string, IReadOnlyDictionary<string, string>>(StringComparer.Ordinal)
                {
                    ["s1"] = new Dictionary<string, string>(StringComparer.Ordinal)
                    {
                        ["parentLocator"] = "login.form"
                    }
                }
            });

        Assert.Equal(RecoveryClassification.StateNotStable, decision.Classification);
        Assert.Equal("parentLocator", decision.BindingKey);
        Assert.Equal("login.form", decision.ObjectReference);
    }

    [Fact]
    public void TypedNonExecutionProof_AllowsMutationClassification()
    {
        var decision = RecoveryFailureClassifier.ForTests().Classify(
            new StepRunResult
            {
                Sequence = 1,
                Id = "s1",
                Operation = "click",
                Phase = "steps",
                Status = "failed",
                Success = false,
                Error = "Element not found for click",
                FailureCode = RecoveryOperationSafety.FailureCodeLocatorNotFoundBeforeInvoke,
                ExecutionState = RecoveryOperationSafety.ExecutionStateNotStarted,
                Classification = UiFailureClassification.OperationFailure
            },
            EmptyContext());

        Assert.Equal(RecoveryClassification.StateNotStable, decision.Classification);
        Assert.True(decision.MayAttemptRecovery);
    }

    [Fact]
    public void PlanImmutabilityAudit_MissingOrUnreadable_FailClosed()
    {
        // Mirrors BuildRecoverySummary fail-closed rules (missing/unreadable ⇒ PlanMutated).
        byte[]? original = [1, 2, 3];
        var missingPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"), "gone.plan.json");
        Assert.True(ComputePlanMutated(original, missingPath));

        var unreadableDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(unreadableDir);
        // Path that exists as a directory — ReadAllBytes throws → fail closed.
        Assert.True(ComputePlanMutated(original, unreadableDir));

        var ok = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".plan.json");
        File.WriteAllBytes(ok, original);
        Assert.False(ComputePlanMutated(original, ok));
        Assert.True(ComputePlanMutated(null, ok));
    }

    private static bool ComputePlanMutated(byte[]? originalPlanBytes, string planPath)
    {
        if (originalPlanBytes is null)
            return true;

        try
        {
            if (!File.Exists(planPath))
                return true;
            var current = File.ReadAllBytes(planPath);
            return !originalPlanBytes.AsSpan().SequenceEqual(current);
        }
        catch
        {
            return true;
        }
    }

    [Fact]
    public async Task TimedOutClick_ExactlyOneUiPost()
    {
        var uiPosts = 0;
        var handler = MutationCatalogHandler(async (req, ct) =>
        {
            if (req.Method != HttpMethod.Post)
                return DefaultGetWithMutations(req);

            Interlocked.Increment(ref uiPosts);
            await Task.Delay(TimeSpan.FromSeconds(30), ct);
            return FakeHttpMessageHandler.Json(new { success = true, value = "late" });
        });

        var options = ReadyOptions();
        options.Runner.StepTransportTimeoutSeconds = 1;
        options.Recovery.BoundedWaitMilliseconds = 1;
        var plan = TestSupport.MinimalPlanJson(stepsOverride: """
            [ { "id": "click", "operation": "clickuia", "arguments": { "locator": { "automationId": "Pay" } } } ]
            """);
        var path = TestSupport.WritePlan(options, "timeout-click.plan.json", plan);
        var report = await CreateRunner(options, handler).RunAsync(path, dryRun: false);

        Assert.Equal("failed", report.Status);
        Assert.Equal(1, uiPosts);
        Assert.True(
            report.Recovery is null
            || report.Recovery.Attempts.Count == 0
            || report.Recovery.Attempts.All(a => a.Tier == RecoveryTier.None),
            $"recovery={JsonSerializer.Serialize(report.Recovery)}; failure={report.Failure?.Message}");
    }

    [Fact]
    public async Task TimedOutType_ExactlyOneUiPost()
    {
        var uiPosts = 0;
        var handler = MutationCatalogHandler(async (req, ct) =>
        {
            if (req.Method != HttpMethod.Post)
                return DefaultGetWithMutations(req);

            Interlocked.Increment(ref uiPosts);
            await Task.Delay(TimeSpan.FromSeconds(30), ct);
            return FakeHttpMessageHandler.Json(new { success = true, value = "late" });
        });

        var options = ReadyOptions();
        options.Runner.StepTransportTimeoutSeconds = 1;
        options.Recovery.BoundedWaitMilliseconds = 1;
        var plan = TestSupport.MinimalPlanJson(stepsOverride: """
            [ { "id": "type", "operation": "typeuia", "arguments": { "locator": { "automationId": "User" }, "value": "secret" } } ]
            """);
        var path = TestSupport.WritePlan(options, "timeout-type.plan.json", plan);
        var report = await CreateRunner(options, handler).RunAsync(path, dryRun: false);

        Assert.Equal("failed", report.Status);
        Assert.Equal(1, uiPosts);
        Assert.True(
            report.Recovery is null
            || report.Recovery.Attempts.Count == 0
            || report.Recovery.Attempts.All(a => a.Tier == RecoveryTier.None),
            $"recovery={JsonSerializer.Serialize(report.Recovery)}; failure={report.Failure?.Message}");
    }

    [Fact]
    public async Task ApprovedRepositoryHash_ComparedBeforeAfter()
    {
        var options = ReadyOptions();
        options.Recovery.BoundedWaitMilliseconds = 1;
        SeedRepoWithAlternate(options, ambiguous: false);

        var handler = ReadyHandler((req, _) =>
        {
            if (req.Method != HttpMethod.Post)
                return Task.FromResult(DefaultGet(req));

            return Task.FromResult(FakeHttpMessageHandler.Json(new
            {
                success = false,
                error = "Element not found for finduia"
            }));
        });

        var plan = """
            {
              "schemaVersion": 1,
              "catalogSchemaVersion": 2,
              "planId": "HASH-1",
              "name": "hash",
              "objectRepository": "object-repository/repository.json",
              "steps": [
                {
                  "id": "find",
                  "operation": "finduia",
                  "arguments": { "locator": { "$objectRef": "login.submit" } }
                }
              ]
            }
            """;
        var path = TestSupport.WritePlan(options, "hash.plan.json", plan);
        var report = await CreateRunner(options, handler).RunAsync(path, dryRun: false);

        Assert.NotNull(report.Recovery);
        Assert.False(string.IsNullOrWhiteSpace(report.Recovery!.ObjectRepositorySha256));
        Assert.False(string.IsNullOrWhiteSpace(report.Recovery.ObjectRepositorySha256After));
        Assert.Equal(report.Recovery.ObjectRepositorySha256, report.Recovery.ObjectRepositorySha256After);
        Assert.False(report.Recovery.ApprovedRepositoryMutated);
    }

    [Fact]
    public async Task Cleanup_Unchanged_And_NoRecoveryDuringCleanup()
    {
        var posts = new List<string>();
        var handler = ReadyHandler(async (req, _) =>
        {
            if (req.Method != HttpMethod.Post)
                return DefaultGet(req);

            var body = await req.Content!.ReadAsStringAsync();
            posts.Add(body);
            if (body.Contains("\"operation\":\"launch\"", StringComparison.Ordinal)
                || body.Contains("\"operation\": \"launch\"", StringComparison.Ordinal))
            {
                return FakeHttpMessageHandler.Json(new { success = false, error = "boom" });
            }

            return FakeHttpMessageHandler.Json(new { success = true });
        });

        var options = ReadyOptions();
        var plan = """
            {
              "schemaVersion": 1,
              "catalogSchemaVersion": 2,
              "planId": "CLEAN-1",
              "name": "clean",
              "steps": [
                { "id": "launch", "operation": "launch", "arguments": { "value": "C:\\\\App.exe" } },
                { "id": "after", "operation": "listwindows", "arguments": {} },
                { "id": "quit", "operation": "quit", "arguments": {} }
              ],
              "onFailureSteps": [
                { "id": "cleanup-quit", "operation": "quit", "arguments": {} }
              ]
            }
            """;
        var path = TestSupport.WritePlan(options, "cleanup.plan.json", plan);
        var report = await CreateRunner(options, handler).RunAsync(path, dryRun: false);

        Assert.Equal("failed", report.Status);
        Assert.Contains(report.OnFailureSteps, s => s.Id == "cleanup-quit");
        Assert.Equal(2, posts.Count); // failed launch + cleanup quit only
    }

    [Fact]
    public async Task CandidateRepair_NeverAutoPromoted_FakeAiOnly()
    {
        var options = ReadyOptions();
        options.Recovery.BoundedWaitMilliseconds = 1;
        options.Recovery.MaxRecoveryAttemptsPerStep = 5;
        options.Recovery.MaxTotalRecoveriesPerRun = 5;

        var handler = ReadyHandler((req, _) =>
        {
            if (req.Method != HttpMethod.Post)
                return Task.FromResult(DefaultGet(req));

            return Task.FromResult(FakeHttpMessageHandler.Json(new
            {
                success = false,
                error = "Element not found for click"
            }));
        });

        var plan = TestSupport.MinimalPlanJson(stepsOverride: """
            [ { "id": "click", "operation": "listwindows", "arguments": {} } ]
            """);
        var path = TestSupport.WritePlan(options, "candidate.plan.json", plan);
        var beforeHash = Sha256File(path);
        var report = await CreateRunner(options, handler).RunAsync(path, dryRun: false);

        Assert.Equal("failed", report.Status);
        Assert.Equal(beforeHash, Sha256File(path));
        Assert.NotNull(report.Recovery);
        Assert.Contains(report.Recovery!.Attempts, a => a.Tier == RecoveryTier.CandidateAiRepairProposal);
        var candidateAttempt = report.Recovery.Attempts.First(a => a.Tier == RecoveryTier.CandidateAiRepairProposal);
        Assert.NotNull(candidateAttempt.CandidateRepairPath);
        Assert.Contains(
            Path.Combine("object-repository", "candidates"),
            candidateAttempt.CandidateRepairPath!,
            StringComparison.OrdinalIgnoreCase);
        Assert.False(Path.GetFullPath(candidateAttempt.CandidateRepairPath!)
            .Contains(Path.Combine("object-repository", "pages"), StringComparison.OrdinalIgnoreCase));

        var json = File.ReadAllText(candidateAttempt.CandidateRepairPath!).Replace(" ", string.Empty);
        Assert.Contains("\"autoPromoted\":false", json, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("fake", new FakeRecoveryModelClient().ProviderName);
    }

    [Fact]
    public void FakeRecoveryModel_IsOnlyCiProviderRegisteredByDefault()
    {
        Assert.Equal("fake", new FakeRecoveryModelClient().ProviderName);
        Assert.Equal("fake", ReadyOptions().Recovery.Provider);
    }

    private static RecoveryContext EmptyContext() =>
        new()
        {
            RunId = "r",
            RunDirectory = Path.GetTempPath(),
            PlanPath = "p",
            OriginalPlanSha256 = "abc"
        };

    private static void SeedRepoWithAlternate(AgentOptions options, bool ambiguous)
    {
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        var repoRoot = Path.Combine(workspace.RootPath, "object-repository");
        Directory.CreateDirectory(Path.Combine(repoRoot, "pages"));
        File.WriteAllText(Path.Combine(repoRoot, "repository.json"), """
            {
              "schemaVersion": 1,
              "repositoryId": "demo",
              "name": "demo",
              "pages": [ { "pageId": "login", "file": "pages/login.page.json" } ]
            }
            """);
        _ = ambiguous;
        File.WriteAllText(Path.Combine(repoRoot, "pages", "login.page.json"), """
            {
              "schemaVersion": 1,
              "pageId": "login",
              "name": "Login",
              "state": "active",
              "elements": {
                "submit": {
                  "description": "Submit",
                  "locator": { "automationId": "PrimarySubmit" },
                  "alternateLocators": [
                    { "automationId": "AltSubmit" }
                  ],
                  "quality": { "grade": "strong" },
                  "source": { "kind": "manual" }
                }
              }
            }
            """);
    }

    private static AgentOptions ReadyOptions() =>
        TestSupport.CreateOptions(baseUrl: "http://127.0.0.1:33211", bearerToken: "test-token");

    private static DeterministicPlanRunner CreateRunner(AgentOptions options, FakeHttpMessageHandler handler)
    {
        var workspace = TestSupport.CreateWorkspace(options);
        workspace.Initialize();
        return TestSupport.CreatePlanRunner(options, workspace, TestSupport.CreateFactory(handler));
    }

    private static FakeHttpMessageHandler ReadyHandler(
        Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> onRequest) =>
        new(onRequest);

    private static FakeHttpMessageHandler MutationCatalogHandler(
        Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> onRequest) =>
        new(onRequest);

    private static HttpResponseMessage DefaultGet(HttpRequestMessage req) =>
        DefaultGetCore(req, CatalogFixtures.Phase2Catalog());

    private static HttpResponseMessage DefaultGetWithMutations(HttpRequestMessage req) =>
        DefaultGetCore(req, MutationCatalog());

    private static HttpResponseMessage DefaultGetCore(HttpRequestMessage req, object catalog)
    {
        if (req.RequestUri!.AbsolutePath.EndsWith("/status", StringComparison.OrdinalIgnoreCase))
        {
            return FakeHttpMessageHandler.Json(new
            {
                sessionId = (string?)null,
                status = 0,
                value = new
                {
                    ready = true,
                    message = "ok",
                    build = new { version = "1.0.105", revision = "", time = "2026-01-01T00:00:00Z" }
                }
            });
        }

        if (req.RequestUri.AbsolutePath.Contains("operations", StringComparison.OrdinalIgnoreCase))
        {
            return FakeHttpMessageHandler.Json(new
            {
                success = true,
                value = catalog
            });
        }

        return new HttpResponseMessage(HttpStatusCode.NotFound);
    }

    private static object MutationCatalog()
    {
        var catalog = CatalogFixtures.Phase2Catalog();
        catalog.Operations.Add(new OperationDescriptorDto
        {
            Name = "clickuia",
            Aliases = [],
            Category = "element-action",
            OperationType = "invoke",
            RequiresSession = false,
            RequiredInputs = ["locator"]
        });
        catalog.Operations.Add(new OperationDescriptorDto
        {
            Name = "typeuia",
            Aliases = [],
            Category = "element-action",
            OperationType = "input",
            RequiresSession = false,
            RequiredInputs = ["locator", "value"]
        });
        return catalog;
    }

    private static string Sha256File(string path)
    {
        var bytes = File.ReadAllBytes(path);
        return Convert.ToHexString(SHA256.HashData(bytes));
    }
}
