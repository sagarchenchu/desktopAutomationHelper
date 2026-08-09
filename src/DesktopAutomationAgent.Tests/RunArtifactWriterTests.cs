using System.Text.Json;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Execution;

namespace DesktopAutomationAgent.Tests;

public class RunArtifactWriterTests
{
    [Fact]
    public void WritesAtomicallyAndNeverOverwrites()
    {
        var dir = Path.Combine(Path.GetTempPath(), "da-agent-runs-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var writer = new RunArtifactWriter();
        var report = SampleReport();

        writer.WriteRunReport(dir, report);
        Assert.True(File.Exists(Path.Combine(dir, "run.json")));
        Assert.Equal("written", report.ArtifactWriteStatus);

        var ex = Assert.ThrowsAny<Exception>(() => writer.WriteRunReport(dir, report));
        Assert.Contains("must not be overwritten", ex.Message, StringComparison.OrdinalIgnoreCase);

        Directory.Delete(dir, recursive: true);
    }

    [Fact]
    public void RedactsSensitivePropertiesFromReport()
    {
        var dir = Path.Combine(Path.GetTempPath(), "da-agent-runs-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var report = SampleReport();
        report = new RunReport
        {
            ReportSchemaVersion = 1,
            RunId = report.RunId,
            Status = "passed",
            ExitCode = 0,
            PlanPath = "plans/a.plan.json",
            PlanId = "A",
            PlanName = "A",
            PlanSha256 = "abc",
            StartedAtUtc = DateTimeOffset.UtcNow,
            FinishedAtUtc = DateTimeOffset.UtcNow,
            Steps =
            [
                new StepRunResult
                {
                    Sequence = 1,
                    Id = "s1",
                    Operation = "listwindows",
                    Phase = "steps",
                    Status = "passed",
                    Success = true,
                    CaptureResponse = true,
                    Arguments = new Dictionary<string, JsonElement>
                    {
                        ["connectionString"] = JsonSerializer.SerializeToElement("Server=.;Password=x")
                    },
                    ResponseValue = JsonSerializer.SerializeToElement(new { token = "abc", title = "ok" })
                }
            ],
            Failure = new RunFailure
            {
                Classification = UiFailureClassification.OperationFailure,
                Message = "Bearer secret-token failed"
            }
        };

        new RunArtifactWriter().WriteRunReport(dir, report);
        var json = File.ReadAllText(Path.Combine(dir, "run.json"));
        Assert.Contains("[REDACTED]", json);
        Assert.DoesNotContain("Server=.;Password=x", json);
        Assert.DoesNotContain("secret-token", json);

        Directory.Delete(dir, recursive: true);
    }

    private static RunReport SampleReport() =>
        new()
        {
            ReportSchemaVersion = 1,
            RunId = "20260101T000000000Z-deadbeef",
            Status = "passed",
            ExitCode = 0,
            PlanPath = "plans/a.plan.json",
            PlanId = "A",
            PlanName = "A",
            PlanSha256 = "abc",
            StartedAtUtc = DateTimeOffset.UtcNow,
            FinishedAtUtc = DateTimeOffset.UtcNow
        };
}
