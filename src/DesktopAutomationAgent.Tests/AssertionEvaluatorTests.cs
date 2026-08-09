using System.Text.Json;
using DesktopAutomationAgent.Execution;
using DesktopAutomationAgent.Plans;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Tests;

public class AssertionEvaluatorTests
{
    private readonly AssertionEvaluator _evaluator =
        new(Options.Create(TestSupport.CreateOptions()));

    [Fact]
    public void StringEquality()
    {
        Assert.True(Eval("\"Dashboard\"", "", "equals", "\"Dashboard\"")[0].Passed);
    }

    [Fact]
    public void CaseInsensitiveStringEquality()
    {
        var assertion = new PlanAssertion
        {
            Path = "",
            Operator = "equals",
            Expected = JsonSerializer.SerializeToElement("dashboard"),
            IgnoreCase = true
        };
        var actual = JsonSerializer.SerializeToElement("Dashboard");
        Assert.True(_evaluator.Evaluate(actual, [assertion])[0].Passed);
    }

    [Fact]
    public void NumericEquality()
    {
        Assert.True(Eval("10", "", "equals", "10")[0].Passed);
        Assert.True(Eval("10.0", "", "equals", "10")[0].Passed);
    }

    [Fact]
    public void ObjectEqualityIndependentOfPropertyOrder()
    {
        Assert.True(Eval("{\"b\":2,\"a\":1}", "", "equals", "{\"a\":1,\"b\":2}")[0].Passed);
    }

    [Fact]
    public void OrderedArrayEquality()
    {
        Assert.True(Eval("[1,2]", "", "equals", "[1,2]")[0].Passed);
        Assert.False(Eval("[2,1]", "", "equals", "[1,2]")[0].Passed);
    }

    [Fact]
    public void StringContains()
    {
        Assert.True(Eval("\"Account Summary\"", "", "contains", "\"Account\"")[0].Passed);
    }

    [Fact]
    public void ArrayContains()
    {
        Assert.True(Eval("[\"a\",\"b\"]", "", "contains", "\"b\"")[0].Passed);
    }

    [Fact]
    public void RegexSuccessAndMismatch()
    {
        Assert.True(Eval("\"abc123\"", "", "matchesRegex", "\"^[a-z]+\\\\d+$\"")[0].Passed);
        Assert.False(Eval("\"abc\"", "", "matchesRegex", "\"^\\\\d+$\"")[0].Passed);
    }

    [Fact]
    public void RegexTimeoutFailsPredictably()
    {
        var options = TestSupport.CreateOptions();
        options.Runner.RegexTimeoutMilliseconds = 1;
        var evaluator = new AssertionEvaluator(Options.Create(options));
        var actual = JsonSerializer.SerializeToElement(new string('a', 20000) + "b");
        var assertion = new PlanAssertion
        {
            Path = "",
            Operator = "matchesRegex",
            Expected = JsonSerializer.SerializeToElement("(a+)+$")
        };
        Assert.False(evaluator.Evaluate(actual, [assertion])[0].Passed);
    }

    [Fact]
    public void BooleanAndNullAssertions()
    {
        Assert.True(Eval("true", "", "isTrue", null)[0].Passed);
        Assert.True(Eval("false", "", "isFalse", null)[0].Passed);
        Assert.True(Eval("null", "", "isNull", null)[0].Passed);
        Assert.True(Eval("1", "", "isNotNull", null)[0].Passed);
    }

    [Fact]
    public void JsonPointerObjectAndArrayTraversal()
    {
        var json = """{"items":[{"name":"Account"}]}""";
        Assert.True(Eval(json, "/items/0/name", "equals", "\"Account\"")[0].Passed);
    }

    [Fact]
    public void MissingPathFails()
    {
        var result = Eval("{\"a\":1}", "/missing", "isNotNull", null)[0];
        Assert.False(result.Passed);
        Assert.Contains("not found", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void WrongActualTypeFailsContains()
    {
        Assert.False(Eval("123", "", "contains", "\"1\"")[0].Passed);
    }

    private IReadOnlyList<AssertionResult> Eval(
        string actualJson,
        string path,
        string op,
        string? expectedJson)
    {
        var actual = JsonSerializer.Deserialize<JsonElement>(actualJson);
        var assertion = new PlanAssertion
        {
            Path = path,
            Operator = op,
            Expected = expectedJson is null ? null : JsonSerializer.Deserialize<JsonElement>(expectedJson)
        };
        return _evaluator.Evaluate(actual, [assertion]);
    }
}
