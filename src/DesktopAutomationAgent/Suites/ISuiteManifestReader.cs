namespace DesktopAutomationAgent.Suites;

public interface ISuiteManifestReader
{
    SuiteValidationResult ValidateFile(string path);

    KeyValidationResult ValidateKeys(IEnumerable<string> keys);
}
