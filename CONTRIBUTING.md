# Contributing

Thanks for helping improve Desktop Automation Helper.

## Before you start

- Search existing issues and pull requests before opening a duplicate.
- Use an issue to discuss substantial behavior or API changes first.
- Never post bearer tokens, credentials, customer data, or private automation
  artifacts. Report vulnerabilities according to [SECURITY.md](SECURITY.md).
- Keep changes focused. Separate unrelated fixes into separate pull requests.

## Development setup

This project targets .NET 8 and Windows desktop UI Automation.

```powershell
dotnet restore DesktopAutomationHelper.slnx
dotnet build DesktopAutomationHelper.slnx --configuration Release --no-restore
dotnet test src/DesktopAutomationDriver.Tests/DesktopAutomationDriver.Tests.csproj --configuration Release --no-build
dotnet test src/DesktopAutomationAgent.Tests/DesktopAutomationAgent.Tests.csproj --configuration Release --no-build
```

Windows is the authoritative test environment for driver and UI Automation
behavior. Add or update tests for observable behavior changes.

## Pull requests

1. Create a branch from the latest `main`.
2. Follow the existing C# style and avoid unrelated formatting changes.
3. Update documentation and examples when public behavior changes.
4. Run the relevant build and test commands.
5. Complete the pull-request template and link related issues.

By contributing, you agree that your contribution is licensed under this
repository's [MIT License](LICENSE) and that project interactions follow the
[Code of Conduct](CODE_OF_CONDUCT.md).
