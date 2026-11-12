# Desktop Automation Test App

WinForms harness used to exercise `DesktopAutomationDriver` actions against a
known, stable UI surface before Gate 0 Assistive acceptance and broader agentic
work.

Project: `src/DesktopAutomationTestApp`

## Why it exists

Driver unit tests do not replace a real desktop UI. This app provides:

- Stable `AutomationId` values (`Control.Name`) listed in `UiIds`
- Click / double-click / right-click targets
- Type and clear
- Drop-down select and type-ahead combo
- Radio buttons and checkbox
- DataGridView with column headers
- Child window, second top-level window, and modal dialog for `switchwindow` /
  dialog Assistive scenarios
- An event log control for quick human or automated observation

It does **not** reference the driver or agent projects.

## Build / run (Windows)

```cmd
dotnet build src/DesktopAutomationTestApp/DesktopAutomationTestApp.csproj -c Release
dotnet run --project src/DesktopAutomationTestApp -c Release
```

Publish:

```cmd
dotnet publish src/DesktopAutomationTestApp/DesktopAutomationTestApp.csproj ^
  -c Release -r win-x64 --self-contained false -o ./publish/test-app
```

## Suggested driver smoke matrix

Launch the test app, then launch the driver. Create a session against the test
app process/window and try:

| Capability | Control AutomationId | Example operation |
|---|---|---|
| Click | `ClickButton` | `click` / `clickuia` |
| Double-click | `DoubleClickButton` | `doubleclick` / `doubleclickuia` |
| Right-click / context menu | `RightClickButton` | Assistive right-click → Copy |
| Type | `NameEdit` | `type` / `typeuia` |
| Clear | `ClearNameButton` + `NameEdit` | `click` then assert empty |
| Drop-down select | `CountryCombo` | `select` / combo UIA ops |
| Type-ahead | `TypeAheadCombo` | type `Bos` → `Boston` |
| Radio | `RoleRadioAdmin` | `click` / select |
| Checkbox | `AgreeCheckBox` | `check` / `uncheck` |
| Grid headers | `SampleGrid` | `gettableheaders` |
| Grid cell | `SampleGrid` | `clickgridcell` / `gettable` |
| Child window | `OpenChildWindowButton` → `ChildWindow` | `switchwindow` |
| Second window | `OpenSecondWindowButton` → `SecondWindow` | `switchwindow` |
| Dialog OK/Cancel | `OpenDialogButton` → `DialogOkButton` | Assistive dialog scenarios |
| Menu path | `File` → `Open Child Window` | `clicklogicalmenupath` |

Observe outcomes in `EventLog` and `StatusLabel`.

## Gate 0 relationship

This harness is a good target application for
`docs/assistive-windows-acceptance-results.md` once you have an interactive
Windows desktop (local or Vultr RDP). It is not a substitute for that checklist.
