# Desktop Automation Driver

A Windows Desktop Application Automation Driver built on top of [FlaUI](https://github.com/FlaUI/FlaUI) (UIA2 and UIA3). It exposes a WebDriver-compatible HTTP REST API, similar to Winium and pywinauto, letting you automate any Windows desktop application from any language.

## Features

- **UIA2 and UIA3 support** - choose the UI Automation backend that best fits your application.
- **Standalone EXE** - publish as a single self-contained `.exe`; no additional runtime required.
- **WebDriver-compatible REST API** - sessions, element finding, actions, and screenshots.
- **Multiple element location strategies** - `automation id`, `name`, `class name`, `tag name`, `xpath`, and more.
- **Rich actions** - click, double-click, right-click, send keys (with WebDriver special key codes), clear, get text, get attribute, take screenshot.

---

## Getting Started

### Running the Driver

Launch the `DesktopAutomationDriver.exe`. It listens on **port 4723** by default.

```cmd
DesktopAutomationDriver.exe
```

To use a custom port:

```cmd
DesktopAutomationDriver.exe --urls http://0.0.0.0:9515
```

---

## REST API Reference

All responses follow the WebDriver response envelope:

```json
{ "sessionId": "<id>", "status": 0, "value": { ... } }
```

### POST /session

Creates a new automation session (launch app or attach to running process).

**Launch an application:**

```json
{
  "desiredCapabilities": {
    "app": "C:\\Windows\\System32\\notepad.exe",
    "uiaType": "UIA3",
    "launchDelay": 1000
  }
}
```

**Attach to running process:**

```json
{
  "desiredCapabilities": {
    "appName": "notepad",
    "uiaType": "UIA3"
  }
}
```

### DELETE /session/{sessionId}

Closes the session and disposes the automation backend.

### POST /session/{sessionId}/element

Finds the first matching element.

Supported `using` strategies:

| Strategy | Description |
|---|---|
| `automation id` / `id` | Match by AutomationId |
| `name` / `link text` | Exact match on Name |
| `partial link text` | Substring match on Name |
| `class name` | Match by ClassName |
| `tag name` | Match by ControlType (e.g. `Button`, `Edit`) |
| `xpath` | Pattern: `//ControlType[@AutomationId='x']` |

### POST /session/{sessionId}/element/{elementId}/click

Clicks the element.

### POST /session/{sessionId}/element/{elementId}/value

Types text into the element. Supports WebDriver key codes.

```json
{ "value": ["Hello World", "\uE007"] }
```

### GET /session/{sessionId}/element/{elementId}/text

Returns the visible text of the element.

### GET /session/{sessionId}/element/{elementId}/attribute/{name}

Returns the named attribute (`Name`, `AutomationId`, `ClassName`, `Value`, etc.).

### GET /session/{sessionId}/screenshot

Returns a Base64-encoded PNG screenshot of the application window.

Full API reference: see [Controllers/ElementController.cs](src/DesktopAutomationDriver/Controllers/ElementController.cs) and [Controllers/SessionController.cs](src/DesktopAutomationDriver/Controllers/SessionController.cs).

---

## Example: Automating Notepad (Python)

```python
import requests

BASE = "http://localhost:4723"

session = requests.post(f"{BASE}/session", json={
    "desiredCapabilities": {
        "app": r"C:\Windows\System32\notepad.exe",
        "uiaType": "UIA3"
    }
}).json()
session_id = session["sessionId"]

el = requests.post(f"{BASE}/session/{session_id}/element",
                   json={"using": "class name", "value": "Edit"}).json()
el_id = el["value"]["elementId"]

requests.post(f"{BASE}/session/{session_id}/element/{el_id}/value",
              json={"value": ["Hello from Desktop Automation Driver!"]})

requests.delete(f"{BASE}/session/{session_id}")
```

---

## Building from Source

> **Note:** Building and running requires **Windows** (UIA2/UIA3 are Windows-only APIs).

```cmd
dotnet build DesktopAutomationHelper.slnx --configuration Release
dotnet test src/DesktopAutomationDriver.Tests --configuration Release
```

### Publish standalone EXE

```cmd
dotnet publish src/DesktopAutomationDriver ^
  --configuration Release ^
  --runtime win-x64 ^
  --self-contained true ^
  -p:PublishSingleFile=true ^
  --output ./publish
```

---

## Architecture

```
DesktopAutomationDriver
+-Controllers/
|  +-StatusController.cs        GET /status
|  +-SessionController.cs       POST/GET/DELETE /session
|  +-ElementController.cs       All /session/{id}/element/* endpoints
+-Models/
|  +-Request/                   Request DTOs
|  +-Response/                  Response DTOs + WebDriverResponse<T>
+-Services/
|  +-ISessionManager.cs         Session lifecycle interface
|  +-SessionManager.cs          Thread-safe session store
|  +-IAutomationService.cs      Element operations interface
|  +-AutomationService.cs       FlaUI-backed element finding and actions
|  +-AutomationSession.cs       Per-session state and element cache
+-Program.cs                    ASP.NET Core setup, port configuration
```

## License

MIT
