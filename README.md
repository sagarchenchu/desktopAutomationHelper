# Desktop Automation Driver

A Windows Desktop Application Automation Driver built on top of [FlaUI](https://github.com/FlaUI/FlaUI) (UIA2 and UIA3). It exposes a WebDriver-compatible HTTP REST API, similar to Winium and pywinauto, letting you automate any Windows desktop application from any language.

Designed for **shared Citrix / RDS / VDI machines** where multiple users run simultaneously: each Windows account gets a deterministic, collision-free port and a unique Bearer token.

## Features

- **UIA2 and UIA3 support** - choose the UI Automation backend that best fits your application.
- **Standalone EXE** - publish as a single self-contained `.exe`; no additional runtime required.
- **Per-user port isolation** - port is derived from the Windows login name (FNV-1a hash → 30000–39999); no two users on the same machine will collide.
- **Bearer token authentication** - a random token is generated at startup and required on every call (except `GET /verify`).
- **`/verify` bootstrap endpoint** - unauthenticated; returns the token and port so clients always know how to connect.
- **Probe discovery port 9102** - the driver also binds a fixed well-known port for `/verify`; first user to start the driver claims it; subsequent users still work via their own port.
- **WebDriver-compatible REST API** - sessions, element finding, actions, and screenshots.
- **Multiple element location strategies** - `automation id`, `name`, `class name`, `tag name`, `xpath`, and more.
- **Rich actions** - click, double-click, right-click, send keys (WebDriver key codes), clear, get text, get attribute, take screenshot.

---

## Getting Started

### Running the Driver

Launch `DesktopAutomationDriver.exe`. On startup it prints a full connection banner:

```
╔══════════════════════════════════════════════════════════════╗
║          Desktop Automation Driver — Started                 ║
╠══════════════════════════════════════════════════════════════╣
║  Windows user  : DOMAIN\alice                                ║
║  Main port     : 33201                                       ║
║  Bearer token  : 7kP2mXqR...base64...                       ║
╠══════════════════════════════════════════════════════════════╣
║  ENDPOINTS  (require: Authorization: Bearer <token>)         ║
║                                                              ║
║  GET  http://localhost:33201/verify    <- no auth, returns token
║  ...                                                         ║
╠══════════════════════════════════════════════════════════════╣
║  PROBE (no auth)                                             ║
║  http://localhost:9102/verify  (active)                      ║
╚══════════════════════════════════════════════════════════════╝
  Authorization header to use:
    Authorization: Bearer 7kP2mXqR...
```

### Discovering the token at runtime

Use either URL to retrieve the connection info without any prior knowledge:

```http
GET http://localhost:9102/verify
GET http://localhost:<user-port>/verify
```

Response:

```json
{
  "status": 0,
  "value": {
    "running": true,
    "username": "DOMAIN\\alice",
    "port": 33201,
    "probePort": 9102,
    "token": "7kP2mXqR...",
    "authorizationHeader": "Bearer 7kP2mXqR..."
  }
}
```

### Authentication

Every request (except `GET /verify`) must include:

```http
Authorization: Bearer <token>
```

Missing or invalid tokens receive `HTTP 401`:

```json
{ "status": 401, "value": { "error": "unauthorized", "message": "..." } }
```

---

## REST API Reference

All responses follow the WebDriver response envelope:

```json
{ "sessionId": "<id>", "status": 0, "value": { ... } }
```

### GET /verify  *(no auth)*

Returns driver status, connection port, and the Bearer token. Use this to bootstrap
a client session on any shared machine.

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

Full API reference: see [Controllers/ElementController.cs](src/DesktopAutomationDriver/Controllers/ElementController.cs).

---

## Example: Automating Notepad (Python)

```python
import requests

PROBE = "http://localhost:9102"   # fixed well-known probe port

# 1. Discover the token and user-specific port (no auth needed)
info = requests.get(f"{PROBE}/verify").json()["value"]
BASE  = f"http://localhost:{info['port']}"
AUTH  = {"Authorization": info["authorizationHeader"]}

# 2. Create session — all calls from here need the Authorization header
session = requests.post(f"{BASE}/session", headers=AUTH, json={
    "desiredCapabilities": {
        "app": r"C:\Windows\System32\notepad.exe",
        "uiaType": "UIA3"
    }
}).json()
session_id = session["sessionId"]

# 3. Find the text area
el = requests.post(f"{BASE}/session/{session_id}/element", headers=AUTH,
                   json={"using": "class name", "value": "Edit"}).json()
el_id = el["value"]["elementId"]

# 4. Type text
requests.post(f"{BASE}/session/{session_id}/element/{el_id}/value", headers=AUTH,
              json={"value": ["Hello from Desktop Automation Driver!"]})

# 5. Clean up
requests.delete(f"{BASE}/session/{session_id}", headers=AUTH)
```

> **Note:** If the probe port (9102) is taken by another user's driver, query the
> user-specific port directly. The port for a given Windows account is always
> `30000 + fnv1a32(username.lower()) % 10000`.

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
|  +-VerifyController.cs         GET /verify  (no auth — bootstrap endpoint)
|  +-StatusController.cs         GET /status
|  +-SessionController.cs        POST/GET/DELETE /session
|  +-ElementController.cs        All /session/{id}/element/* endpoints
+-Middleware/
|  +-BearerTokenMiddleware.cs    Validates Authorization: Bearer <token>
+-Models/
|  +-Request/                    Request DTOs
|  +-Response/                   Response DTOs + WebDriverResponse<T> + VerifyResponse
+-Services/
|  +-IDriverContext.cs           Per-user context interface (port, token, username)
|  +-DriverContext.cs            FNV-1a port derivation + random token at startup
|  +-ISessionManager.cs          Session lifecycle interface
|  +-SessionManager.cs           Thread-safe session store
|  +-IAutomationService.cs       Element operations interface
|  +-AutomationService.cs        FlaUI-backed element finding and actions
|  +-AutomationSession.cs        Per-session state and element cache
+-Program.cs                     Host setup, probe-server (port 9102), startup banner
```

## License

MIT
