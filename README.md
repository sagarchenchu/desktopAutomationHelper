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

## POST /ui — Unified Automation Endpoint

All desktop automation operations are exposed through a single endpoint:

```
POST http://127.0.0.1:{user-port}/ui
Authorization: Bearer <token>
Content-Type: application/json
```

**Request envelope:**

```json
{
  "operation": "<operation-name>",
  "locator":  { "automationId": "...", "name": "...", "className": "...", "controlType": "...", "xpath": "..." },
  "locator2": { "automationId": "...", "name": "...", "className": "...", "controlType": "...", "xpath": "..." },
  "value": "...",
  "index": 0,
  "columnIndex": 0
}
```

All fields except `operation` are optional and depend on the operation being used.

**Success response:**

```json
{ "status": 0, "value": { ... } }
```

**Error response:**

```json
{ "status": 400, "value": { "error": "...", "message": "..." } }
```

---

### Session & Window Management

| Operation | Required fields | Description | Response `value` |
|---|---|---|---|
| `launch` | `value` (exe path) | Launch an application and start a session | `{ "sessionId": "...", "app": "..." }` |
| `close` / `quit` | — | Close the active application and session | `null` |
| `closewindow` | `value` (partial title) | Close a window by partial title match | `null` |
| `maximize` | — | Maximize the active window | `null` |
| `minimize` | — | Minimize the active window | `null` |
| `switchwindow` | `value` (partial title) | Focus a window by partial title match | `null` |
| `refresh` | — | Re-attach to the main window of the active application | `null` |
| `screenshot` | `value` (optional file path) | Take a screenshot; returns Base64 PNG or saves to file | `{ "screenshot": "<base64>" }` |
| `listelements` | `locator` | List all matching elements and their properties | `{ "elements": [ ... ] }` |
| `listwindows` | — | List all top-level windows visible on the desktop | `{ "windows": [ ... ] }` |

**Example — launch:**

```json
{ "operation": "launch", "value": "C:\\Windows\\System32\\notepad.exe" }
```

---

### Element Query

| Operation | Required fields | Description | Response `value` |
|---|---|---|---|
| `exists` | `locator` | Check if an element exists (no wait) | `{ "exists": true/false }` |
| `waitfor` | `locator`, `value` (timeout seconds) | Wait up to N seconds for an element to appear | `{ "found": true/false }` |
| `isenabled` | `locator` | Check whether the element is enabled | `{ "isEnabled": true/false }` |
| `isvisible` | `locator` | Check whether the element is visible | `{ "isVisible": true/false }` |
| `isclickable` | `locator` | Check whether the element is enabled and visible | `{ "isClickable": true/false }` |
| `ischecked` | `locator` | Check the toggle/check state of an element | `{ "isChecked": true/false }` |
| `getvalue` | `locator` | Get the Value pattern value (e.g. text in an Edit) | `{ "value": "..." }` |
| `gettext` | `locator` | Get the visible text of an element | `{ "text": "..." }` |
| `getname` | `locator` | Get the UIA Name property | `{ "name": "..." }` |
| `getcontroltype` | `locator` | Get the UIA ControlType string | `{ "controlType": "Button" }` |
| `getselected` | `locator` | Get the currently selected item from a combo or list | `{ "selected": "..." }` |
| `gettable` / `gettabledata` | `locator` | Read all rows and headers from a grid/table | `{ "headers": [...], "rows": [[...], ...] }` |
| `gettableheaders` | `locator` | Read only the column headers from a grid/table | `{ "headers": [...] }` |

---

### Position Comparison

These operations require **both** `locator` and `locator2`.

| Operation | Description | Response `value` |
|---|---|---|
| `isrightof` | Whether element 1 is to the right of element 2 | `{ "isRightOf": true/false }` |
| `isleftof` | Whether element 1 is to the left of element 2 | `{ "isLeftOf": true/false }` |
| `isabove` | Whether element 1 is above element 2 | `{ "isAbove": true/false }` |
| `isbelow` | Whether element 1 is below element 2 | `{ "isBelow": true/false }` |
| `getposition` | Bounding rectangles + all four relative positions | `{ "element1": {...}, "element2": {...}, "isRightOf": ..., ... }` |

---

### Element Actions

| Operation | Required fields | Description | Response `value` |
|---|---|---|---|
| `click` | `locator` | Click the element (Invoke pattern or mouse click) | `null` |
| `doubleclick` | `locator` | Double-click the element | `null` |
| `rightclick` | `locator` | Right-click the element | `null` |
| `hover` | `locator` | Move the mouse pointer over the element | `null` |
| `focus` | `locator` | Give keyboard focus to the element | `null` |
| `type` | `locator`, `value` | Type text into the element | `null` |
| `clear` | `locator` | Clear the text content of an Edit element | `null` |
| `sendkeys` | `locator`, `value` | Send key sequences, e.g. `{ENTER}`, `{TAB}`, `{F5}` | `null` |
| `scroll` | `locator` | Scroll the element into view | `null` |
| `check` | `locator` | Set a toggle/checkbox element to the checked state | `null` |
| `uncheck` | `locator` | Set a toggle/checkbox element to the unchecked state | `null` |
| `select` | `locator`, `value` or `index` | Select a ComboBox item by name or zero-based index | `null` |
| `selectaid` | `locator`, `value` (AutomationId) | Select a ComboBox item by AutomationId | `null` |
| `typeandselect` | `locator`, `value` | Type a filter string and select the first matching dropdown item | `null` |
| `clickgridcell` | `locator`, `index` (row), `columnIndex` (col) | Click a grid cell at the given row/column | `null` |
| `doubleclickgridcell` | `locator`, `index` (row), `columnIndex` (col) | Double-click a grid cell at the given row/column | `null` |
| `draganddrop` | `locator` (source), `locator2` (target) | Drag the source element and drop it onto the target element | `null` |

#### draganddrop — request example

```json
{
  "operation": "draganddrop",
  "locator":  { "automationId": "sourceId",  "controlType": "ListItem" },
  "locator2": { "automationId": "targetId",  "controlType": "List" }
}
```

#### sendkeys — supported key tokens

Key tokens are wrapped in braces inside the `value` string, e.g. `"Hello{ENTER}"`.

| Token | Key |
|---|---|
| `{ENTER}` / `{RETURN}` | Enter / Return |
| `{TAB}` | Tab |
| `{ESC}` / `{ESCAPE}` | Escape |
| `{BACKSPACE}` / `{BS}` | Backspace |
| `{DELETE}` / `{DEL}` | Delete |
| `{INSERT}` / `{INS}` | Insert |
| `{HOME}` | Home |
| `{END}` | End |
| `{UP}` | Arrow Up |
| `{DOWN}` | Arrow Down |
| `{LEFT}` | Arrow Left |
| `{RIGHT}` | Arrow Right |
| `{PGUP}` / `{PAGEUP}` | Page Up |
| `{PGDN}` / `{PAGEDOWN}` | Page Down |
| `{SPACE}` | Space |
| `{F1}`–`{F12}` | Function keys |

---

### Alert / Dialog Handling

| Operation | Required fields | Description | Response `value` |
|---|---|---|---|
| `alertok` | — | Click the OK / Yes / Save button of the topmost modal dialog | `{ "success": true }` |
| `alertcancel` | — | Click the Cancel / No button of the topmost modal dialog | `{ "success": true }` |
| `alertclose` | — | Close the topmost modal dialog via the Window pattern | `{ "success": true }` |

---

### Locator reference

All `locator` and `locator2` objects support the following properties (combined with AND logic):

| Property | Description |
|---|---|
| `automationId` | UIA AutomationId property |
| `name` | UIA Name property (exact match) |
| `className` | UIA ClassName property |
| `controlType` | UIA control type string, e.g. `Button`, `Edit`, `ComboBox` |
| `xpath` | XPath-style expression (overrides all other properties when set) |

**XPath syntax examples:**

```
//Button[@Name='OK']
//*[@AutomationId='myId']
//ComboBox[@Name='Status']/ListItem[@Name='Active']
//Edit[@AutomationId='search' and @Name='Search']
//ListItem[@Name='Item'][2]
```

---

## Recording Metadata

When `POST /record/start` launches an application, the response now includes the captured
screen resolution plus the launched window's initial `x`, `y`, `width`, `height`,
`windowState`, and `isFullScreen` details. The exported recording JSON also includes the
same `screen` / `launch` metadata and a per-action `pointerContext` block for assistive
and right-click-driven actions so hook coordinates, live cursor coordinates, and the
resolved fallback point can be inspected later.

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
