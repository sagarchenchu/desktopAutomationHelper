# Desktop Automation Driver

A Windows Desktop Application Automation Driver built on top of [FlaUI](https://github.com/FlaUI/FlaUI) (UIA3). It exposes a REST API that lets you automate any Windows desktop application from any language or test framework.

Designed for **shared Citrix / RDS / VDI machines** where multiple users run simultaneously: each Windows account gets a deterministic, collision-free port and a unique Bearer token.

## Features

- **UIA3 support** via FlaUI — automate any Win32, WPF, WinForms, or UWP application.
- **Standalone EXE** — publish as a single self-contained `.exe`; no additional runtime required.
- **Per-user port isolation** — port is derived from the Windows login name (FNV-1a hash → 30000–39999); no two users on the same machine will collide.
- **Bearer token authentication** — a cryptographically random token is generated at startup and required on every call (except `GET /verify`).
- **`/verify` bootstrap endpoint** — unauthenticated; returns the token and port so clients always know how to connect.
- **Probe discovery port 9102** — a fixed well-known port for `/verify`; first user to start the driver claims it; subsequent users still work via their own port.
- **Unified `POST /ui` endpoint** — 40+ automation operations (launch, click, type, select, drag-and-drop, table reading, and more) through one request envelope.
- **Simple convenience endpoints** — purpose-built thin wrappers (`/launch`, `/click/name`, `/select/combobox/name`, etc.) for simple scripting.
- **WebDriver session API** — WebDriver-compatible `/session` endpoints for tools that expect that protocol.
- **Action recording** — built-in overlay records clicks and keyboard input as a JSON sequence that can be replayed via the API.
- **Built-in retry** — element-finding operations automatically retry for up to 5 seconds so callers rarely need explicit waits.
- **Failure screenshots** — on any error, a screenshot is automatically captured and its path is returned in the error response.

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
║  UNIFIED ENDPOINT (require: Authorization: Bearer <token>)   ║
║                                                              ║
║  POST http://127.0.0.1:33201/ui  <- all operations           ║
║  ...                                                         ║
╠══════════════════════════════════════════════════════════════╣
║  PROBE (no auth)                                             ║
║  http://localhost:9102/verify  (active)                      ║
╚══════════════════════════════════════════════════════════════╝
  Authorization header to use:
    Authorization: Bearer 7kP2mXqR...
```

### Discovering the token at runtime

Use either URL to retrieve connection details without any prior knowledge:

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

> **Note:** If the probe port (9102) is already claimed by another user's driver instance,
> `probePort` will be `null` in the response. Compute the user-specific port directly:
> `30000 + fnv1a32(username.toLowerCase()) % 10000`.

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

## REST API Overview

The driver exposes four groups of endpoints:

| Group | Base path | Purpose |
|---|---|---|
| Bootstrap | `GET /verify`, `GET /status` | Discover connection details and check readiness |
| Unified | `POST /ui` | All 40+ automation operations through one envelope |
| Simple | `/launch`, `/click/*`, `/select/*`, `/alert/*`, … | Thin, purpose-built wrappers for common operations |
| WebDriver session | `/session`, `/session/{id}/element/*` | WebDriver-compatible session + element protocol |
| Recording | `/record/*` | Start, monitor and stop action recording |

All automation endpoints bind **only** to `127.0.0.1` (IPv4 loopback) so the driver is not reachable from outside the machine.

---

## Bootstrap Endpoints

### GET /verify  *(no auth required)*

Returns driver status, the user-specific port, and the Bearer token needed for all other calls.

**Response:**

```json
{
  "status": 0,
  "value": {
    "running": true,
    "username": "DOMAIN\\alice",
    "port": 33201,
    "probePort": 9102,
    "token": "<base64-token>",
    "authorizationHeader": "Bearer <base64-token>"
  }
}
```

### GET /status

Returns driver readiness and version information.

**Response:**

```json
{
  "status": 0,
  "value": {
    "ready": true,
    "message": "Desktop Automation Driver is running",
    "build": {
      "version": "1.0.0",
      "time": "2024-01-01T00:00:00+00:00"
    }
  }
}
```

---

## POST /ui — Unified Automation Endpoint

The primary entry point for all desktop automation. All operations share the same request/response envelope.

```
POST http://127.0.0.1:{user-port}/ui
Authorization: Bearer <token>
Content-Type: application/json
```

### Request envelope

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

All fields except `operation` are optional and their meaning depends on the operation.

### Response envelope

**Success (`HTTP 200`):**

```json
{
  "success": true,
  "value": { ... }
}
```

**Error (`HTTP 400` / `404` / `500`):**

```json
{
  "success": false,
  "error": "Human-readable error message",
  "screenshotPath": "C:\\path\\to\\failure_20240101_120000_000.png"
}
```

A failure screenshot is automatically captured and its path is included in every error response.

---

### Session & Window Management

| Operation | Required fields | Description | Response `value` |
|---|---|---|---|
| `launch` | `value` (exe path) | Launch an application and start a session | `{ "sessionId": "...", "app": "..." }` |
| `close` / `quit` | — | Close the active application and session | `null` |
| `closewindow` | `value` (partial title) | Close any window whose title contains the value | `null` |
| `maximize` | — | Maximize the active window | `null` |
| `minimize` | — | Minimize the active window | `null` |
| `switchwindow` | `value` (partial title) | Bring a window to the foreground and make it active | `{ "title": "..." }` |
| `refresh` | — | Re-attach to the main window of the active application | `null` |
| `screenshot` | `value` (optional output path) | Take a PNG screenshot; saves to the given path or a temp file | `{ "path": "..." }` |
| `listelements` | `value` (optional control-type filter) | List all descendant elements and their UIA properties | `[ { "name", "automationId", "className", "controlType", "enabled", "visible" }, ... ]` |
| `listwindows` | `value` (optional title filter) | List all top-level windows of the active application | `[ { "title", "automationId", "className" }, ... ]` |

**Example — launch:**

```json
{ "operation": "launch", "value": "C:\\Windows\\System32\\notepad.exe" }
```

**Example — switchwindow:**

```json
{ "operation": "switchwindow", "value": "Notepad" }
```

---

### Element Query

All element-query operations perform a built-in retry for up to **5 seconds** (500 ms polling interval) before throwing.

| Operation | Required fields | Description | Response `value` |
|---|---|---|---|
| `exists` | `locator` | Immediately check if an element exists (no retry) | `{ "exists": true/false }` |
| `waitfor` | `locator`; `value` = timeout seconds (default `10`) | Wait up to N seconds for an element to be visible and enabled | `{ "found": true }` |
| `isenabled` | `locator` | Check whether the element is enabled | `{ "enabled": true/false }` |
| `isvisible` | `locator` | Check whether the element is visible (not off-screen) | `{ "visible": true/false }` |
| `isclickable` | `locator` | Check whether the element is enabled and visible | `{ "clickable": true/false }` |
| `ischecked` | `locator` | Check the toggle/checked state (Toggle or SelectionItem pattern) | `{ "checked": true/false }` |
| `getvalue` | `locator` | Get the Value pattern value (e.g. text in an Edit field) | `{ "value": "..." }` |
| `gettext` | `locator` | Get the visible text via Text, Value, or Name pattern | `{ "text": "..." }` |
| `getname` | `locator` | Get the UIA Name property | `{ "name": "..." }` |
| `getcontroltype` | `locator` | Get the UIA ControlType string | `{ "controlType": "Button" }` |
| `getselected` | `locator` | Get the currently selected item from a combo or list | `{ "selected": "..." }` |
| `gettable` / `gettabledata` | `locator` | Read all rows and column headers from a grid/table element | Table data array |
| `gettableheaders` | `locator` | Read only the column headers from a grid/table element | Headers array |

---

### Position Comparison

These operations require **both** `locator` (element 1) and `locator2` (element 2).

| Operation | Description | Response `value` |
|---|---|---|
| `isrightof` | Whether element 1 is to the right of element 2 | `{ "isRightOf": true/false }` |
| `isleftof` | Whether element 1 is to the left of element 2 | `{ "isLeftOf": true/false }` |
| `isabove` | Whether element 1 is above element 2 | `{ "isAbove": true/false }` |
| `isbelow` | Whether element 1 is below element 2 | `{ "isBelow": true/false }` |
| `getposition` | Bounding rectangles of both elements and all four relative-position booleans | `{ "element1": {rect}, "element2": {rect}, "isRightOf", "isLeftOf", "isAbove", "isBelow" }` |

---

### Element Actions

| Operation | Required fields | Description | Response `value` |
|---|---|---|---|
| `click` | `locator` | Click the element (UIA Invoke pattern preferred, falls back to mouse click) | `null` |
| `doubleclick` | `locator` | Double-click the element | `null` |
| `rightclick` | `locator` | Right-click the element | `null` |
| `hover` | `locator` | Move the mouse pointer to the element's clickable point | `null` |
| `focus` | `locator` | Give keyboard focus to the element | `null` |
| `type` | `locator`, `value` | Focus the element and type the given text | `null` |
| `clear` | `locator` | Clear the text content of an Edit element | `null` |
| `sendkeys` | `locator`, `value` | Focus the element and send a key sequence (see [key tokens](#sendkeys-key-tokens)) | `null` |
| `scroll` | `locator` | Scroll the element into view via the ScrollItem pattern | `null` |
| `check` | `locator` | Set a toggle/checkbox element to the checked state | `null` |
| `uncheck` | `locator` | Set a toggle/checkbox element to the unchecked state | `null` |
| `select` | `locator`; `value` (item name) **or** `index` (0-based) | Select a ComboBox/List item by name or zero-based position | `null` |
| `selectaid` | `locator`, `value` (item AutomationId) | Select a ComboBox item by the item's AutomationId | `null` |
| `typeandselect` | `locator`, `value` | Type a filter string into an editable combo and select the best-matching item | `null` |
| `clickgridcell` | `locator`, `index` (row), `columnIndex` (col) | Click the grid cell at the given zero-based row/column | `null` |
| `doubleclickgridcell` | `locator`, `index` (row), `columnIndex` (col) | Double-click the grid cell at the given zero-based row/column | `null` |
| `draganddrop` | `locator` (source), `locator2` (target) | Drag the source element and drop it onto the target element | `null` |

**Example — draganddrop:**

```json
{
  "operation": "draganddrop",
  "locator":  { "automationId": "sourceItem", "controlType": "ListItem" },
  "locator2": { "automationId": "targetList",  "controlType": "List" }
}
```

**Example — typeandselect:**

```json
{
  "operation": "typeandselect",
  "locator": { "automationId": "countryCombo" },
  "value": "United"
}
```

---

### Alert / Dialog Handling

These operations act on the topmost modal dialog and succeed silently (no error) when no dialog is present.

| Operation | Required fields | Description | Response `value` |
|---|---|---|---|
| `alertok` | — | Click the **OK / Yes / Save** button of the topmost modal dialog | `{ "success": true }` |
| `alertcancel` | — | Click the **Cancel / No** button of the topmost modal dialog | `{ "success": true }` |
| `alertclose` | — | Close the topmost modal dialog via the UIA Window pattern | `{ "success": true }` |
| `popupok` | `value` (partial window title) | Bring the named popup to the foreground, press Enter, and wait for it to close | `{ "success": true }` |

---

### sendkeys Key Tokens

Key tokens are wrapped in braces inside the `value` string, e.g. `"Hello{ENTER}"`. Literal text outside braces is typed character-by-character.

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
| `{F1}`–`{F12}` | Function keys F1 through F12 |

---

### Locator Reference

All `locator` and `locator2` objects support these properties, combined with AND logic:

| Property | Description |
|---|---|
| `automationId` | UIA AutomationId property |
| `name` | UIA Name property (exact match) |
| `className` | UIA ClassName property |
| `controlType` | UIA control type string, e.g. `Button`, `Edit`, `ComboBox`, `CheckBox` |
| `xpath` | XPath-style expression; **overrides all other properties when set** |

**XPath syntax examples:**

```
//Button[@Name='OK']
//*[@AutomationId='myId']
//ComboBox[@Name='Status']/ListItem[@Name='Active']
//Edit[@AutomationId='search' and @Name='Search']
//ListItem[@Name='Item'][2]
```

Supported XPath attributes: `@Name`, `@AutomationId`, `@ClassName`, `@ControlType`.

---

## Simple Convenience Endpoints

Thin, purpose-built wrappers for the most common operations. All return `{ "success": true }` or `{ "success": false, "error": "..." }`.

```
POST http://127.0.0.1:{user-port}/<endpoint>
Authorization: Bearer <token>
Content-Type: application/json
```

### Window / Session

| Endpoint | Request body | Description |
|---|---|---|
| `POST /launch` | `{ "exePath": "C:\\path\\app.exe" }` | Launch an application |
| `POST /close` | `{ "app": "Notepad" }` | Close the first window whose title contains the value |
| `POST /listallwindows` | `{ "window": "Notepad" }` | Returns `{ "found": true/false }` — checks whether any matching window is open |
| `POST /switchwindow` | `{ "windowTitle": "Notepad" }` | Bring a window to the foreground |
| `POST /maximize` | `{ "window": "Notepad" }` | Switch to the named window and maximize it |

### Click

| Endpoint | Request body | Description |
|---|---|---|
| `POST /click/name` | `{ "name": "OK" }` | Click element by UIA Name |
| `POST /click/aid` | `{ "automationId": "btnOK" }` | Click element by AutomationId |
| `POST /click/advanced` | `{ "name": "Save", "controlType": "Button" }` | Click element by Name and/or ControlType |

### Double-click

| Endpoint | Request body | Description |
|---|---|---|
| `POST /doubleclick/name` | `{ "name": "MyFile.txt" }` | Double-click element by UIA Name |
| `POST /doubleclick/aid` | `{ "automationId": "listItem1" }` | Double-click element by AutomationId |

### ComboBox Selection

| Endpoint | Request body | Description |
|---|---|---|
| `POST /select/combobox/name` | `{ "combobox": "Country", "itemName": "United States" }` | Find ComboBox by Name and select an item by its visible text |
| `POST /select/combobox/aid` | `{ "combobox": "Country", "automationId": "item_us" }` | Find ComboBox by Name and select an item by its AutomationId |

### Alert / Dialog

| Endpoint | Request body | Description |
|---|---|---|
| `POST /alert/ok` | *(empty)* | Click OK / Yes on the topmost modal dialog |
| `POST /alert/cancel` | *(empty)* | Click Cancel / No on the topmost modal dialog |
| `POST /alert/close` | *(empty)* | Close the topmost modal dialog via the Window pattern |

---

## Recording API

The recording feature opens a transparent always-on-top overlay bar at the top of the primary screen. User interactions are captured as a JSON action sequence that can be exported and replayed via the `/ui` endpoint.

### POST /record/start

Opens the overlay and optionally launches an application.

**Request body** (all fields optional):

```json
{
  "exePath": "C:\\path\\to\\app.exe",
  "outputPath": "C:\\recordings\\",
  "waitSeconds": 60
}
```

| Field | Description |
|---|---|
| `exePath` | Full path to an executable to launch before recording begins |
| `outputPath` | Directory (or full file path) for the exported JSON file. Defaults to `%TEMP%\DesktopAutomationHelper\Recordings\` |
| `waitSeconds` | Number of seconds before the recording auto-stops and exports |

**Overlay keyboard shortcuts:**

| Shortcut | Action |
|---|---|
| `Ctrl+P` | **Passive** mode — mouse clicks and keyboard input are captured automatically |
| `Ctrl+A` | **Assistive** mode — right-click any element to choose an action from a context menu |
| `Ctrl+S` | Stop recording and export to JSON |

**Response:**

```json
{
  "success": true,
  "message": "Recording started.",
  "launch": { ... },
  "outputPath": "C:\\recordings\\"
}
```

### GET /record/status

Returns whether recording is active, the current mode, and the number of actions captured.

**Response:**

```json
{
  "status": 0,
  "value": {
    "isActive": true,
    "mode": "Passive",
    "startedAt": "2024-01-01T12:00:00+00:00",
    "actionsCount": 12
  }
}
```

### GET /record/actions

Returns all recorded actions collected so far (or after recording stopped).

**Response:**

```json
{
  "status": 0,
  "value": {
    "startedAt": "2024-01-01T12:00:00+00:00",
    "stoppedAt": "2024-01-01T12:05:00+00:00",
    "mode": "Passive",
    "exportedFilePath": "C:\\recordings\\recording_20240101_120000.json",
    "actions": [
      {
        "actionType": "Click",
        "timestamp": "2024-01-01T12:00:05+00:00",
        "mode": "Passive",
        "description": "Click on OK Button",
        "value": null,
        "element": {
          "name": "OK",
          "automationId": "btnOK",
          "className": "Button",
          "controlType": "Button",
          "boundingRectangle": "{X=100,Y=200,Width=80,Height=30}",
          "suggestedXPath": "//Button[@AutomationId='btnOK']"
        }
      }
    ]
  }
}
```

**Action types:** `Click`, `DoubleClick`, `RightClick`, `Hover`, `Select`, `Type`, `TypeAndSelect`, `IsVisible`, `IsClickable`, `IsEnabled`, `IsDisabled`, `IsEditable`, `GetTableHeaders`, `GetTableData`, `Assert`, `IsChecked`, `SelectCheckBox`, `ClearText`, `GetValue`, `Expand`, `Collapse`, `Maximize`, `Minimize`, `CloseWindow`, `SwitchWindow`, `SetValue`, `Scroll`, `DragAndDrop`.

### POST /record/stop

Stops the active recording, writes the JSON export file, and returns the full `RecordingExport` payload. Safe to call even if recording was already stopped (idempotent).

---

## WebDriver Session API

A WebDriver-compatible session API for tools that expect the protocol. All endpoints require `Authorization: Bearer <token>`.

All responses use the WebDriver envelope:

```json
{ "sessionId": "<guid>", "status": 0, "value": { ... } }
```

### POST /session

Creates a new automation session.

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

**Attach to a running process:**

```json
{
  "desiredCapabilities": {
    "appName": "notepad",
    "uiaType": "UIA3"
  }
}
```

**Response:**

```json
{
  "sessionId": "<guid>",
  "status": 0,
  "value": {
    "sessionId": "<guid>",
    "capabilities": {
      "app": "C:\\Windows\\System32\\notepad.exe",
      "uiaType": "UIA3",
      "processId": 12345
    }
  }
}
```

### GET /sessions

Lists all active session IDs.

### GET /session/{sessionId}

Returns information about the session (app path, process ID, UIA type).

### DELETE /session/{sessionId}

Closes the session and disposes the automation backend.

### POST /session/{sessionId}/element

Finds the **first** matching element and returns its cached element ID.

**Request:**

```json
{ "using": "automation id", "value": "OKButton" }
```

Supported `using` strategies:

| Strategy | Description |
|---|---|
| `automation id` / `id` | Match by AutomationId |
| `name` / `link text` | Exact match on UIA Name |
| `partial link text` | Substring match on UIA Name |
| `class name` | Match by ClassName |
| `tag name` | Match by ControlType (e.g. `Button`, `Edit`) |
| `xpath` | XPath-style pattern (see [Locator Reference](#locator-reference)) |

### POST /session/{sessionId}/elements

Finds **all** matching elements and returns a list of element IDs.

### POST /session/{sessionId}/element/{elementId}/element

Finds a child element within the specified parent element.

### POST /session/{sessionId}/element/{elementId}/elements

Finds all child elements within the specified parent element.

### POST /session/{sessionId}/element/{elementId}/click

Performs a mouse click on the element.

### POST /session/{sessionId}/element/{elementId}/doubleclick

Performs a double-click on the element.

### POST /session/{sessionId}/element/{elementId}/rightclick

Performs a right-click on the element.

### POST /session/{sessionId}/element/{elementId}/value

Types text into the element. Supports WebDriver special key codes.

```json
{ "value": ["Hello World", "\uE007"] }
```

`\uE007` is the WebDriver Enter key code.

### POST /session/{sessionId}/element/{elementId}/clear

Clears the text content of the element.

### GET /session/{sessionId}/element/{elementId}/text

Returns the visible text of the element.

### GET /session/{sessionId}/element/{elementId}/name

Returns the control type name of the element (e.g. `"Button"`, `"Edit"`).

### GET /session/{sessionId}/element/{elementId}/attribute/{name}

Returns the named UIA property. Common attribute names: `Name`, `AutomationId`, `ClassName`, `Value`, `IsEnabled`, `IsOffscreen`, `ControlType`, `HelpText`, `BoundingRectangle`.

### GET /session/{sessionId}/element/{elementId}/enabled

Returns `true` if the element is enabled.

### GET /session/{sessionId}/element/{elementId}/displayed

Returns `true` if the element is visible on screen.

### GET /session/{sessionId}/screenshot

Returns a Base64-encoded PNG screenshot of the application window.

---

## Example: Automating Notepad (Python)

```python
import requests

PROBE = "http://localhost:9102"   # fixed well-known probe port

# 1. Discover the token and user-specific port (no auth needed)
info = requests.get(f"{PROBE}/verify").json()["value"]
BASE = f"http://localhost:{info['port']}"
AUTH = {"Authorization": info["authorizationHeader"]}

# 2. Launch Notepad via the unified /ui endpoint
requests.post(f"{BASE}/ui", headers=AUTH, json={
    "operation": "launch",
    "value": r"C:\Windows\System32\notepad.exe"
})

# 3. Type some text
requests.post(f"{BASE}/ui", headers=AUTH, json={
    "operation": "type",
    "locator": {"controlType": "Edit"},
    "value": "Hello from Desktop Automation Driver!"
})

# 4. Save the file (Ctrl+S)
requests.post(f"{BASE}/ui", headers=AUTH, json={
    "operation": "sendkeys",
    "locator": {"controlType": "Edit"},
    "value": "^s"
})

# 5. Close Notepad
requests.post(f"{BASE}/ui", headers=AUTH, json={
    "operation": "close"
})
```

### WebDriver session style (Python)

```python
import requests

PROBE = "http://localhost:9102"
info  = requests.get(f"{PROBE}/verify").json()["value"]
BASE  = f"http://localhost:{info['port']}"
AUTH  = {"Authorization": info["authorizationHeader"]}

# Create session
session    = requests.post(f"{BASE}/session", headers=AUTH, json={
    "desiredCapabilities": {"app": r"C:\Windows\System32\notepad.exe", "uiaType": "UIA3"}
}).json()
session_id = session["sessionId"]

# Find the text area and type into it
el    = requests.post(f"{BASE}/session/{session_id}/element", headers=AUTH,
                      json={"using": "class name", "value": "Edit"}).json()
el_id = el["value"]["elementId"]
requests.post(f"{BASE}/session/{session_id}/element/{el_id}/value", headers=AUTH,
              json={"value": ["Hello!"]})

# Clean up
requests.delete(f"{BASE}/session/{session_id}", headers=AUTH)
```

---

## Building from Source

> **Note:** Building and running requires **Windows** (UIA3 is a Windows-only API).
> On Linux, `<EnableWindowsTargeting>true</EnableWindowsTargeting>` is set in the project file so the solution compiles for CI, but the produced binary will only run on Windows.

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
├── Controllers/
│   ├── VerifyController.cs              GET /verify  (no auth — bootstrap)
│   ├── StatusController.cs              GET /status
│   ├── SessionController.cs             POST/GET/DELETE /session, GET /sessions
│   ├── ElementController.cs             /session/{id}/element/* (WebDriver protocol)
│   ├── UiController.cs                  POST /ui  (unified automation endpoint)
│   ├── SimpleAutomationController.cs    /launch, /click/*, /select/*, /alert/*, …
│   └── RecordingController.cs           /record/start, /record/stop, /record/status, /record/actions
├── Middleware/
│   └── BearerTokenMiddleware.cs         Validates Authorization: Bearer <token> on all routes except /verify
├── Models/
│   ├── Request/                         UiRequest, UiLocator, FindElementRequest, CreateSessionRequest, StartRecordingRequest, …
│   ├── Response/                        UiResponse, WebDriverResponse<T>, VerifyResponse, SessionResponse, StatusResponse, …
│   └── Recording/                       RecordingExport, RecordedAction, ElementInfo, ActionType, RecordingMode
└── Services/
    ├── IDriverContext / DriverContext    FNV-1a user-port derivation (30000–39999) + random Bearer token at startup
    ├── ISessionManager / SessionManager Thread-safe WebDriver session store
    ├── IAutomationService / AutomationService  FlaUI-backed element finding and actions (WebDriver session style)
    ├── AutomationSession                Per-session FlaUI state and element GUID cache
    ├── IUiSessionContext / UiSessionContext     Active-session singleton used by /ui and simple endpoints
    ├── IUiService / UiService           All 40+ /ui operations backed by FlaUI UIA3
    ├── IRecordingService / RecordingService     Recording overlay, action capture, and JSON export
    └── AssistivePopupResolver           Resolves assistive-mode right-click context menus

Program.cs  — Kestrel host (127.0.0.1:{user-port}), BearerTokenMiddleware,
              probe server (HttpListener on localhost:9102 → GET /verify only),
              startup banner
```

### Port derivation

Each Windows user gets a stable port in the range **30000–39999**:

```
port = 30000 + FNV1a32(username.toLowerCase().utf8) % 10000
```

The FNV-1a hash is stable across .NET versions and handles non-ASCII usernames correctly.

### Probe server (port 9102)

A second, lightweight `HttpListener` binds to `localhost:9102` and serves only `GET /verify`. The first driver instance that starts on a machine claims port 9102; subsequent users work through their own user-specific port. `probePort` is `null` in the `/verify` response when the probe port is unavailable.

---

## License

MIT
