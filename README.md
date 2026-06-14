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

All responses from WebDriver-compatible endpoints follow the WebDriver response envelope format:

```json
{
  "sessionId": "<string|null>",
  "status": 0,
  "value": { ... }
}
```

The `status` property indicates the status of the operation: `0` denotes success, and non-zero values denote various error statuses.

Below is the detailed specification of all WebDriver-compatible REST endpoints.

---

### GET /verify *(no authentication required)*

Returns the current running status, username, connection port, fixed probe port (if active), and the Bearer token required for other endpoints. Use this to bootstrap your automated client sessions.

- **Request Format:** No request body required.
- **Success Response (HTTP 200):**
  ```json
  {
    "sessionId": null,
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

---

### GET /status

Checks if the driver is running and ready to accept new sessions. Requires Bearer authentication.

- **Request Format:** No request body required.
- **Success Response (HTTP 200):**
  ```json
  {
    "sessionId": null,
    "status": 0,
    "value": {
      "ready": true,
      "message": "Desktop Automation Driver is running",
      "build": {
        "version": "1.0.0",
        "revision": "",
        "time": "2026-06-14T00:46:12Z"
      }
    }
  }
  ```

---

### POST /session

Creates a new automation session (either launching an application or attaching to a running process) and returns the session capability details. Requires Bearer authentication.

- **Request Format (Launch Application):**
  ```json
  {
    "desiredCapabilities": {
      "app": "C:\\Windows\\System32\\notepad.exe",
      "appArguments": "-arg1 val1",
      "appWorkingDir": "C:\\Windows\\System32",
      "uiaType": "UIA3",
      "launchDelay": 1000
    }
  }
  ```
- **Request Format (Attach to Process):**
  ```json
  {
    "desiredCapabilities": {
      "appName": "notepad",
      "uiaType": "UIA3"
    }
  }
  ```
- **Request Fields:**
  | Field | Type | Required | Description |
  |---|---|---|---|
  | `desiredCapabilities` | object | Yes | Capabilities defined for the session. |
  | `desiredCapabilities.app` | string | No* | Full path to application executable. *Either `app` or `appName` must be set. |
  | `desiredCapabilities.appName` | string | No* | Process name of already-running app to attach to. *Either `app` or `appName` must be set. |
  | `desiredCapabilities.appArguments` | string | No | Command-line arguments to pass when launching the application. |
  | `desiredCapabilities.appWorkingDir` | string | No | Working directory for the application process. |
  | `desiredCapabilities.uiaType` | string | No | UI Automation backend type: `"UIA2"` or `"UIA3"` (default: `"UIA3"`). |
  | `desiredCapabilities.launchDelay` | number | No | Delay in milliseconds to wait after launch before beginning automation (default: `1000`). |

- **Success Response (HTTP 200):**
  ```json
  {
    "sessionId": "4b9a1012-706f-409c-be7c-bc7d35368a62",
    "status": 0,
    "value": {
      "sessionId": "4b9a1012-706f-409c-be7c-bc7d35368a62",
      "capabilities": {
        "app": "C:\\Windows\\System32\\notepad.exe",
        "appName": "notepad",
        "uiaType": "UIA3",
        "processId": 1234
      }
    }
  }
  ```

---

### GET /session/{sessionId}

Retrieves capability information about an active session. Requires Bearer authentication.

- **Success Response (HTTP 200):**
  ```json
  {
    "sessionId": "4b9a1012-706f-409c-be7c-bc7d35368a62",
    "status": 0,
    "value": {
      "app": "C:\\Windows\\System32\\notepad.exe",
      "appName": "notepad",
      "uiaType": "UIA3",
      "processId": 1234
    }
  }
  ```

---

### DELETE /session/{sessionId}

Closes the active session and disposes of the automation backend. Requires Bearer authentication.

- **Success Response (HTTP 200):**
  ```json
  {
    "sessionId": "4b9a1012-706f-409c-be7c-bc7d35368a62",
    "status": 0,
    "value": null
  }
  ```

---

### GET /sessions

Lists all active session IDs. Requires Bearer authentication.

- **Success Response (HTTP 200):**
  ```json
  {
    "sessionId": null,
    "status": 0,
    "value": [
      "4b9a1012-706f-409c-be7c-bc7d35368a62"
    ]
  }
  ```

---

### POST /session/{sessionId}/element

Finds the first matching element in the application window scope. Requires Bearer authentication.

- **Request Format:**
  ```json
  {
    "using": "automation id",
    "value": "OKButton"
  }
  ```
- **Request Fields:**
  | Field | Type | Required | Description |
  |---|---|---|---|
  | `using` | string | Yes | Search strategy: `"automation id"`, `"id"`, `"name"`, `"link text"`, `"partial link text"`, `"class name"`, `"tag name"`, or `"xpath"`. |
  | `value` | string | Yes | The value to search for. |

- **Success Response (HTTP 200):**
  ```json
  {
    "sessionId": "4b9a1012-706f-409c-be7c-bc7d35368a62",
    "status": 0,
    "value": {
      "elementId": "element-a1b2c3d4"
    }
  }
  ```

---

### POST /session/{sessionId}/elements

Finds all matching elements within the application window scope. Requires Bearer authentication.

- **Request Format:** Same as `POST /session/{sessionId}/element`.
- **Success Response (HTTP 200):**
  ```json
  {
    "sessionId": "4b9a1012-706f-409c-be7c-bc7d35368a62",
    "status": 0,
    "value": [
      { "elementId": "element-a1b2c3d4" },
      { "elementId": "element-e5f6g7h8" }
    ]
  }
  ```

---

### POST /session/{sessionId}/element/{elementId}/element

Finds a child element starting from the specified parent element scope. Requires Bearer authentication.

- **Request Format:** Same as `POST /session/{sessionId}/element`.
- **Success Response (HTTP 200):**
  ```json
  {
    "sessionId": "4b9a1012-706f-409c-be7c-bc7d35368a62",
    "status": 0,
    "value": {
      "elementId": "child-element-xyz123"
    }
  }
  ```

---

### POST /session/{sessionId}/element/{elementId}/elements

Finds all matching child elements starting from the specified parent element scope. Requires Bearer authentication.

- **Request Format:** Same as `POST /session/{sessionId}/element`.
- **Success Response (HTTP 200):**
  ```json
  {
    "sessionId": "4b9a1012-706f-409c-be7c-bc7d35368a62",
    "status": 0,
    "value": [
      { "elementId": "child-element-xyz123" }
    ]
  }
  ```

---

### POST /session/{sessionId}/element/{elementId}/click

Performs a mouse click on the specified element. Requires Bearer authentication.

- **Success Response (HTTP 200):**
  ```json
  {
    "sessionId": "4b9a1012-706f-409c-be7c-bc7d35368a62",
    "status": 0,
    "value": null
  }
  ```

---

### POST /session/{sessionId}/element/{elementId}/doubleclick

Performs a double-click on the specified element. Requires Bearer authentication.

- **Success Response (HTTP 200):**
  ```json
  {
    "sessionId": "4b9a1012-706f-409c-be7c-bc7d35368a62",
    "status": 0,
    "value": null
  }
  ```

---

### POST /session/{sessionId}/element/{elementId}/rightclick

Performs a right-click on the specified element. Requires Bearer authentication.

- **Success Response (HTTP 200):**
  ```json
  {
    "sessionId": "4b9a1012-706f-409c-be7c-bc7d35368a62",
    "status": 0,
    "value": null
  }
  ```

---

### POST /session/{sessionId}/element/{elementId}/value

Types keys into the specified element. Supports standard characters and special key codes. Requires Bearer authentication.

- **Request Format:**
  ```json
  {
    "value": ["Hello World", "\uE007"]
  }
  ```
- **Request Fields:**
  | Field | Type | Required | Description |
  |---|---|---|---|
  | `value` | string[] | Yes | Array of keys to send. Concatenated and typed in sequence. Standard Unicode WebDriver key codes are supported. |

- **Success Response (HTTP 200):**
  ```json
  {
    "sessionId": "4b9a1012-706f-409c-be7c-bc7d35368a62",
    "status": 0,
    "value": null
  }
  ```

---

### POST /session/{sessionId}/element/{elementId}/clear

Clears the text of an editable element. Requires Bearer authentication.

- **Success Response (HTTP 200):**
  ```json
  {
    "sessionId": "4b9a1012-706f-409c-be7c-bc7d35368a62",
    "status": 0,
    "value": null
  }
  ```

---

### GET /session/{sessionId}/element/{elementId}/text

Retrieves the visible text of the element. Requires Bearer authentication.

- **Success Response (HTTP 200):**
  ```json
  {
    "sessionId": "4b9a1012-706f-409c-be7c-bc7d35368a62",
    "status": 0,
    "value": "My Output Text"
  }
  ```

---

### GET /session/{sessionId}/element/{elementId}/name

Retrieves the control type name of the element (e.g., `"Button"`, `"Edit"`). Requires Bearer authentication.

- **Success Response (HTTP 200):**
  ```json
  {
    "sessionId": "4b9a1012-706f-409c-be7c-bc7d35368a62",
    "status": 0,
    "value": "Edit"
  }
  ```

---

### GET /session/{sessionId}/element/{elementId}/attribute/{name}

Retrieves the value of a specific named attribute or property of the element. Requires Bearer authentication.
Common attributes: `Name`, `AutomationId`, `ClassName`, `IsEnabled`, `IsOffscreen`, `ControlType`, `Value`, `HelpText`, `BoundingRectangle`.

- **Success Response (HTTP 200):**
  ```json
  {
    "sessionId": "4b9a1012-706f-409c-be7c-bc7d35368a62",
    "status": 0,
    "value": "ok_btn_automation_id"
  }
  ```

---

### GET /session/{sessionId}/element/{elementId}/enabled

Checks if the element is currently enabled. Requires Bearer authentication.

- **Success Response (HTTP 200):**
  ```json
  {
    "sessionId": "4b9a1012-706f-409c-be7c-bc7d35368a62",
    "status": 0,
    "value": true
  }
  ```

---

### GET /session/{sessionId}/element/{elementId}/displayed

Checks if the element is currently visible on screen. Requires Bearer authentication.

- **Success Response (HTTP 200):**
  ```json
  {
    "sessionId": "4b9a1012-706f-409c-be7c-bc7d35368a62",
    "status": 0,
    "value": true
  }
  ```

---

### GET /session/{sessionId}/screenshot

Takes a Base64-encoded PNG screenshot of the application window. Requires Bearer authentication.

- **Success Response (HTTP 200):**
  ```json
  {
    "sessionId": "4b9a1012-706f-409c-be7c-bc7d35368a62",
    "status": 0,
    "value": "iVBORw0KGgoAAAANSUhEUgAA..."
  }
  ```

Full API reference: see [Controllers/ElementController.cs](src/DesktopAutomationDriver/Controllers/ElementController.cs).

---

## Simple Automation API

The Simple Automation API exposes lightweight, action-oriented HTTP POST endpoints wrapping the automation service. Each simple automation request yields a consistent `{"success": true}` or `{"success": false, "error": "..."}` response (or `{"found": true/false}` for list operations). This allows easy scripting and execution without handling element-level session IDs or complex WebDriver state machines.

All Simple Automation routes require the Bearer token in the `Authorization` header.

---

### POST /launch

Launches the application at the given path and creates an active session.

- **Request Format:**
  ```json
  {
    "exePath": "C:\\Windows\\System32\\notepad.exe"
  }
  ```
- **Response Format:**
  ```json
  {
    "success": true
  }
  ```
  Or on failure:
  ```json
  {
    "success": false,
    "error": "Error message describing the failure."
  }
  ```

---

### POST /close

Closes the first top-level window whose title contains the given string.

- **Request Format:**
  ```json
  {
    "app": "Notepad"
  }
  ```
- **Response Format:** Same as `POST /launch`.

---

### POST /listallwindows

Checks whether any top-level window whose title contains the given name is open.

- **Request Format:**
  ```json
  {
    "window": "Notepad"
  }
  ```
- **Response Format:**
  ```json
  {
    "found": true
  }
  ```

---

### POST /switchwindow

Brings the window whose title contains `windowTitle` to the foreground and makes it the active window for subsequent operations.

- **Request Format:**
  ```json
  {
    "windowTitle": "Notepad"
  }
  ```
- **Response Format:** Same as `POST /launch`.

---

### POST /maximize

Switches to the window whose title contains `window` and maximizes it.

- **Request Format:**
  ```json
  {
    "window": "Notepad"
  }
  ```
- **Response Format:** Same as `POST /launch`.

---

### POST /click/name

Clicks the first element whose UIA Name matches the given value.

- **Request Format:**
  ```json
  {
    "name": "OK"
  }
  ```
- **Response Format:** Same as `POST /launch`.

---

### POST /click/aid

Clicks the first element whose UIA AutomationId matches the given value.

- **Request Format:**
  ```json
  {
    "automationId": "btnOK"
  }
  ```
- **Response Format:** Same as `POST /launch`.

---

### POST /click/advanced

Clicks the element matching the supplied Name and/or ControlType. At least one of `name` or `controlType` must be provided; supplying both narrows the match.

- **Request Format:**
  ```json
  {
    "name": "Save",
    "controlType": "Button"
  }
  ```
- **Response Format:** Same as `POST /launch`.

---

### POST /doubleclick/name

Double-clicks the first element whose UIA Name matches the given value.

- **Request Format:**
  ```json
  {
    "name": "MyFile.txt"
  }
  ```
- **Response Format:** Same as `POST /launch`.

---

### POST /doubleclick/aid

Double-clicks the first element whose UIA AutomationId matches the given value.

- **Request Format:**
  ```json
  {
    "automationId": "listItem1"
  }
  ```
- **Response Format:** Same as `POST /launch`.

---

### POST /select/combobox/name

Finds a ComboBox by its Name and selects an item by the item's visible text.

- **Request Format:**
  ```json
  {
    "combobox": "CountryCombo",
    "itemName": "United States"
  }
  ```
- **Response Format:** Same as `POST /launch`.

---

### POST /select/combobox/aid

Finds a ComboBox and selects an item within it by its UIA AutomationId.

- **Request Format:**
  ```json
  {
    "combobox": "CountryCombo",
    "automationId": "item_us"
  }
  ```
- **Response Format:** Same as `POST /launch`.

---

### POST /alert/ok

Finds the topmost modal dialog and clicks its OK or Yes button. Succeeds silently when no alert is present.

- **Request Format:** `{}` (optional request body)
- **Response Format:**
  ```json
  {
    "success": true
  }
  ```

---

### POST /alert/cancel

Finds the topmost modal dialog and clicks its Cancel or No button. Succeeds silently when no alert is present.

- **Request Format:** `{}` (optional request body)
- **Response Format:** Same as `POST /alert/ok`.

---

### POST /alert/close

Finds the topmost modal dialog and closes it via the Window pattern. Succeeds silently when no alert is present.

- **Request Format:** `{}` (optional request body)
- **Response Format:** Same as `POST /alert/ok`.

---
\n## POST /ui — Unified Automation Endpoint

All desktop automation operations are exposed through a single endpoint:

```http
POST http://127.0.0.1:{user-port}/ui
Authorization: Bearer <token>
Content-Type: application/json
```

`GET /ui` is also routed to the same handler, but normal clients should use `POST`
because the operation data is supplied as JSON in the request body.

### Request envelope

The POST `/ui` endpoint accepts a rich, consolidated JSON payload (`UiRequest`) allowing for powerful single-endpoint execution and high-performance, low-latency UI Automation.

```json
{
  "operation": "<operation-name>",
  "locator": {
    "mode": "logical",
    "automationId": "...",
    "name": "...",
    "className": "...",
    "controlType": "...",
    "xpath": "..."
  },
  "locator2": {
    "automationId": "...",
    "name": "...",
    "className": "...",
    "controlType": "...",
    "xpath": "..."
  },
  "parentLocator": {
    "automationId": "..."
  },
  "containerLocator": {
    "automationId": "..."
  },
  "value": "...",
  "clickRegion": "LowerRight",
  "itemRegion": "LeftCenter",
  "index": 0,
  "columnIndex": 0,
  "limit": 50,
  "includeDesktopDescendants": false,
  "allowKeyboardFallback": false,
  "timeoutMs": 5000,
  "fast": false,
  "disableAutoFollow": false,
  "useCache": false,
  "preferAttributes": true,
  "preferXPath": false,
  "xpathOnly": false,
  "fallbackToWindowRootIfParentChildNotFound": false,
  "hwnd": 123456,
  "className": "Notepad",
  "processId": 1234,
  "matchMode": "contains",
  "forceKillAttachedProcess": false,
  "action": "button",
  "button": "OK|Yes|Save",
  "desktopSearch": true,
  "sameProcessOnly": false,
  "makeCurrent": true,
  "offsetX": 0,
  "offsetY": 0,
  "dragStart": "center",
  "dragDurationMs": 250,
  "dragSteps": 10,
  "fromX": 100,
  "fromY": 100,
  "toX": 200,
  "toY": 200,
  "x": 150,
  "y": 150,
  "includeOffscreen": true,
  "maxAttempts": 30,
  "delayMs": 150,
  "direction": "down",
  "amount": 1,
  "mode": "auto",
  "wheelDelta": null,
  "verifyScroll": false,
  "scrollDelayMs": 100,
  "state": "exists",
  "pollIntervalMs": 200,
  "returnAllMatches": false,
  "maxMatches": 500,
  "includeDiagnostics": false,
  "allowBestMatch": false,
  "bestMatch": null,
  "useDesktopRoot": false,
  "useActiveWindowRoot": false,
  "softVerification": false,
  "locatorPath": [],
  "criteria": [],
  "searchRoot": "currentWindow",
  "treeView": "control",
  "backend": "hybrid",
  "returnCandidates": false,
  "debug": false,
  "ambiguity": "error"
}
```

#### Complete `UiRequest` Property Reference

| Property | Type | Description |
|---|---|---|
| `operation` | string | **Required**. The case-insensitive operation name (e.g. `"launch"`, `"click"`, `"type"`, `"gettext"`, etc.). |
| `locator` | object (`UiLocator`) | Primary locator options identifying the target UI element. Multiple non-null criteria are combined with AND logic. |
| `locator2` | object (`UiLocator`) | Secondary locator used by position comparison (`isrightof`, `isleftof`, etc.) and drag/drop operations. |
| `parentLocator` | object (`UiLocator`) | Parent container locator. If set, searches the target child element within this container scope only. |
| `containerLocator` | object (`UiLocator`) | Optional scrollable container locator for scroll-into-view loops. |
| `value` | string | Auxiliary parameter whose meaning depends on the operation (e.g. execution path, typing text, key sequence, timeouts, window title, etc.). |
| `clickRegion` | string | Header click target region for grid header dropdown operations. Defaults to `LowerRight`. |
| `itemRegion` | string | Dropdown ListItem target region for dropdown selections. Defaults to `LeftCenter`. |
| `index` | number | Zero-based index used by `"select"` to pick ComboBox items by position, and by `"clickGridCell"` / `"doubleclickGridCell"` to specify row index. |
| `columnIndex` | number | Zero-based column index used by grid cell click operations. |
| `limit` | number | Maximum returned items for list/diagnostic operations. |
| `includeDesktopDescendants`| boolean | When true, listwindows also scans desktop descendants. Defaults to false. |
| `allowKeyboardFallback` | boolean | When true, ComboBox selection may fallback to keyboard type-ahead on visible list failure. |
| `timeoutMs` | number | Overrides the default operation timeout in milliseconds. |
| `fast` | boolean | Forces short timeout (100ms), no desktop popup scanning, and cached elements. |
| `disableAutoFollow` | boolean | Disables automatic popup/dialog window follow checks for this operation. |
| `useCache` | boolean | Enables element locator caching. Returns a previously resolved element on matching criteria. |
| `preferAttributes` | boolean | Attempts attribute-based criteria match (AutomationId, Name, ClassName) before XPath. (Default for fast operations). |
| `preferXPath` | boolean | Evaluates XPath criteria before attribute-based checks. |
| `xpathOnly` | boolean | Restricts lookup exclusively to XPath; fails if locator lacks an XPath expression. |
| `fallbackToWindowRootIfParentChildNotFound` | boolean | Retries a failed parent-child lookup scope against the full window root before giving up. |
| `hwnd` | number | Native window handle (HWND) match filter for window operations. |
| `className` | string | Win32 class name filter for 'switchwindow'. |
| `processId` | number | Process ID override for 'switchwindow'. |
| `matchMode` | string | Title matching mode: `exact`, `contains` (default), `regex`. |
| `forceKillAttachedProcess` | boolean | When true, 'quit' force-kills the attached process tree unconditionally. |
| `action` | string | Popup action type: `button` (default), `close`, `enter`, `escape`, `makecurrent`. |
| `button` | string | Pipe-separated button candidate names to click (e.g. `OK|Yes|Save`). |
| `desktopSearch` | boolean | If true (default), scans desktop root during popup discovery. |
| `sameProcessOnly` | boolean | Restricts popup scanning to windows matching current app PID. |
| `makeCurrent` | boolean | Focuses and sets the discovered popup as active window. |
| `offsetX`, `offsetY` | number | Drag offset movement vectors in pixels for dragbyoffset. |
| `dragStart` | string | Rectangle drag anchor point: `center`, `topLeft`, `bottomLeft`, etc. (Default: `"center"`). |
| `dragDurationMs` | number | Total drag animation duration in milliseconds. Defaults to 250. |
| `dragSteps` | number | Interpolated mouse move step frames for dragbyoffset. |
| `fromX`, `fromY`, `toX`, `toY` | number | Mouse start and end screen coordinates for dragcoordinates. |
| `x`, `y` | number | Targeted screen coordinate offsets for low-level cursor/scroll. |
| `includeOffscreen` | boolean | Includes offscreen elements during scroll-into-view searches. |
| `maxAttempts` | number | Maximum scroll loop retry attempts before timing out. |
| `delayMs` | number | Polling iteration delay in milliseconds during scroll-into-view. |
| `direction` | string | Direction vector: `up`, `down`, `left`, `right`. |
| `amount` | number | Wheel ticks or pattern-specific unit amounts. |
| `mode` | string | Scroll strategy: `auto` (default), `wheel`, `pattern`. |
| `wheelDelta` | number | Raw mouse wheel delta values. Negative = down/left, positive = up/right. |
| `verifyScroll` | boolean | Ensures scroll position actually modified during wheel scrolls. |
| `scrollDelayMs` | number | Settling delay in milliseconds after scrolling completes. Defaults to 100. |
| `state` | string | Target element state check for wait operations (e.g. `exists`, `visible`, `enabled`). |
| `pollIntervalMs` | number | Query polling sleep interval in milliseconds during waits. |
| `returnAllMatches` | boolean | Returns complete matched element list instead of just first match. |
| `maxMatches` | number | Result count limit when returnAllMatches is enabled. |
| `includeDiagnostics` | boolean | Returns verbose trace metadata of resolved candidate scores on success. |
| `allowBestMatch` | boolean | Permits partial best-match scoring when exact attributes are missing. |
| `bestMatch` | string | Best-match fallback text hint. |
| `useDesktopRoot` | boolean | Searches from UIA desktop root, enabling multi-process element lookups. |
| `useActiveWindowRoot` | boolean | Always looks up elements starting from active foreground window. |
| `softVerification` | boolean | Disables throwing errors on post-action check failures (emits warnings instead). |
| `locatorPath` | array | Sequence of locators for deep pywinauto-style path traversals. |
| `criteria` | array | Extra filter constraints for complex traversals. |
| `searchRoot` | string | Start node: `currentWindow`, `desktop`, `foreground`, `activePopup`, `parent`. |
| `treeView` | string | Selection of UIA tree views: `control` (default), `content`, `raw`. |
| `backend` | string | Automation technology stack choice: `uia`, `win32`, `hybrid` (default). |
| `returnCandidates` | boolean | Returns diagnostic candidates list on failure or ambiguous resolution. |
| `debug` | boolean | Enables verbose pywinauto-style diagnostic tracing. |
| `ambiguity` | string | Resolution handling on duplicates: `error`, `first`, `all`. |

---

### Response formats

Success Response Envelope (HTTP 200):

```json
{
  "success": true,
  "value": { ... }
}
```

The type and schema of the `value` property depends on the requested operation (e.g., check operations return boolean `{"exists": true}`, element detail returns metadata object, listing returns lists, click/actions return `null`).

Error Response Envelope:
Client/request errors return `400 BadRequest`, not-found/operation failures return `404 NotFound`, and unhandled internal driver failures return `500 InternalServerError`.

```json
{
  "success": false,
  "value": null,
  "error": "Detailed error message describing what failed.",
  "screenshotPath": "C:\\Users\\alice\\AppData\\Local\\Temp\\DesktopAutomationHelper\\Failures\\failure-screenshot.png"
}
```

---

### Locator reference

All `locator`, `locator2`, `parentLocator`, and `containerLocator` objects are defined by the `UiLocator` schema, supporting detailed attribute and pywinauto-style advanced scoring filters.

| Property | Type | Description |
|---|---|---|
| `mode` | string | Locator strategy mode hint (e.g. `"logical"` for deep visual menu traversal). |
| `name` | string | Exact match on UIA Name attribute (element label). |
| `nameRegex` | string | Regex match on UIA Name attribute. |
| `automationId` | string | Match on UIA AutomationId attribute. |
| `automationIdRegex` | string | Regex match on UIA AutomationId. |
| `className` | string | Match on UIA ClassName attribute. |
| `classNameRegex` | string | Regex match on ClassName. |
| `controlType` | string | Match on UIA ControlType string (e.g. `Button`, `Edit`, `ComboBox`, `Window`). |
| `xpath` | string | XPath-style traversal expression. When set, other attribute properties are bypassed. |
| `hwnd` | number | Exact native window handle (HWND). If provided, resolves the element directly from window. |
| `processId` | number | Filter by Win32 Process ID. |
| `controlId` | number | Filter by Win32 Control ID (DlgCtrlID). |
| `frameworkId` | string | Filter by target framework identifier (e.g. `"WPF"`, `"WinForm"`, `"Win32"`). |
| `runtimeId` | string | Exact UIA RuntimeId match (e.g. `"42.1234"`). Useful for pinning single UI elements. |
| `value` | string | Match on ValuePattern value. |
| `valueRegex` | string | Regex match on ValuePattern value. |
| `text` | string | Match on TextPattern content (or fallback element Name). |
| `visible` | boolean | Visibility status filter (true = visible only, false = hidden only). |
| `enabled` | boolean | Enabled status filter. |
| `offscreen` | boolean | Offscreen status filter. |
| `matchMode` | string | Global match mode for strings: `exact` (default), `contains`, `startswith`, `regex`. |
| `nameMatchMode` | string | Overrides `matchMode` specifically for Name. |
| `automationIdMatchMode` | string | Overrides `matchMode` for AutomationId. |
| `classNameMatchMode` | string | Overrides `matchMode` for ClassName. |
| `valueMatchMode` | string | Overrides `matchMode` for Value. |
| `textMatchMode` | string | Overrides `matchMode` for Text. |
| `foundIndex` | number | Scored, filtered candidate index selection (matches pywinauto `found_index`). |
| `ctrlIndex` | number | Zero-based pre-filtered raw candidate retrieval index (matches pywinauto `ctrl_index`). |
| `depth` | number | Maximum UIA tree traversal depth (Default: 20). |
| `topLevelOnly` | boolean | Restricts candidates to direct parent-root children (depth = 1). |
| `activeOnly` | boolean | Restricts element query search to active foreground window scope. |
| `includeOffscreen` | boolean | Scans and returns elements physically off-screen. |
| `role` | string | Filter by ARIA role / LocalizedControlType. |
| `bestMatch` | string | Weight name scoring hint used by advanced fuzzy search. |

---
\n### `/ui` operations

#### Session, window, and inspection

| Operation | Required fields | Description | Response `value` |
|---|---|---|---|
| `launch` | `value` | Launch an executable and make it the active session. | `{ "sessionId": "...", "app": "..." }` |
| `close`, `quit` | — | Close the active application/session. | `null` |
| `closewindow` | `value` | Close a window whose title contains `value`. | `null` |
| `maximize` | — | Maximize the active session window. | `null` |
| `minimize` | — | Minimize the active session window. | `null` |
| `switchwindow` | `value` | Focus/attach to a window whose title contains `value`. | Window/session metadata. |
| `switchwinodw` (legacy misspelled alias), `switch_window`, `switchto` | `value` | Aliases for `switchwindow`. | Window/session metadata. |
| `refresh` | — | Re-attach to the active application's main window. | Session/window metadata. |
| `screenshot` | optional `value` | Take a screenshot. If `value` is a file path, save it there; otherwise return Base64 PNG. | `{ "screenshot": "<base64>" }` or file metadata |
| `listelements` | optional `locator`, `limit` | List matching elements and properties. | `{ "elements": [ ... ] }` |
| `listwindows` | optional `includeDesktopDescendants`, `limit` | List visible top-level windows; optionally include desktop descendants. | `{ "windows": [ ... ] }` |
| `getcurrentroot` | optional `locator` | Debug the active root/window context. | Root/window metadata. |
| `findlocator` | `locator` | Debug locator resolution and return matched element metadata. | Element metadata. |

#### Element query

| Operation | Required fields | Description | Response `value` |
|---|---|---|---|
| `exists` | `locator` | Check whether an element exists without waiting. | `{ "exists": true/false }` |
| `waitfor` | `locator`, `value` | Wait up to `value` seconds for an element. | `{ "found": true/false }` |
| `isenabled` | `locator` | Check whether an element is enabled. | `{ "isEnabled": true/false }` |
| `isvisible` | `locator` | Check whether an element is visible. | `{ "isVisible": true/false }` |
| `isclickable` | `locator` | Check whether an element is visible and enabled. | `{ "isClickable": true/false }` |
| `iseditable` | `locator` | Check whether an element supports text editing/value input. | `{ "isEditable": true/false }` |
| `ischecked` | `locator` | Check checkbox/toggle state. | `{ "isChecked": true/false }` |
| `getvalue` | `locator` | Read the UIA Value pattern value. | `{ "value": "..." }` |
| `gettext` | `locator` | Read visible text/name/value text. | `{ "text": "..." }` |
| `getname` | `locator` | Read UIA Name. | `{ "name": "..." }` |
| `getcontroltype` | `locator` | Read UIA ControlType. | `{ "controlType": "Button" }` |
| `getselected` | `locator` | Read selected combo/list item. | `{ "selected": "..." }` |
| `gettable`, `gettabledata` | `locator` | Read grid/table headers and row data. | `{ "headers": [...], "rows": [[...], ...] }` |
| `gettableheaders` | `locator` | Read only grid/table headers. | `{ "headers": [...] }` |

#### Position comparison

These operations require both `locator` and `locator2`.

| Operation | Description | Response `value` |
|---|---|---|
| `isrightof` | Whether element 1 is to the right of element 2. | `{ "isRightOf": true/false }` |
| `isleftof` | Whether element 1 is to the left of element 2. | `{ "isLeftOf": true/false }` |
| `isabove` | Whether element 1 is above element 2. | `{ "isAbove": true/false }` |
| `isbelow` | Whether element 1 is below element 2. | `{ "isBelow": true/false }` |
| `getposition` | Return both rectangles and all relative position checks. | `{ "element1": {...}, "element2": {...}, "isRightOf": ..., ... }` |

#### Element actions

| Operation | Required fields | Description | Response `value` |
|---|---|---|---|
| `click` | `locator` | Click/invoke an element. | `null` |
| `doubleclick` | `locator` | Double-click an element. | `null` |
| `rightclick` | `locator` | Right-click an element. Uses FlaUI first, then physical right-click fallback when needed. | `{ "rightClicked": true, "strategy": "...", "element": {...} }` |
| `contextmenupath` | `locator`, `value` | Right-click an element, then click an item (or nested path) in the context menu that appears. `value` is a `>`-separated path such as `Export Excel` or `Export > Export Excel`. | `{ "selected": "Export Excel", "strategy": "context-menu-path", "element": {...} }` |
| `hover` | `locator` | Move mouse over an element. | `null` |
| `focus` | `locator` | Give keyboard focus to an element. | `null` |
| `type` | `locator`, `value` | Type text into an element. Date pickers use segmented date typing when applicable. | `null` or typed date metadata |
| `typedate` | `locator`, `value` | Type a WinForms `SysDateTimePick32` date as MM → DD → YYYY segments. | `{ "typed": true, "strategy": "date-segments", ... }` |
| `clear` | `locator` | Clear an editable element. Tries ValuePattern, TextBox, Ctrl+A+Backspace, then Ctrl+A+Delete. | `{ "cleared": true, "strategy": "...", "element": {...} }` |
| `sendkeys` | `value`; optional `locator` | Send text/key tokens. Supports aliases like `CTRL+A`, `BACKSPACE`, `DELETE`, `ENTER`. When `locator` is provided the element is focused first; when omitted, keys are sent to the currently focused element in the active window. | `{ "sent": true, "original": "...", "normalized": "...", "target": {...} }` |
| `scroll` | `locator` | Scroll element into view (UIA ScrollItem pattern). | `null` |
| `mousescroll` / `wheelscroll` | optional `locator`, optional `value` | Scroll the mouse wheel. `value` is the number of wheel clicks (positive = up, negative = down) or `"up"` / `"down"` (±3 clicks). Defaults to 3 clicks down. When `locator` is provided the cursor moves to the element first. | `{ "scrolled": true, "wheelClicks": n, "direction": "...", "target": {...} }` |
| `check` | `locator` | Set checkbox/toggle to checked. | `null` |
| `uncheck` | `locator` | Set checkbox/toggle to unchecked. | `null` |
| `select` | `locator`, `value` or `index` | Select combo/list item by visible text or zero-based index. | `null` |
| `selectaid` | `locator`, `value` | Select combo/list item by AutomationId. | `null` |
| `selectcomboboxitem` | `locator`, `value` | Select a ComboBox item by text using expanded dropdown search. | Selection metadata |
| `typeandselect` | `locator`, `value` | Type filter text and select the first matching dropdown item. | `null` |
| `clickgridcell` | `locator`, `index`, `columnIndex` | Click a grid cell by row and column. | `null` |
| `doubleclickgridcell` | `locator`, `index`, `columnIndex` | Double-click a grid cell by row and column. | `null` |
| `openheaderdropdown` | `locator`, optional `clickRegion` | Open a grid header dropdown and list popup items. | `{ "opened": true, "listFound": true/false, "items": [...] }` |
| `selectheaderdropdownitem` | `locator`, `value`, optional `clickRegion`, `itemRegion` | Open a grid header dropdown and select a matching `ListItem`. | `{ "selected": "...", "header": "..." }` |
| `draganddrop` | `locator`, `locator2` | Drag source element to target element. | `null` |

#### Menu operations

| Operation | Required fields | Description | Response `value` |
|---|---|---|---|
| `clickmenu` | `locator`, `value` | Click a visual menu path starting from the located menu/header. | Selection metadata |
| `clickmenupath` | `locator`, `value` | Click a menu path. Uses logical mode when `locator.mode` is `logical`; otherwise visual traversal. | Selection metadata |
| `clicklogicalmenupath` | `locator`, `value` | Click a logical UIA menu path. | Selection metadata |
| `clickmenulogical` | `locator`, `value` | Alias for `clicklogicalmenupath`. | Selection metadata |
| `menupath` | `locator`, `value` | Alias for `clicklogicalmenupath`. | Selection metadata |
| `inspectlogicalmenu` | `locator` | Inspect logical menu descendants/candidates for debugging. | Menu debug metadata |
| `inspectmenupathcandidates` | `value` | Inspect candidate menu parent chains for a path such as `DQA>Level 17`. | Candidate metadata |
| `dumpmenus`, `dumplogicalmenus` | optional `locator`, `limit` | Dump logical menu items for diagnostics. | Menu dump metadata |
| `selectdynamicmenuitem` | `locator`, `value` | Compatibility operation for a dynamic submenu item; delegates to dynamic path traversal. | `{ "selected": "...", "parent": "...", "dropdown": "..." }` |
| `selectdynamicmenupath` | `locator`, `value` | Open a dynamic root `MenuItem` and select child path `Child>Leaf` or full `Root>Child>Leaf`. | `{ "selected": "Child>Leaf", "parent": "...", "dropdown": "..." }` |
| `contextmenupath` | `locator`, `value` | Right-click the located element, then navigate the application's context menu using the `>`-separated path in `value` and click the final item. Supports multi-level submenus (e.g. `Export > Export Excel`). See also the **Element actions** table for examples. | `{ "selected": "Export > Export Excel", "strategy": "context-menu-path", "element": {...} }` |

Dynamic menu example:

```json
{
  "operation": "selectdynamicmenupath",
  "locator": { "name": "File", "controlType": "MenuItem" },
  "value": "Recent>Report1"
}
```

Logical menu example with duplicate leaf names:

```json
{
  "operation": "menupath",
  "locator": { "name": "DQA", "controlType": "MenuItem", "mode": "logical" },
  "value": "DQA>Level 17"
}
```

When multiple leaves share the same label, include the parent chain in `value`; for example
`DQA>Level 17` is matched by its full chain and will not select `Tag1>Level 17`.

#### Alert and popup operations

| Operation | Required fields | Description | Response `value` |
|---|---|---|---|
| `alertok` | — | Click OK/Yes/Save on the topmost modal dialog. | `{ "success": true }` |
| `alertcancel` | — | Click Cancel/No on the topmost modal dialog. | `{ "success": true }` |
| `alertclose` | — | Close the topmost modal dialog via Window pattern. | `{ "success": true }` |
| `popupok` | optional `locator` | Click OK on a popup/dialog, optionally scoped by locator. | `{ "success": true }` |

#### Request examples

Launch Notepad:

```json
{ "operation": "launch", "value": "C:\\Windows\\System32\\notepad.exe" }
```

Type text:

```json
{
  "operation": "type",
  "locator": { "className": "Edit", "controlType": "Edit" },
  "value": "Hello from /ui"
}
```

Drag and drop:

```json
{
  "operation": "draganddrop",
  "locator":  { "automationId": "sourceId", "controlType": "ListItem" },
  "locator2": { "automationId": "targetId", "controlType": "List" }
}
```

#### `sendkeys` tokens

Key tokens are wrapped in braces inside the `value` string, e.g. `"Hello{ENTER}"`.

Modifier prefixes (AutoIt style):
- `^x` or `^{KEYNAME}` — Ctrl + key, e.g. `^a` = Ctrl+A, `^{DELETE}` = Ctrl+Delete
- `+x` or `+{KEYNAME}` — Shift + key, e.g. `+a` = Shift+A, `+{TAB}` = Shift+Tab
- `%x` or `%{KEYNAME}` — Alt + key, e.g. `%f` = Alt+F, `%{F4}` = Alt+F4

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
| `{WIN}` / `{LWIN}` / `{RWIN}` | Windows key |
| `{PRINTSCREEN}` / `{PRTSC}` | Print Screen |
| `{CAPSLOCK}` | Caps Lock |
| `{NUMLOCK}` | Num Lock |
| `{SCROLLLOCK}` | Scroll Lock |
| `{PAUSE}` | Pause / Break |
| `{APPS}` | Application / Context Menu key |

#### Common `sendkeys` examples

| Goal | `value` |
|---|---|
| Select all | `^a` |
| Copy | `^c` |
| Cut | `^x` |
| Paste | `^v` |
| Undo | `^z` |
| Redo | `^y` |
| Save | `^s` |
| Backspace | `{BACKSPACE}` |
| Delete | `{DELETE}` |
| Enter / confirm | `{ENTER}` |
| Escape / cancel | `{ESC}` |

Select all and delete (keyboard alternative to `clear`):

```json
{ "operation": "sendkeys", "value": "^a{DELETE}" }
```

Send CTRL+A to a specific element:

```json
{
  "operation": "sendkeys",
  "locator": { "automationId": "myTextBox", "controlType": "Edit" },
  "value": "^a"
}
```

User-friendly alias form (`CTRL+A`) also works:

```json
{
  "operation": "sendkeys",
  "locator": { "automationId": "txtName", "controlType": "Edit" },
  "value": "CTRL+A"
}
```

Send Enter / Backspace / Delete to the focused control:

```json
{ "operation": "sendkeys", "value": "ENTER" }
```

```json
{ "operation": "sendkeys", "value": "BACKSPACE" }
```

```json
{ "operation": "sendkeys", "value": "DELETE" }
```

Clear an input field:

```json
{
  "operation": "clear",
  "locator": { "automationId": "txtName", "controlType": "Edit" }
}
```

Right-click an element:

```json
{
  "operation": "rightclick",
  "locator": { "name": "Processdate", "controlType": "Header" }
}
```

Right-click an element and click a context menu item (e.g. **Copy**):

```json
{
  "operation": "contextmenupath",
  "locator": { "name": "Grid", "controlType": "Table" },
  "value": "Copy"
}
```

Response:

```json
{
  "sessionId": null,
  "status": 0,
  "value": {
    "selected": "Copy",
    "strategy": "context-menu-path",
    "element": { "name": "Grid", "controlType": "Table", "automationId": "dgData" }
  }
}
```

Right-click and click a **nested** context menu item (e.g. **Export → Export Excel**):

```json
{
  "operation": "contextmenupath",
  "locator": { "name": "Grid", "controlType": "Table" },
  "value": "Export > Export Excel"
}
```

Response:

```json
{
  "sessionId": null,
  "status": 0,
  "value": {
    "selected": "Export > Export Excel",
    "strategy": "context-menu-path",
    "element": { "name": "Grid", "controlType": "Table", "automationId": "dgData" }
  }
}
```

#### `mousescroll` examples

Scroll down 5 clicks over a list:

```json
{
  "operation": "mousescroll",
  "locator": { "automationId": "myList", "controlType": "List" },
  "value": "-5"
}
```

Scroll down using keyword value:

```json
{
  "operation": "mousescroll",
  "locator": { "automationId": "dglnd", "controlType": "Table" },
  "value": "down"
}
```

Scroll down 5 clicks with `wheelscroll`:

```json
{
  "operation": "wheelscroll",
  "locator": { "automationId": "dglnd", "controlType": "Table" },
  "value": "-5"
}
```

Scroll up 3 clicks at the current cursor position:

```json
{ "operation": "mousescroll", "value": "up" }
```

---

## Recording API

Recording endpoints open an always-on-top overlay and export captured actions to JSON.
All recording endpoints require the Bearer token.

### POST /record/start

Starts a recording session. If a `/ui` session is already active, recording is scoped to
that application's process/window. If `exePath` is supplied, the driver launches that
application first and records it as the target.

Request body fields are optional:

```json
{
  "exePath": "C:\\Windows\\System32\\notepad.exe",
  "outputPath": "C:\\Temp\\recordings",
  "waitSeconds": 60
}
```

| Field | Description |
|---|---|
| `exePath` | Full path to an executable to launch before recording. Omit to record the current/active target. |
| `outputPath` | Directory or full JSON file path for the export. Defaults to `%TEMP%\\DesktopAutomationHelper\\Recordings\\`. |
| `waitSeconds` | Optional auto-stop timeout. Omit to stop manually with Ctrl+S or `POST /record/stop`. |

Success response:

```json
{
  "success": true,
  "message": "Recording started.",
  "launch": {
    "success": true,
    "processId": 1234,
    "windowTitle": "Untitled - Notepad",
    "window": {
      "title": "Untitled - Notepad",
      "x": 100,
      "y": 100,
      "width": 1200,
      "height": 800,
      "windowState": "Normal",
      "isMaximized": false,
      "isMinimized": false,
      "isFullScreen": false
    }
  },
  "outputPath": "C:\\Temp\\recordings"
}
```

If another recording is already active, the endpoint returns `409` with a WebDriver error
envelope.

### Recording modes

After `POST /record/start`, use the overlay hotkeys to select the mode:

| Hotkey | Mode | Details |
|---|---|---|
| `Ctrl+P` | Passive | Automatically captures mouse clicks and keyboard actions with the element under the cursor. Use this for quick observation-style recordings. |
| `Ctrl+A` | Assistive | Right-click an element to open an action menu, then choose what to record/perform. Use this when you need deterministic locators, assertions, menu actions (including right-click context menus), grid/header dropdown actions, or dynamic menu paths. |
| `Ctrl+S` | Stop | Stops recording and writes the JSON export. Equivalent to `POST /record/stop`. |

Assistive mode records richer action metadata, including the selected action type,
resolved element locator, optional target element, menu path details, pointer diagnostics,
and operation-specific metadata. `/playback` replays Assistive actions only; Passive
actions are exported for inspection but skipped during playback.

Assistive dynamic menu example: selecting `File > Recent > Report1` is exported as a
path-oriented action such as:

```json
{
  "actionType": "click",
  "mode": "assistive",
  "operation": "selectdynamicmenupath",
  "value": "File>Recent>Report1",
  "description": "Select dynamic menu path File>Recent>Report1"
}
```

Assistive right-click context menu example: right-clicking a grid and selecting **Copy**
(or a nested path like **Export > Export Excel**) is exported as:

```json
{
  "actionType": "MenuPathClick",
  "mode": "Assistive",
  "operation": "contextmenupath",
  "value": "Copy",
  "element": { "name": "Grid", "controlType": "Table", "automationId": "dgData" },
  "metadata": {
    "fallbackPoint": "640,480",
    "strategy": "context-menu-path"
  },
  "description": "Select context menu path Copy on Grid Table"
}
```

Nested submenu example (`Export > Export Excel`):

```json
{
  "actionType": "MenuPathClick",
  "mode": "Assistive",
  "operation": "contextmenupath",
  "value": "Export > Export Excel",
  "element": { "name": "Grid", "controlType": "Table", "automationId": "dgData" },
  "metadata": {
    "fallbackPoint": "640,480",
    "strategy": "context-menu-path"
  },
  "description": "Select context menu path Export > Export Excel on Grid Table"
}
```

These recorded actions are replayed directly by `/playback` using `contextmenupath`.

### GET /record/status

Returns the active state, current mode, start time, and captured action count.

```json
{
  "sessionId": null,
  "status": 0,
  "value": {
    "isActive": true,
    "mode": "Assistive",
    "startedAt": "2026-05-10T08:25:31.095+00:00",
    "actionsCount": 3
  }
}
```

### GET /record/actions

Returns the current recording export, including actions captured so far and the exported
file path if the recording has already stopped.

### POST /record/stop

Stops the active recording session, writes the JSON export, and returns the same export
format as `/record/actions`. Calling it after the overlay was stopped with `Ctrl+S` is safe.

### Recording export format

```json
{
  "sessionId": null,
  "status": 0,
  "value": {
    "startedAt": "2026-05-10T08:25:31.095+00:00",
    "stoppedAt": "2026-05-10T08:26:10.000+00:00",
    "mode": "Assistive",
    "screen": { "x": 0, "y": 0, "width": 1920, "height": 1080 },
    "launch": { "success": true, "processId": 1234, "windowTitle": "Untitled - Notepad" },
    "exportedFilePath": "C:\\Temp\\recordings\\recording-20260510.json",
    "actions": [
      {
        "actionType": "Click",
        "timestamp": "2026-05-10T08:25:45.000+00:00",
        "mode": "Assistive",
        "element": { "name": "OK", "controlType": "Button" },
        "description": "Click on OK Button",
        "value": null,
        "operation": "click",
        "menuPath": null,
        "targetElement": null,
        "pointerContext": { },
        "metadata": { }
      }
    ]
  }
}
```

Recorded action types include `Click`, `MenuPathClick`, `DoubleClick`, `RightClick`,
`Hover`, `Select`, `Type`, `TypeAndSelect`, `IsVisible`, `IsClickable`, `IsEnabled`,
`IsDisabled`, `IsEditable`, `GetTableHeaders`, `GetTableData`, `Assert`, `IsChecked`,
`SelectCheckBox`, `ClearText`, `GetValue`, `Expand`, `Collapse`, `Maximize`, `Minimize`,
`CloseWindow`, `SwitchWindow`, `SetValue`, `Scroll`, and `DragAndDrop`.

---

## POST /playback

Replays Assistive recording JSON by converting recorded actions to `/ui` operations and
executing them sequentially. Passive actions are skipped because they do not always contain
enough deterministic action metadata.

```http
POST http://127.0.0.1:{user-port}/playback
Authorization: Bearer <token>
Content-Type: application/json
```

Accepted request formats:

1. Raw recording export object returned by `/record/actions` or `/record/stop`:

```json
{
  "startedAt": "2026-05-10T08:25:31.095+00:00",
  "mode": "Assistive",
  "actions": [ ... ]
}
```

2. A bare actions array:

```json
[
  { "actionType": "Click", "mode": "Assistive", "element": { "name": "OK", "controlType": "Button" } }
]
```

3. A wrapper with options:

```json
{
  "recording": {
    "mode": "Assistive",
    "actions": [ ... ]
  },
  "continueOnError": true,
  "delayMs": 250
}
```

4. A WebDriver-style value wrapper containing a recording export:

```json
{
  "value": {
    "mode": "Assistive",
    "actions": [ ... ]
  }
}
```

Playback options:

| Field | Description |
|---|---|
| `continueOnError` | Defaults to false. When true, playback continues after failed actions and reports each failure. |
| `delayMs` | Optional delay between successfully executed actions. Defaults to 0. |

Response:

```json
{
  "sessionId": null,
  "status": 0,
  "value": {
    "totalActions": 3,
    "executedActions": 2,
    "skippedActions": 1,
    "failedActions": 0,
    "completed": true,
    "actions": [
      {
        "index": 0,
        "success": true,
        "skipped": false,
        "skipReason": null,
        "operation": "click",
        "actionType": "Click",
        "description": "Click on OK Button",
        "error": null,
        "result": null
      }
    ]
  }
}
```

Playback action mapping:

| Recorded action | `/ui` operation |
|---|---|
| `Click` | `click`, `sendkeys`, `alertok`, or `alertcancel` depending on recorded value/description |
| `MenuPathClick` | `contextmenupath` (when recorded via right-click context menu) or `clicklogicalmenupath` (when recorded via logical menu bar) |
| `DoubleClick` | `doubleclick` |
| `RightClick` | `rightclick` |
| `Hover` | `hover` |
| `Select` | `select` |
| `Type` | `type` or `sendkeys` |
| `TypeAndSelect` | `typeandselect` |
| `IsVisible` | `isvisible` |
| `IsClickable` | `isclickable` |
| `IsEnabled` | `isenabled` |
| `IsEditable` | `iseditable` |
| `GetTableHeaders` | `gettableheaders` |
| `GetTableData` | `gettabledata` |
| `IsChecked` | `ischecked` |
| `SelectCheckBox` | `check` or `uncheck` |
| `ClearText` | `clear` |
| `GetValue` | `getvalue` |
| `Expand`, `Collapse` | `click` |
| `Maximize` | `maximize` |
| `Minimize` | `minimize` |
| `CloseWindow` | `closewindow` |
| `SwitchWindow` | `switchwindow` |
| `SetValue` | `type` |
| `Scroll` | `scroll` |

If a recorded action has an explicit `operation`, playback uses it directly. Unsupported
or non-deterministic actions are marked as skipped with a `skipReason`.

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
