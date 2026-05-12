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

```http
POST http://127.0.0.1:{user-port}/ui
Authorization: Bearer <token>
Content-Type: application/json
```

`GET /ui` is also routed to the same handler, but normal clients should use `POST`
because the operation data is supplied as JSON in the request body.

### Request envelope

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
  "value": "...",
  "clickRegion": "LowerRight",
  "itemRegion": "LeftCenter",
  "index": 0,
  "columnIndex": 0,
  "limit": 50,
  "includeDesktopDescendants": false
}
```

| Field | Type | Used by | Notes |
|---|---|---|---|
| `operation` | string | all operations | Required. Case-insensitive. |
| `locator` | object | element/menu/grid operations | Primary element locator. Multiple non-null properties are combined with AND logic. |
| `locator2` | object | position and drag/drop operations | Secondary element locator. |
| `value` | string | operation-specific | Executable path, text, key sequence, timeout, menu path, item name, window title, screenshot path, etc. |
| `clickRegion` | string | header dropdown operations | Optional header click target. Defaults to `LowerRight`. |
| `itemRegion` | string | header dropdown item selection | Optional list-item click target. Defaults to `LeftCenter`. |
| `index` | number | combo/grid operations | Zero-based combo item index or grid row index. |
| `columnIndex` | number | grid operations | Zero-based grid column index. |
| `limit` | number | list/debug operations | Optional maximum number of returned items. |
| `includeDesktopDescendants` | boolean | `listwindows` | When true, also scans desktop descendants. Defaults to false. |

### Response formats

Success:

```json
{ "success": true, "value": { } }
```

Client/request errors return `400`; not-found/operation failures return `404`; unexpected
failures return `500`. Error responses include a message and may include a failure
screenshot path if screenshot capture is configured:

```json
{
  "success": false,
  "value": null,
  "error": "'operation' is required.",
  "screenshotPath": "C:\\...\\failure.png"
}
```

### Locator reference

All `locator` and `locator2` objects support these properties:

| Property | Description |
|---|---|
| `mode` | Optional mode hint. Menu operations can use `"logical"` for logical menu traversal. |
| `automationId` | UIA AutomationId property. |
| `name` | UIA Name property, exact match. |
| `className` | UIA ClassName property. |
| `controlType` | UIA control type string, e.g. `Button`, `Edit`, `ComboBox`, `MenuItem`. |
| `xpath` | XPath-style expression. When set, all other locator properties are ignored. |

XPath examples:

```text
//Button[@Name='OK']
//*[@AutomationId='myId']
//ComboBox[@Name='Status']/ListItem[@Name='Active']
//Edit[@AutomationId='search' and @Name='Search']
//ListItem[@Name='Item'][2]
```

### `/ui` operations

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
| `rightclick` | `locator` | Right-click an element. | `null` |
| `hover` | `locator` | Move mouse over an element. | `null` |
| `focus` | `locator` | Give keyboard focus to an element. | `null` |
| `type` | `locator`, `value` | Type text into an element. Date pickers use segmented date typing when applicable. | `null` or typed date metadata |
| `typedate` | `locator`, `value` | Type a WinForms `SysDateTimePick32` date as MM → DD → YYYY segments. | `{ "typed": true, "strategy": "date-segments", ... }` |
| `clear` | `locator` | Clear an editable element. | `null` |
| `sendkeys` | `value`; optional `locator` | Send text/key tokens. When `locator` is provided the element is focused first; when omitted, keys are sent to the currently focused element in the active window. | `null` |
| `scroll` | `locator` | Scroll element into view (UIA ScrollItem pattern). | `null` |
| `mousescroll` / `wheelscroll` | optional `locator`, optional `value` | Scroll the mouse wheel. `value` is the number of wheel clicks (positive = up, negative = down) or `"up"` / `"down"` (±3 clicks). Defaults to 3 clicks up. When `locator` is provided the cursor moves to the element first. | `null` |
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

Select all and delete (clear via keyboard):

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

#### `mousescroll` examples

Scroll down 5 clicks over a list:

```json
{
  "operation": "mousescroll",
  "locator": { "automationId": "myList", "controlType": "List" },
  "value": "-5"
}
```

Scroll up 3 clicks (default) at the current cursor position:

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
| `Ctrl+A` | Assistive | Right-click an element to open an action menu, then choose what to record/perform. Use this when you need deterministic locators, assertions, menu actions, grid/header dropdown actions, or dynamic menu paths. |
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
| `MenuPathClick` | `clicklogicalmenupath` |
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
