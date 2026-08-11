# Assistive Jira / BDD Recording

Deterministic Assistive Mode enhancement that associates optional Jira keys and BDD statements with recorded actions, then exports candidate Page Object files and a BDD-to-action map on stop.

This feature does **not** call Jira APIs, interpret BDD with AI/NLP, auto-promote page objects, or generate Phase 2 plans.

## User workflow

1. Start recording with `POST /record/start` (unchanged).
2. Press **Ctrl+A** to enter Assistive Mode.
3. Right-click a target application element.
4. Open **Jira / BDD Recording**.
5. Choose **Start Jira Recording…** and enter a key such as `ABC-1234`.
6. Optionally choose **Take BDD Statement (Next Action)…** or **Take BDD Statement (Multiple Actions)…**.
7. Perform Assistive menu actions as usual.
8. For multiple-action BDD, choose **Finish Current BDD Statement** when the group is complete.
9. Press **Ctrl+S** (or `POST /record/stop`) to export.

Entering Jira/BDD text never creates a `RecordedAction`. Only a later successful Assistive action receives the association.

## Menu labels

| Item | When enabled |
|------|----------------|
| Start Jira Recording… | Before a Jira key is set |
| Current Jira: … | Always shown (disabled) |
| Take BDD Statement (Next Action)… | After Jira key is set |
| Take BDD Statement (Multiple Actions)… | After Jira key is set |
| Finish Current BDD Statement | After Jira key is set |
| Cancel Pending BDD Statement | After Jira key is set; only succeeds if the group has no actions yet |

## Next-action vs multiple-action BDD

- **Next Action:** the next successfully recorded Assistive action receives the BDD group, then the group ends.
- **Multiple Actions:** every subsequent Assistive action shares the same `groupId` until **Finish Current BDD Statement**.
- Cancel removes an unused pending group only. Associations already written into events are never deleted.
- A new BDD statement cannot replace an active or pending group; **Finish** or **Cancel** is required first.
- Drag-and-drop consumes BDD on the final `DragAndDrop` action, not when selecting the source.
- Application context-menu recording associates the final `MenuPathClick` / recorded action, not the internal right-click used to open the menu.
- Passive Mode actions never inherit Assistive BDD metadata.

## Overlay status

The overlay may show safe state only, for example:

`Jira ABC-1234 | BDD bdd-0001 armed for next action`

It never displays the full BDD statement (may contain business data / PII). Logs record group ID, character count, and event IDs only.

## Recording JSON fields

Additive optional fields on Assistive events:

- `eventId`, `sequence`
- `jiraKey`
- `bdd` (`groupId`, `statement`) — omitted entirely when absent
- `window` (`title`, `normalizedTitle`, `processId`, `nativeWindowHandle`)
- `pageId`, `objectRef`, `targetObjectRef`

Existing recordings without these fields continue to deserialize and play back. Playback ignores the new annotation fields and uses the same operation mapping as before.

### Example single-action event

```json
{
  "eventId": "evt-000001",
  "sequence": 1,
  "actionType": "doubleClick",
  "mode": "assistive",
  "jiraKey": "ABC-1234",
  "bdd": {
    "groupId": "bdd-0001",
    "statement": "And double click on ABC"
  },
  "window": {
    "title": "Welcome",
    "normalizedTitle": "welcome",
    "processId": 1234,
    "nativeWindowHandle": "0x001A04F2"
  },
  "pageId": "welcome",
  "objectRef": "welcome.abc",
  "element": {
    "name": "ABC",
    "automationId": "btnABC",
    "controlType": "Button"
  },
  "description": "Double Click on ABC Button"
}
```

## Artifact directory layout

```text
<outputPath>/
  recording_20260810_162000.json
  assistive-artifacts/
    ABC-1234/
      rec-20260810-162000-7f81c4a2/
        bdd-action-map.json
        page-objects/
          welcome.page.json
          print-dialog.page.json
```

Without a Jira key:

- The normal recording JSON is still written.
- Candidate page objects may be written under `assistive-artifacts/unassigned/<recordingId>/`.
- No BDD map is invented for a missing Jira key.

`RecordingExport.artifacts` summarizes sidecar paths and warnings when present; `exportedFilePath` behavior is unchanged.

## Page Object candidates

- One candidate file per distinct normalized window title used by Assistive actions.
- Same title reuses the same `pageId` within the recording session.
- Different titles create different pages.
- `state` is always `candidate`; `source.kind` is always `capture`.
- Never writes `automation/object-repository/repository.json` or `pages/`.
- Never auto-promotes to active/approved.
- Locators use Phase 3 fields only (no bounding rectangles / coordinates / SuggestedXPath).
- Candidate locators are validated against Phase 3 runtime controlType rules (not only JSON Schema). Unsupported types such as `HeaderItem` omit `controlType` when `automationId` is present; otherwise the element is listed under `unresolved`.
- Schema: `automation/schemas/page-object.schema.json`.

## BDD action map

- Schema: `automation/schemas/bdd-action-map.schema.json`
- Groups by `groupId` (not statement text), so two separate “And click OK” statements remain distinct.
- Actions without BDD appear in `unmappedEventIds`.
- Does not duplicate typed values or secrets.

## Window title handling

- Page IDs are derived from Unicode-normalized, lowercased, kebab titles (`^[a-z][a-z0-9-]{0,63}$`).
- Digit-leading titles receive a `p-` prefix.
- Long titles / collisions use a deterministic short SHA-256 suffix.
- HWND is diagnostic only and never used as a page ID or locator.
- Dynamic title normalization beyond this phase is a known limitation / future enhancement.

## PII / security

- Runtime recordings, maps, and candidates may contain PII and remain gitignored.
- Do not log full BDD statements, typed values, or credentials.
- Artifact paths are containment-checked under the recording output directory.
- Symlink / junction / reparse-point escapes under `assistive-artifacts` are rejected.
- Sidecars are staged then renamed into place; primary recording JSON is replaced atomically.
- Export is single-flight (`NotStarted` → `InProgress` → `Completed`); concurrent `/record/status` waits for or reuses the same export.
- Recording lifecycle is `Idle` → `Active` → `Stopping` → `Idle`; starts are rejected for the whole stop/export window, not only after export `InProgress`.
- `POST /record/stop` claims `Stopping` and runs export itself (does not depend on the overlay existing or closing).
- Failed primary export keeps `Stopping` until retry succeeds (session cannot be erased by a new start).
- A new recording cannot start while a previous export is still in progress.
- Sidecar creation failures are reported separately from primary-summary rewrite failures (and do not claim sidecars were written when they were not).
- Partial staging directories are cleaned on failure.

## Manual Windows acceptance checklist

Unit tests do **not** exercise the real overlay/menu workflow. After merges, run this interactive checklist on Windows.
Record results in [`docs/assistive-windows-acceptance-results.md`](assistive-windows-acceptance-results.md).

1. Launch a simple WPF/WinForms test app.
2. Start recording; press Ctrl+A.
3. Start Jira recording with `ABC-1234`.
4. Arm a next-action BDD; double-click a button; confirm association on that action only.
5. Arm a multiple-action BDD; perform two actions; finish; perform another action without BDD.
6. Switch to a second titled window; perform Type, Type-and-Select, table, popup, Switch Window, and dropdown actions; confirm distinct page files.
7. **Click that opens a dialog** — confirm the recorded page/window is the source window, not untitled.
8. **Click OK / Cancel that closes a popup** — confirm the recorded page/window is the popup title captured before close.
9. **Close Window** — confirm window title was captured before close.
10. Press Ctrl+S while concurrently polling `/record/status` (or `/record/actions`) and confirm a single artifact directory with no failure warning from a raced export.
11. Confirm recording JSON, one candidate page per title, and BDD map grouping.
12. Replay via `/playback` and confirm behavior is unchanged.
13. Repeat without Jira/BDD and confirm ordinary Assistive recording still works.
14. **Immediate stop before overlay is fully displayed** — call `POST /record/start`, then immediately call `POST /record/stop` before the overlay is fully shown. Confirm: stop succeeds; exactly one primary `recording_*.json` is created; lifecycle returns to Idle; a new recording can start afterward; no orphan overlay and no partial artifact directory remains.

## Out of scope

- Jira REST / reading BDD from Jira
- AI interpretation / Copilot / self-healing / auto locator repair
- Automatic promotion to active page objects
- Full object-tree capture (future Debug Mode)
- BDD-to-Phase-2-plan compilation
- Screenshots / HTML reports / notifications
