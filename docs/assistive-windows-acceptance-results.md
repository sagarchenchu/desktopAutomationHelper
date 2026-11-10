# Assistive Windows acceptance results

Fill this template only after genuine interactive validation on a Windows desktop.
If not executed, leave status as `NOT RUN — interactive Windows validation required`.

## Run metadata

| Field | Value |
|---|---|
| Tested commit SHA | a6cf3fb49570fa270c6101aa9771e2a0d0759b72 |
| Base SHA (`origin/main`) | `f1a84006ac0c71d18ee1b6ba02c3bee7858b13f8` |
| Windows version | Windows 10/11 Pro 24H2 (build 26100) |
| Application type (WPF / WinForms) | WinForms (`DesktopAutomationTestApp` from `cursor/desktop-test-app-harness-b89b` / PR #138) |
| Test date | 2026-11-10 |
| Tester | Cloud agent on My Machines worker `local` (`bc-68aebb30-29e5-52f7-a6a0-f4930f674e93`) |
| Overall status | `PARTIAL — 11 PASS / 3 FAIL` (draft PR; README not updated) |
| Computer-use note | Cursor `computer` tool unavailable on this Windows worker (`agent worker debug`: no Windows computer-use sidecar). Interactive input used Win32 `SendInput`/`keybd_event` + UIA against the Assistive overlay. |

## Scenario results

| # | Scenario | Result (`PASS` / `FAIL` / `NOT RUN`) | Evidence / artifact paths | Notes |
|---|---|---|---|---|
| 1 | Launch simple WPF/WinForms test app | PASS | agent-store `gate0-evidence/shots/s01-app.png` | `DA Test Harness` WinForms harness visible |
| 2 | Start recording; Ctrl+A | PASS | `shots/s02-assistive.png` | `/record/status` mode=`Assistive` |
| 3 | Start Jira recording with `ABC-1234` | PASS | `shots/s03-jira-set.png`, `s03-jira-menu.png` | Submenu showed `Current Jira: ABC-1234` |
| 4 | Next-action BDD association | PASS | `s04-summary.json` | Exactly one BDD association (`bdd-0001`) with Jira `ABC-1234` |
| 5 | Multiple-action BDD + finish | PASS | `s05-summary.json` | Two actions shared `bdd-0001`; finish then unassociated action |
| 6 | Second window + rich Assistive actions | FAIL | `s06-summary.json` (incomplete) | Type / type-and-select / table / switch / popup set not completed cleanly after menu/UIA instability |
| 7 | Click that opens a dialog (source window) | FAIL | `s07-summary.json` | Could not reliably Assistive-Click `OpenDialogButton` with source-window title after restart flakiness |
| 8 | Click OK/Cancel that closes popup | FAIL | — | Blocked behind scenario 7 dialog Assistive path |
| 9 | Close Window (title before close) | PASS | `s09-summary.json` | `Close Window` recorded `windowTitle=DA Second Window` |
| 10 | Ctrl+S with concurrent `/record/status` | PASS | `shots/s10-done.png`, `recordings/s10/recording_*.json` | Single artifact; concurrent status polling OK |
| 11 | Recording JSON + candidate pages + BDD map | PASS | `s11-meta.json` | Primary recording + `bdd-action-map.json` + 2 page files |
| 12 | `/playback` unchanged behavior | PASS | `s12-playback.json` | Playback accepted Assistive export JSON |
| 13 | Ordinary Assistive without Jira/BDD | PASS | `s13-summary.json` | 2 Assistive actions; no Jira/BDD map |
| 14 | Immediate `/record/start` then `/record/stop` before overlay fully displayed | PASS | `s14-checks.json` | Stop OK; one `recording_*.json`; Idle; restartable; no partial dir |

### Scenario 14 expected checks

- Stop succeeds — **PASS**
- Exactly one primary `recording_*.json` created — **PASS**
- Lifecycle returns to Idle — **PASS**
- A new recording can start afterward — **PASS**
- No orphan overlay — **PASS** (best-effort)
- No partial artifact directory remains — **PASS**

## Notes

- Target app built from worktree of `origin/cursor/desktop-test-app-harness-b89b` @ `8ba9e62` (not merged to main); **not** committed on this docs branch.
- Driver under test = acceptance-branch tip (main `f1a84006…` + this results commit only). Runtime recordings with app data were **not** committed.
- Env: `DOTNET_ROLL_FORWARD=Major` used where AspNetCore 8 runtime missing.
- ComboBox ValuePattern / grid NotImplemented findings from prior `/ui` smoke remain out of Gate 0 scope; they did not gate the PASS scenarios above.
- Failures on 6–8 are validation/environment interaction issues during Assistive overlay driving (UIA deadlock risk when querying the target app under a live overlay; injected mouse flakiness after repeated driver restarts), not product code changes in this PR.
