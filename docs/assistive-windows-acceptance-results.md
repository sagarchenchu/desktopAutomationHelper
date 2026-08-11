# Assistive Windows acceptance results

Fill this template only after genuine interactive validation on a Windows desktop.
If not executed, leave status as `NOT RUN — interactive Windows validation required`.

## Run metadata

| Field | Value |
|---|---|
| Tested commit SHA | |
| Windows version | |
| Application type (WPF / WinForms) | |
| Test date | |
| Tester | |
| Overall status | `NOT RUN — interactive Windows validation required` |

## Scenario results

| # | Scenario | Result (`PASS` / `FAIL` / `NOT RUN`) | Evidence / artifact paths | Notes |
|---|---|---|---|---|
| 1 | Launch simple WPF/WinForms test app | NOT RUN | | |
| 2 | Start recording; Ctrl+A | NOT RUN | | |
| 3 | Start Jira recording with `ABC-1234` | NOT RUN | | |
| 4 | Next-action BDD association | NOT RUN | | |
| 5 | Multiple-action BDD + finish | NOT RUN | | |
| 6 | Second window + rich Assistive actions | NOT RUN | | |
| 7 | Click that opens a dialog (source window) | NOT RUN | | |
| 8 | Click OK/Cancel that closes popup | NOT RUN | | |
| 9 | Close Window (title before close) | NOT RUN | | |
| 10 | Ctrl+S with concurrent `/record/status` | NOT RUN | | |
| 11 | Recording JSON + candidate pages + BDD map | NOT RUN | | |
| 12 | `/playback` unchanged behavior | NOT RUN | | |
| 13 | Ordinary Assistive without Jira/BDD | NOT RUN | | |
| 14 | Immediate `/record/start` then `/record/stop` before overlay fully displayed | NOT RUN | | |

### Scenario 14 expected checks

- Stop succeeds
- Exactly one primary `recording_*.json` created
- Lifecycle returns to Idle
- A new recording can start afterward
- No orphan overlay
- No partial artifact directory remains

## Notes

_
