# Phase 4A Windows acceptance

Recorded on the self-hosted Windows worker against branch
`feature/agent-phase-4a-semantic-observation` (draft PR #140; do not merge).

- Final tip SHA: e929f28025ba1edc1a612f544afab61efdc8b1f6
- Product/acceptance code tip: `74a6e40764aec2b96475ae1735f79aba3f461ee4`

## Targets

| Target | App | Evidence |
|---|---|---|
| WPF (required) | `tools/DesktopAutomationWpfSmokeApp` (`DA WPF Smoke`) | observe main + child dialog |
| WinForms | harness `DesktopAutomationTestApp` from PR #138 worktree (`DA Test Harness`) | observe main, child window, modal dialog |
| Child dialog | WinForms `DA Sample Dialog` (modal) and WPF `DA WPF Child Dialog` | observed while open |

Harness PR #138 alone is **not** sufficient: WPF observation uses the committed
WPF smoke app on this branch.

## Results

| Check | Result |
|---|---|
| Observe WPF window | PASS - DesktopAutomationWpfSmokeApp, sessionId=1, button/textbox/combo/list/disabled/menu |
| Observe WPF child dialog | PASS - title DA WPF Child Dialog, 10 nodes |
| Observe WinForms window | PASS - DesktopAutomationTestApp, sessionId=1, menu/button/textbox/combo/DataGrid/SampleGrid/disabled |
| Observe WinForms child window | PASS - title DA Child Window |
| Observe WinForms modal dialog | PASS - title DA Sample Dialog, isModal=true |
| No UI mutation from observe-only | PASS - EventLog getvalue unchanged across observe |
| Process/session/window scope | PASS - processName/processId/sessionId match the focused app |

### Snapshot hashes (acceptance run)

- WPF main: `99f12a52932cc18c3278ad50e2d57fcf8fe88e33a12311932abc76605acbfae2`
- WPF dialog: `855d4b2adb00fe4bf19570d956bdf3351f21c06a90c07bcb3fdfe9a9f8791de3`
- WinForms main: `dde2b278ce0f8ccc3e285da88098ffee15adecbd2547efd5263c68592b573f9c`
- WinForms child: `581571f3d0f03f7df5b39d59abba40b36fb55cebf3e772438f067fb289fe95e9`
- WinForms dialog: `5e422a346f80160f937eb52a963d0de2630a351d7f1bc34bc46b27ec26e6fc78`

Raw JSON summaries live under agent-store docs/acceptance/phase4a-*-observe.json
and local artifacts/phase4a-acceptance/ (not required in git).

## Roadmap status

Windows acceptance for Phase 4A **is required** before merge (WPF + WinForms when
available + actual child dialog). It is **not** optional / "not blocking."
