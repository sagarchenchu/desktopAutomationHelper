# Phase 8 - Windows acceptance

- Tip tested (GitHub-resolvable): [`28b8712db032e867bb8433f1f2bd69161d85004d`](https://github.com/sagarchenchu/desktopAutomationHelper/commit/28b8712db032e867bb8433f1f2bd69161d85004d) (`28b8712`)
- Review tip (CI + inspect re-smoke): [`65a332695d24489e492efd8417e0f3abd5445871`](https://github.com/sagarchenchu/desktopAutomationHelper/commit/65a332695d24489e492efd8417e0f3abd5445871) (`65a3326`)
- Branch: `feature/driver-phase-8-debug-vision`
- PR: #148 (**draft / READY FOR REVIEW** — not merged)
- When: 2026-11-13T18:10:05.5397620+00:00
- Overall: **PASS** (S1–S9 on tip `28b8712`)
- Note: S1–S9 PASS on `28b8712`. During that run, live `/debug/inspect` HTTP 500ed on `IntPtr` `WindowHandle` (scenarios 1–4 still PASS via on-disk `candidate.json`). Fixed on `65a3326` via hex `windowHandle` converter; `/debug/inspect` HTTP re-smoke **OK**. #149 closed; evidence owned by #148.

| # | Scenario | Expected | Result | Evidence |
|---|---|---|---|---|
| 1 | Debug Mode start/stop + inspect on WinForms test app | Highlight + inspect without click | PASS | `phase8-01-winforms-debug-result.json` |
| 2 | Debug Mode on WPF smoke app | Same as above | PASS | `phase8-02-wpf-debug-result.json` |
| 3 | Candidate locator unique verify | Unique OK; no auto-promote | PASS | `phase8-03-candidate-verify-result.json` |
| 4 | Yellow highlight does not intercept clicks | Click reaches underlying control | PASS | `phase8-04-highlight-passthrough-result.json` |
| 5 | CustomDrawnCanvas vision path (Enabled=true) | Candidate with confidence+bounds; no auto-promote | PASS | `phase8-05-customdrawn-vision-result.json` |
| 6 | Vision disabled by default | No vision candidates | PASS | `phase8-06-vision-default-off-result.json` |
| 7 | Stale/DPI/layout change | Coordinate action denied | PASS | `phase8-07-stale-dpi-guards-result.json` |
| 8 | Sensitive screenshot | Not persisted by default | PASS | `phase8-08-sensitive-screenshot-result.json` |
| 9 | Hover probe | Mouse probe, not UIA capability | PASS | `phase8-09-hover-probe-result.json` |