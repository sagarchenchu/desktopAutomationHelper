# Phase 8 — Windows acceptance (template)

Keep PR #148 draft until this document records PASS for both UIA and custom-drawn fixtures.

Interactive desktop required (no headless).

| # | Scenario | Expected | Result | Notes |
|---|---|---|---|---|
| 1 | Debug Mode start/stop + Ctrl+D on WinForms test app | Highlight + inspect without click | PENDING | |
| 2 | Debug Mode on WPF smoke app | Same as above | PENDING | |
| 3 | Candidate locator unique verify | Unique OK; ambiguous rejected | PENDING | |
| 4 | Yellow highlight does not intercept clicks | Click reaches underlying control | PENDING | |
| 5 | CustomDrawnCanvas vision path (Enabled=true) | Candidate with confidence+bounds; no auto-promote | PENDING | |
| 6 | Vision disabled by default | No vision candidates | PENDING | |
| 7 | Stale/DPI/layout change | Coordinate action denied | PENDING | |
| 8 | Sensitive screenshot | Not persisted by default | PENDING | |

Overall: **PENDING**
