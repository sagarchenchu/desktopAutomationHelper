# Phase 4B acceptance evidence (sixth re-review)

Committed so PR reviewers can audit without the local `artifacts/` tree.

| Run | Folder / file |
|---|---|
| Dry-run | [`20261112T060542511Z-5b049e2f/`](./20261112T060542511Z-5b049e2f/) · [`dry-run.json`](./dry-run.json) |
| Navigate | [`20261112T060544139Z-9d62233f/`](./20261112T060544139Z-9d62233f/) · [`navigate.json`](./navigate.json) |
| Modal + HWND backtrack | [`20261112T060740200Z-f6a6c903/`](./20261112T060740200Z-f6a6c903/) · [`modal-second.json`](./modal-second.json) |
| Non-modal second-window crawl | [`20261112T060745583Z-8ad995de/`](./20261112T060745583Z-8ad995de/) · [`second-window.json`](./second-window.json) |

Policies: [`acceptance.policy.json`](./acceptance.policy.json), [`modal-second.policy.json`](./modal-second.policy.json), [`second-window.policy.json`](./second-window.policy.json).

Sixth-review proofs:
- Modal: `Open Dialog` → `modalDialog`, then `backtrack:close:0x…` via WindowPattern.
- Second window: `Open Second` → `newWindow` (not staleRef-only), `Ping Second` → `success`, HWND WindowPattern close.
- Switch paths pass `hwnd`+`processId`; frontier candidates carry target HWND/PID.
