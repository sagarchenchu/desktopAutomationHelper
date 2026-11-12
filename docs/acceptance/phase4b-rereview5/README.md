# Phase 4B acceptance evidence (fifth re-review)

Committed so PR reviewers can audit without the local `artifacts/` tree.

| Run | Folder / file |
|---|---|
| Dry-run | [`20261112T024749897Z-09fc64db/`](./20261112T024749897Z-09fc64db/) · [`dry-run.json`](./dry-run.json) |
| Navigate | [`20261112T024751178Z-f74fe604/`](./20261112T024751178Z-f74fe604/) · [`navigate.json`](./navigate.json) |
| Modal + HWND backtrack | [`20261112T024837101Z-d4723e31/`](./20261112T024837101Z-d4723e31/) · [`modal-second.json`](./modal-second.json) |
| Non-modal second-window crawl | [`20261112T024849579Z-1206ebc4/`](./20261112T024849579Z-1206ebc4/) · [`second-window.json`](./second-window.json) |

Policies: [`acceptance.policy.json`](./acceptance.policy.json), [`modal-second.policy.json`](./modal-second.policy.json), [`second-window.policy.json`](./second-window.policy.json).

Second-window proof: `Open Second` → `outcome=newWindow` (not staleRef-only), then `Ping Second` → `outcome=success` on `DA Explorer Second`, then HWND `backtrack:close` via WindowPattern.
