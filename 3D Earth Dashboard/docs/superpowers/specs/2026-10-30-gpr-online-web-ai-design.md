# Design: Global Purchase Report — Online WEB + DATA with Working AI Assistant

Date: 2026-10-30 · Status: approved concept, pending plan · Owner: Anthony (anthonymankaho?)

## Intent and the why

Today the 3D Earth dashboard (`Global_Purchase_3D_Earth_TEMP.html`) is shared as a file
(TeamPlayer/Drive/email). Recipients get a fully working report but an **offline AI
Assistant**, because the panel calls the loopback relay `127.0.0.1:8765` — always the
viewer's own machine. Anthony wants a hosted online version where **the AI assistant works
for every recipient**, refreshed monthly by pushing new data, with nothing for teammates
to install.

## Goals

1. One public-web URL (reachable only by emailed links) serving the dashboard + a working
   Qwen chat assistant for all named recipients.
2. Monthly refresh = push new DATA only; the WEB shell rarely changes.
3. Zero credentials in any shared file or repo. Qwen bearer lives only as a Cloudflare
   Worker secret.
4. Bound runaway use with a per-app monthly token cap; fail closed with a friendly message.
5. Keep the existing file:// TEMP workflow and the :8765 loopback relay untouched.

## Non-goals

- Per-user login, accounts, or Cloudflare Access (audience spans 6 company domains +
  personal Gmail — single-tenant Google gating fails; decided).
- Mobile app, analytics, editing capabilities, write-back to GitHub from the browser.
- Replacing TeamPlayer/Drive distribution — the online link is an *additional* channel.

## Anti-goals (what makes this a failure even if it "works")

- API key visible in page source, repos, logs, or network responses.
- A stranger who guesses a URL draining Anthony's Qwen credits unbounded.
- Chat transcripts riding along in shared/saved copies.
- The heavy monthly PowerShell rebuild pipeline (Main.xaml untouched elsewhere; this
  project's PS scripts) breaking because the online variant mutated its outputs.

## Constraints (facts, verified 2026-10-30)

- **Cloudflare Free**: 100k requests/day **shared account-wide**, hard stop (never bills,
  no card); ≤100 Workers; Pages effectively unlimited static traffic; KV 100k reads +
  1k writes/day. Team-scale chat traffic ≈ hundreds/day → rounding error.
- **Qwen plan**: credit-metered, **45,500 credits/month**, auto-renewal. Observed blend
  ≤66.6 credits / 1M tokens (input-dominated, cache hits). qwen3.8-flash list: $0.15/M in,
  $0.47/M out; 1 credit ≈ $0.0025; thinking tokens bill as output.
- Anthony's own burn pace (2,379 credits/day, 10-16→10-30) will likely hit the plan wall
  ~2026-11-04, before renewal. Accepted: every app (MPL, EIC, web) then shows the same
  "quota exhausted until renewal" state.
- **GitHub dual-account gotcha**: anthonymankaho vs AnthonyManhk share one `gh` credential —
  repo owner must be confirmed before first push (question for Anthony).
- **Desktop is a leak path** (Desktop → OneDrive → private GitHub auto-sync at 13:00).
  Local clone dirs and any config containing the bearer must live OUTSIDE Desktop
  (use `C:\Users\User\GitHub-Clones\...`). The wrangler bearer token file likewise.
- Model standard: **qwen3.8-flash**, pinned server-side; the browser never supplies model,
  limits, or budgets.

## Decisions — as Anthony decided them

| # | Decision | Decided by |
|---|---|---|
| D1 | Two GitHub repos: **WEB** (dashboard shell + Worker code) and **DATA** (monthly purchase JSON) | Anthony |
| D2 | Cloudflare (Worker + hosting) as the online home | Anthony |
| D3 | **No passwords/logins.** Access control = the same philosophy as Drive: links emailed only to the 12 named recipients | Anthony |
| D4 | Quota bounded by a hard cap instead of identity: **0.3M tokens per calendar month, per AI app** (web = its own 0.3M, counted from provider `usage`) | Anthony |
| D5 | Planning mix assumption for budgeting: 20% input / 80% output (conservative; web reality is ~82% input, cheaper) | Anthony |
| D6 | Model pinned to **qwen3.8-flash** for all AIs | Anthony |

## Decisions — mine unless overruled (builder calls)

| # | Call | Rationale |
|---|---|---|
| B1 | One Cloudflare **Worker with static assets** serves the shell, `/api/health`, `/api/chat`, and the data JSON from a **same origin** (single subdomain; no custom domain) | Same-origin kills CORS entirely and lets one token-cookie gate all routes; cookie issues across pages.dev↔workers.dev avoided |
| B2 | Data flows: PS rebuild additionally emits `gpr-data-<period>.json` → `wrangler kv key put` (atomic swap) → Worker serves current key. DATA repo keeps the JSON as source of truth/version history | KV read per request is free-tier-safe (100k reads/day); git push alone can't reach Cloudflare without a deploy; KV = one-command monthly update |
| B3 | The "unguessable" part: a **random path segment per recipient cohort** (e.g. `/r/<32-hex>/`) issued at email time; Worker maps segment → viewer bucket, serves shell+data+chat only under a known segment. Unknown segment → 404, no fingerprint of existence | Anthony declined real auth; obscurity+caps is the agreed lock; per-cohort segments let one person's leak be revoked by deleting one mapping line (revoke = that cohort gets 404s, others unaffected) — this restores per-recipient blocking for free |
| B4 | Quota counters in **KV** keyed `quota:<month>:gpr-web`, cap 0.3M tokens; warn ≥80% in-panel; fail closed at 100%; per-bucket min-interval between messages; upstream 401/429 → "used up until renewal" message | Same semantics as the loopback relay, so both variants behave identically |
| B5 | One HTML source tree; build stamps `RELAY`: file:// variant → `http://127.0.0.1:8765`; web variant → `/api`. Panel code otherwise unchanged (already allowlisted context, `textContent`, in-memory history) | The panel was designed for exactly this swap |
| B6 | Repos: WEB and DATA **private**; DATA contains real vendor names/amounts | Matches Drive-copy exposure, not more |
| B7 | Worker secret via `wrangler secret put QWEN_TOKEN` from the (off-Desktop) clone; token value never printed, never logged by Worker (metadata-only logs) | Skill rule |
| B8 | Chat identity: per-cohort bucket name (e.g. `gpr-team-a`) for quota keys — not real names | Enough for blocking/revocation, no personal data |

## What is left to the plan/build stage

- Exact JSON schema split out of the current embedded-data island (rows + option sets +
  summary fields the KPI/chart layer uses), versioned by `period` field.
- Where the build step lives: extend the existing PS rebuild to emit JSON + a small
  `deploy-web.ps1` (wrangler deploy + kv put), both under `GitHub-Clones\GPR-Web\`.
- wrangler project scaffold, directory layout of the two repos, CI: none (manual push;
  Cloudflare Git-integration optional later).
- Mobile/Safari check of the hosted page (existing verification habits apply).

## Verification (must all pass before declaring done)

1. Grep WEB + DATA repos and the built page for the bearer prefix → zero hits.
2. Fresh headless-Chromium load of the online URL → real chat reply (provider `usage`
   returned, counted into the KV quota).
3. Unknown/garbled path segment → 404, no data JSON, no chat.
4. Kill/expire the cohort segment → that link's panel and data both die; others live.
5. Simulate quota exhaustion (temp cap = tiny) → friendly fail-closed message, dashboard
   still browsable, log shows no prompt contents.
6. `<img onerror>`-style model reply renders inert.
7. KV write counters/day and request counters observed against free limits after first
   real month.
8. file:// TEMP sibling re-tested unchanged: still degrades offline without relay.

## Open questions for Anthony (blocking for first push)

- **Q1: which GitHub account owns the two repos — anthonymankaho or AnthonyManhk?**
  (dual-account `gh` gotcha: wrong active credential looks like "Repository not found")

## Known accepted risks

- Forwarded link ⇒ access moves with it (same as a forwarded Drive copy today).
- Plan wall before renewal stops ALL bearer-sharing apps at once (observed pace ⇒ ~Nov 4;
  monitor with the Qwen Usage VS Code extension).
- Cloudflare Free subdomains are unbranded URLs; recipients may find them less
  "official-looking" than a TeamPlayer path.
