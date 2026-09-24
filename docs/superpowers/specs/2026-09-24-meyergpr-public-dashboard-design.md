# MEYERGPR: Public Dashboard Mirror for Global Purchase Report

## Summary / 摘要

Joe (senior management) can't open the monthly interactive HTML report as an email/Slack attachment — his client shows no download option. The fix: publish the finalized monthly HTML report to a **new, separate, public GitHub repository** (`anthonymankaho/MEYERGPR`) via GitHub Pages, so recipients open it as a browser link instead of a file attachment. A lightweight client-side login screen gates casual/accidental viewing (not determined attackers — see Non-Goals).

Joe（高階主管）沒辦法在 email/Slack 附件裡打開每月的互動式 HTML 報表 — 他的信箱沒有下載選項。解法：把定案後的每月 HTML 報表發布到一個**全新、獨立、公開的 GitHub repo**（`anthonymankaho/MEYERGPR`），透過 GitHub Pages 讓收件人用瀏覽器網址開啟，不用下載附件。前端登入畫面用來擋掉「誤傳/誤開」的情況（不是為了擋駭客 — 見下方「非目標」）。

**Origin**: user relayed a detailed implementation prompt drafted with Copilot's help (2026-09-24). This spec incorporates that prompt's requirements, adjusted per two decisions made during brainstorming: (1) stay on GitHub Free plan → hosting repo must be public; (2) therefore use a **new, separate** repo rather than publishing from the existing `global-purchase-report-claude` repo, to keep 1.3GB of confidential source data private.

**來源**：使用者轉述了與 Copilot 討論後的詳細實作提示詞（2026-09-24）。本規格納入該提示詞的需求，並依腦力激盪過程中的兩個決定調整：（1）維持 GitHub 免費方案 → host用的repo必須公開；（2）因此改用**全新、獨立**的repo，而非直接從現有的 `global-purchase-report-claude` repo發布，以保護1.3GB機密原始資料的私密性。

## Non-Goals / 非目標

- **Not real security.** This is a static site with client-side JS. Anyone with the URL can bypass the login via dev tools, view-source, or by disabling JavaScript, and read every row of the embedded dataset. The login screen and unguessable URL exist to prevent *accidental* forwarding/opening by non-target recipients — not to withstand a motivated attacker. This must be stated explicitly in both the new repo's README and (briefly) in an on-page note, per Copilot's own prompt requirement.
- **Not integrated into the existing monthly PowerShell pipeline automation.** Publishing to MEYERGPR is a manual, explicit step the user runs after finalizing the month's report — matching this project's established "review first, explicit go-ahead" pattern. No auto-trigger.
- **Not a backend, database, or paid auth service.** Confirmed out of scope per the original prompt.

**不是真正的安全機制。** 這是純前端JS的靜態網站。任何拿到網址的人都能用開發者工具、看原始碼、或關閉JavaScript繞過登入，讀到內嵌資料集的每一列。登入畫面和不易猜到的網址只是為了防止「非目標收件人」意外轉發/打開 — 不是為了擋住有心人士。這點必須明確寫在新repo的README，以及頁面上的簡短提示（照Copilot提示詞的要求）。
**不會整合進現有的每月PowerShell流程自動化。** 發布到MEYERGPR是使用者在當月報表定案後手動執行的明確步驟 — 符合這個專案一貫「先審查、明確下指令才執行」的模式，不會自動觸發。
**不是後端、資料庫、或付費驗證服務。** 依原始提示詞範圍排除。

## Decisions Made During Brainstorming / 腦力激盪期間的決定

| Decision | Choice |
|---|---|
| GitHub plan | Free (stay free) |
| Hosting repo | **New, separate** repo — not the existing `global-purchase-report-claude` |
| Repo name | `MEYERGPR` (user's own choice — not fully random, but not descriptive of contents either; acceptable given the stated accidental-exposure threat model) |
| Repo owner | `anthonymankaho` (same account as the source project) |
| Repo visibility | Public (required for Pages on Free plan) |
| Login credentials | `Meyer` / `Meyer` (Copilot's default, user confirmed keep) |
| Deployment method | GitHub Pages, deploy from `main` branch, root — no GitHub Actions (no build step needed) |
| Monthly update trigger | Manual — user runs a copy/commit/push step after `Generate-PurchaseReportHtml.ps1 -Publish` |

## Architecture / 架構

Two static HTML files, no build step, no backend:

```
MEYERGPR/                      (new public repo, root = Pages site root)
├── index.html                 ← login screen (new file, written from scratch)
├── report.html                ← the existing dashboard, copied in verbatim + one injected <script> block
├── robots.txt                 ← disallow all
├── .nojekyll                  ← disable GitHub Pages' default Jekyll processing
└── README.md                  ← what this repo is, security limitations, how it's updated monthly
```

### `index.html` (new file)

Plain HTML/CSS/vanilla JS, single file, no external dependencies (keep it simple and fast — this is the first thing every viewer sees).

Requirements (from the user's prompt, verbatim scope):
- Username + password fields, both required (no empty submit).
- Default: username `Meyer` (case-insensitive compare), password `Meyer` (case-sensitive compare).
- Login button + Enter-key submit.
- Show/Hide password toggle.
- Focus on username field on page load.
- Accessible labels, keyboard navigable.
- Responsive (desktop/tablet/mobile).
- On success: `sessionStorage.setItem('globalPurchaseReportAuthenticated', 'true')`, then `location.href = 'report.html'`.
- On failure: non-technical error message (e.g. "Username or password is incorrect."), no hint about which field was wrong.
- If `sessionStorage.globalPurchaseReportAuthenticated === 'true'` already on load, skip straight to `report.html` (so a refresh of a bookmarked `index.html` doesn't force re-login mid-session).
- No credentials in URL/query string/history. No console logging of username/password. No password in localStorage (sessionStorage only, and only a boolean flag — never the password itself).
- `<meta name="robots" content="noindex,nofollow">` in `<head>`.
- Visual style: simple, professional, loosely matches the report's own light-mode palette (doesn't need to match exactly — this is a 10-second gate, not a branded product).

### `report.html` (existing file + one injected snippet)

Source: the exact output of `Generate-PurchaseReportHtml.ps1 -Publish` for the month being published — copied byte-for-byte except for two additions:

1. **Auth guard**, injected as the very first `<script>` in `<head>`, before anything else loads:
   ```html
   <script>
     if (sessionStorage.getItem('globalPurchaseReportAuthenticated') !== 'true') {
       location.replace('index.html');
     }
   </script>
   ```
   This blocks direct navigation to `report.html` without a valid session (Copilot's requirement) — note the Non-Goals caveat: a user who disables JS entirely bypasses this, which is accepted per the threat model.

2. **Logout button**, injected into the existing page's header/nav area (exact insertion point decided during implementation, once the current `report.html` structure is inspected) — a small fixed-position button:
   ```html
   <button onclick="sessionStorage.removeItem('globalPurchaseReportAuthenticated'); location.href='index.html';">Logout</button>
   ```

3. `<meta name="robots" content="noindex,nofollow">` added to `<head>` alongside the existing meta tags.

No other change to `report.html`'s content, charts, filters, or data — this is strictly an unmodified copy plus the two snippets above.

### `robots.txt`

```
User-agent: *
Disallow: /
```

### `.nojekyll`

Empty file. Prevents GitHub Pages' default Jekyll build step from processing/mangling the large generated HTML (underscores, `_`-prefixed patterns Jekyll treats specially, etc.).

### `README.md` (new repo's own, distinct from the source project's README)

Per the original prompt's Part 7, documents: purpose, how to test locally, GitHub Pages settings used, the monthly update procedure, login/logout behavior, sessionStorage behavior, security limitations (the Non-Goals section above, in plain language), and rollback (see below). Does **not** include the actual password in a way an unauthorized viewer of this public repo could read it — actually, note: since Meyer/Meyer is already documented as "the default" in the *private* project's own SOP/README, and this new repo is itself public, the password **must not** appear in this repo's README at all (per the original prompt's explicit rule 12/"do not place the default password in README.md if the repo could be accessed by unauthorised users" — this repo IS public, so the rule applies). State only that credentials were shared separately (e.g. "ask Anthony for the login").

## Monthly Update Process / 每月更新流程

A new script in the *source* project (`Global Purchase Report Claude\Publish-MeyerGprMirror.ps1` or similar name, decided at implementation time):

1. Takes `-ReportMonth "yyyy-MM"`.
2. Locates that month's published HTML (`Global Purchase Report RPA\REPORT\FIN_Global_Purchase -Report_html_<YYYYMM>.html`) — reuses the existing dated-snapshot naming convention (see the "two REPORT folders" note already in the main README).
3. Copies it to a local clone/working copy of the `MEYERGPR` repo as `report.html`, injecting the two snippets from above (idempotent — always re-injects fresh, doesn't assume the prior month's copy already had them, in case the source file changed).
4. Commits and pushes — the user runs this manually, reviews the diff/output before confirming, per this project's established pacing preference. No auto-push.

This keeps the MEYERGPR repo itself very simple (just the static files, no PowerShell/report-generation logic duplicated into it) and keeps the actual report-building logic in exactly one place (the existing `HTML_Report_Generator`).

## Error Handling / 錯誤處理

- If the expected month's HTML file doesn't exist yet, the publish script throws with a clear message (matching the existing scripts' style: "run Generate-PurchaseReportHtml.ps1 -Publish first").
- If `report.html`'s structure changes in a way that breaks the injection point assumption (e.g. `<head>` tag missing — shouldn't happen since it's machine-generated, but defensive check anyway), the script aborts before writing rather than producing a broken page.
- Browser-side: if `sessionStorage` is unavailable (very old browser, or privacy mode blocking it in some browsers) — the login would appear to "not stick" across page loads. Acceptable degradation for this audience (internal recipients on standard corporate browsers); not engineering around it.

## Testing / Validation / 測試與驗證

Local, before every push (no CI/automated tests — matches this project's existing manual-verification pattern for Excel/report work):
1. Open `index.html` via `file://` — verify: empty submit rejected, wrong username rejected, wrong password rejected, correct credentials redirect to `report.html`, Enter key submits, show/hide password toggle works, username field has focus on load.
2. Open `report.html` directly via `file://` (simulating a bookmarked/shared direct link) without having logged in — verify it redirects to `index.html`.
3. Log in, confirm `report.html` loads with all existing functionality intact (charts render, filters work, drilldown works) — this is the existing report, so this step is mostly a sanity check that the injected snippet didn't break anything.
4. Click Logout — verify it returns to `index.html` and a subsequent direct `report.html` visit redirects again.
5. Refresh `report.html` mid-session (after logging in) — verify it does NOT redirect to login (session persists across a refresh in the same tab).

Deployment validation (once pushed, before sharing the link with anyone):
6. Visit the real `https://anthonymankaho.github.io/MEYERGPR/` URL fresh (private/incognito window) — confirm it lands on the login screen, not the report.
7. Confirm `robots.txt` and the `noindex` meta tags are present in the deployed pages (view-source on the live URL).

## Rollback / 回復方式

- **Unpublish immediately**: GitHub repo Settings → Pages → set source to "None" — takes the live site down within a few minutes, without deleting any files or history.
- **Full removal**: delete the `MEYERGPR` repository entirely. Since it never contains anything except the static dashboard files (no source data, no scripts of record), there's nothing to preserve — the source of truth stays in `Global Purchase Report Claude`.
- **Rotate the URL** (if the current one leaks further than intended): rename the repo — GitHub Pages URLs follow the repo name, so a rename immediately changes the public URL (old links break, which is the point).

## Open Items for Implementation-Time Decisions / 實作時再決定的細節

These don't block writing the plan, but are called out so the implementer (or Claude, during planning) inspects the real file before deciding:
- Exact byte offset/insertion point for the Logout button inside `report.html` (depends on the report's actual current header markup — inspect at implementation time, don't guess now).
- Whether `report.html`'s existing `<head>` already has any conflicting `<meta name="robots">` or CSP-like tag that needs merging rather than duplicating.
- Final visual polish of `index.html` (exact colors/fonts) — loosely match the report's existing light-mode palette per the original prompt, finalize during implementation.
