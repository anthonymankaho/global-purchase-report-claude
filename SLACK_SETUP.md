# Slack 一次性設定 / One-time Slack setup

這是**您需要自己在瀏覽器中完成的一次性設定**——我無法代勞，因為需要您（或公司 Slack 管理員）登入並在公司 workspace 建立/安裝一個 Slack App。完成之後，`New-MonthlySlackPost.ps1` 就能直接使用，不需要每月重新設定。

This is a **one-time setup you (or your company's Slack admin) do yourselves in the browser** — I can't do this part, since it requires logging into your company's Slack workspace and creating/installing an app there. Once done, `New-MonthlySlackPost.ps1` just works every month, no repeat setup needed.

> ⚠️ **可能需要 Slack 管理員協助**：公司 Slack 剛開始使用，如果您的帳號沒有「安裝 App 到 workspace」的權限，Step 3 可能需要請 IT/Slack 管理員協助安裝。
> ⚠️ **May need your Slack admin's help**: since the company Slack is brand new, if your account doesn't have permission to "install an app to the workspace," Step 3 may need IT/your Slack admin to do it.

## Step 1 — 建立 Slack App / Create a Slack App

1. 前往 https://api.slack.com/apps ，用您的公司 Slack 帳號登入。
   Go to https://api.slack.com/apps and log in with your company Slack account.
2. 點擊 **Create New App** → **From scratch**。
   Click **Create New App** → **From scratch**.
3. App Name 填 `Global Purchase Report Bot`（或任意名稱），選擇公司的 workspace。
   App Name: `Global Purchase Report Bot` (or anything), pick your company's workspace.

## Step 2 — 設定權限 / Configure permissions

1. 左側選單：**OAuth & Permissions**。
   Left menu: **OAuth & Permissions**.
2. 往下捲到 **Scopes → Bot Token Scopes**，新增：
   Scroll to **Scopes → Bot Token Scopes**, add:
   - `chat:write`（發訊息到 bot 已加入的頻道 / post to channels the bot has joined）
   - `chat:write.public`（可發到任何公開頻道，不用先手動邀請 bot 加入 / post to any public channel without manually inviting the bot first）
   - `files:write`（上傳並分享檔案附件，例如互動式 HTML 報表 / upload and share file attachments, e.g. the interactive HTML report）

## Step 3 — 安裝到 workspace / Install to workspace

1. 頁面最上方：**Install to Workspace** → 允許權限。
   Top of the same page: **Install to Workspace** → approve the permissions.
2. 安裝完成後，這個頁面會顯示 **Bot User OAuth Token**（開頭是 `xoxb-...`）。這個就是之後要用的密鑰，請先複製起來，稍後貼給我一次就好（我會加密儲存，不會存在對話紀錄裡）。
   After installing, this page shows a **Bot User OAuth Token** (starts with `xoxb-...`). This is the key the script will use — copy it, you'll paste it once (I'll encrypt and store it, not keep it in chat history).

## Step 4 — 決定要發到哪個頻道 / Decide which channel to post to

建一個新頻道（例如 `#global-purchase-report`）或選一個現有頻道，讓13位同事之後陸續加入。頻道名稱之後要傳給腳本。

Create a new channel (e.g. `#global-purchase-report`) or pick an existing one, for the 13 colleagues to join over time. You'll pass this channel name to the script.

## Step 5 — 執行一次性儲存腳本 / Run the one-time token-save script

```powershell
cd "C:\Users\User\Desktop\Global Purchase Report Claude"
.\Save-SlackToken.ps1
```

會提示貼上 Step 3 的 Bot Token，加密儲存到 `Credentials\slack_bot_token.enc`（已加入 `.gitignore`）。只需要做一次。

Prompts you to paste the Step 3 Bot Token, encrypts and saves it to `Credentials\slack_bot_token.enc` (already in `.gitignore`). Only needs to be done once.

## Step 6 — 每月使用方式 / Monthly usage

```powershell
.\New-MonthlySlackPost.ps1 -ReportMonth "2026-07" -Channel "global-purchase-report"
```

會發一則跟 Gmail 草稿內容一致的通知到指定頻道（純文字通知，不含檔案附件——原因同 Gmail：檔案太大）。

Posts a notice matching the Gmail draft's content to the given channel (plain text notice, no file attachment — same reason as Gmail: the file is too large).
