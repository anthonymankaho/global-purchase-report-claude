# PowerShell 執行時間報告 / PowerShell Runtime Report

> ⚠️ **誠實說明**：本次月結（2026-07）大部分腳本是您在自己的 PowerShell 視窗獨立執行的，我沒有在旁邊即時計時，
> 所以下表大多數「執行時間」欄位是**沒有實測數據**的，標成「未計時」，不是我編造一個數字充數。
> 唯一有實測數據的是 `Publish-OutsideReport.ps1`（我這次有全程監控）。
>
> ⚠️ **Honest note**: most scripts in this month's (2026-07) run were executed independently in your own
> PowerShell window — I wasn't timing them live, so most "Runtime" cells below are genuinely **not measured**,
> marked as such rather than a made-up number. The only script with real measured data is
> `Publish-OutsideReport.ps1` (which I monitored throughout this run).

## 本次月結（2026-07）各腳本執行時間 / This month's (2026-07) script runtimes

| 順序 | 腳本 Script | 執行時間 Runtime | 備註 Notes |
|---|---|---|---|
| 1 | `Build-GlobalPurchaseReport.ps1` | 未計時 Not measured | 您自行執行，未同步觀察 Run independently, not observed live |
| 2 | `Get-PenExchangeRate.ps1` | 未計時 Not measured | 已確認成功（Exchange Rate M欄=2.2973），但未計時 Confirmed successful (Exchange Rate col M = 2.2973), but not timed |
| 3 | 手動 AA2:AB2 Manual AA2:AB2 | 未計時（人工步驟）Not measured (manual step) | — |
| 4 | `Complete-FormulaColumns.ps1` | 未計時 Not measured | 已確認成功（AC:AJ 已填值）Confirmed successful (AC:AJ populated) |
| 5 | `New-ReportAuditSummary.ps1` | 未確認是否執行過 Unconfirmed whether run | — |
| 6 | `Publish-OutsideReport.ps1` | **實測 46+ 分鐘，且疑似卡在收尾階段 Measured 46+ min, appears stuck at shutdown** | 見下方詳細記錄 See detailed log below |

## `Publish-OutsideReport.ps1` 詳細時間記錄（本次唯一實測）/ Detailed timing log (the one real measurement)

| 時間點 Time | 經過分鐘 Elapsed | 觀察 Observation |
|---|---|---|
| 15:54:48 | 0.0 | 程序啟動 Process started |
| — | 15.1 | 仍在 Step 2/3（同步 OUTSIDETEMP.xlsx）Still on Step 2/3 (syncing OUTSIDETEMP.xlsx) |
| 16:11 | ~16.2 | OUTSIDETEMP.xlsx 存檔完成 + 交付檔已產出 OUTSIDETEMP.xlsx saved + deliverable produced |
| — | 25.6 | 進入 Step 3/3（VBA練習檔）Entered Step 3/3 (VBA practice file) |
| 16:27 | ~32.2 | VBA練習檔存檔完成，兩個檔案鎖定檔皆已消失 VBA practice file saved, both lock files gone |
| — | 36.7–46.4 | **CPU 時間幾乎不再增加**（9分鐘內僅增加不到1秒）——資料已安全存檔，但程序本身疑似卡在 `$excel.Quit()` 收尾階段 **CPU time essentially flat** (under 1s added over 9 min) — data is safely saved, but the process itself appears stuck at the `$excel.Quit()` shutdown step |

**結論 Conclusion：** 實際「同步資料＋重新整理樞紐分析表」的工作，大約在 **32 分鐘**內完成（15:54:48 → 16:27）。之後的時間是程序收尾卡住，不是還在處理資料。已在腳本加入手動計算模式（`$excel.Calculation = -4135`），下次執行預期會快很多。

The actual "sync data + refresh PivotTables" work finished in about **32 minutes** (15:54:48 → 16:27). Time after that is the process stuck on shutdown, not still processing data. Manual calculation mode (`$excel.Calculation = -4135`) has been added to the script — the next run should be significantly faster.

## 下次如何自己記錄真實時間 / How to record real timings yourself next time

每支腳本執行前後各加一行，執行完會自動印出花了多久：
Add one line before and after each script — it'll print the elapsed time automatically:

```powershell
$t0 = Get-Date
.\Build-GlobalPurchaseReport.ps1 -ReportMonth $ReportMonth
"{0:N1} 分鐘 / minutes" -f ((Get-Date) - $t0).TotalMinutes
```

把每支腳本量到的實際分鐘數填回上面的表格，這份報告就能逐月累積成真正的效能基準線。
Fill the actual minutes you measure back into the table above — over a few months this becomes a real performance baseline.
