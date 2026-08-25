Attribute VB_Name = "HtmlUpdateModule"
' Macro for a manual "Update HTML Report" button on the report workbook.
' Not auto-installed (this project deliberately avoids macros/VBA elsewhere -
' see the HTML_Report_Generator memory notes) - import this yourself:
'   1. Alt+F11 to open the VBA editor
'   2. File > Import File... > select this .bas file
'      (or: Insert > Module, then paste the Sub below into the new module)
'   3. Save the workbook as .xlsm (macro-enabled) if it isn't already
'   4. Insert a Forms button (Insert > Illustrations > Shapes, or the Forms
'      button control) and set its "Assign Macro" to RunHtmlUpdate
'
' Requires no special Trust Center setting to USE (only auto-generating this
' code via script automation needed "Trust access to the VBA project object
' model", which is off by default - manually pasting/importing it yourself
' works regardless of that setting).

' PREVIEW / TESTING STAGE: -Publish is deliberately left off below, so this
' button only refreshes the TEMP working file, never the approved dated
' snapshot. Once the TEMP output has been manually checked and approved for
' a few months, add " -Publish" back onto the end of psCmd to make this
' button finalize the real monthly deliverable too.
Sub RunHtmlUpdate()
    Dim scriptPath As String
    Dim psCmd As String

    scriptPath = "C:\Users\User\Desktop\Global Purchase Report Claude\Global Purchase Report RPA\HTML_Report_Generator\Generate-PurchaseReportHtml.ps1"

    psCmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & scriptPath & """"

    ' Shell() is non-blocking - control returns to Excel immediately while
    ' PowerShell runs in the background (confirmed safe in sandbox testing:
    ' it opens its own independent Excel COM instance and does not conflict
    ' with this workbook staying open).
    Shell psCmd, vbNormalFocus

    MsgBox "HTML preview generation started in the background." & vbCrLf & _
           "This takes about a minute - check REPORT\FIN_Global_Purchase " & _
           "-Report_html_TEMP.html afterwards." & vbCrLf & vbCrLf & _
           "(Preview only - this does NOT publish the approved monthly report.)", _
           vbInformation, "Update HTML Preview"
End Sub
