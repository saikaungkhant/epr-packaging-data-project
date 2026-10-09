Attribute VB_Name = "Module1"
Option Explicit

Sub FormatSubmission()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Submission")

    ' Header row: bold + light blue fill
    With ws.Range("A1:L1")
        .Font.Bold = True
        .Interior.ThemeColor = xlThemeColorAccent1
        .Interior.TintAndShade = 0.6
    End With

    ' Fit column widths
    ws.Columns("A:L").AutoFit

    ' Freeze top row (needs the sheet to be active)
    ws.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitColumn = 0
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True
End Sub


Function DoChecks() As Boolean
    Dim raw As Worksheet, subm As Worksheet, logWs As Worksheet
    Dim nUnmapped As Long, nNoKg As Long, nHdc As Long, nUnitW As Long
    Dim kgRaw As Double, kgSub As Double
    Dim msg As String
    Dim passed As Boolean, r As Long

    Set raw = ThisWorkbook.Worksheets("Raw_ALL")
    Set subm = ThisWorkbook.Worksheets("Submission")
    Set logWs = ThisWorkbook.Worksheets("Check_Log")

    With Application.WorksheetFunction
        ' 1. Any codes that didn't map?
        nUnmapped = .CountIf(raw.Range("U:U"), "UNMAPPED")

        ' 2. Included rows (no exclude_reason) with no final_kg
        nNoKg = .CountIfs(raw.Range("A:A"), "<>", _
                          raw.Range("AD:AD"), "", _
                          raw.Range("AE:AE"), "")

        ' 3. HDC problems
        nHdc = .CountIf(raw.Range("V:V"), "INVALID HDC MATERIAL") _
             + .CountIf(raw.Range("V:V"), "HDC NO UNITS")

        ' 4. Unit-weight flags not fixed by an override and not excluded
        nUnitW = .CountIfs(raw.Range("AB:AB"), "CHECK UNIT WEIGHT", _
                           raw.Range("AC:AC"), "", _
                           raw.Range("AD:AD"), "")

        ' 5. Totals: Raw_ALL vs Submission
        kgRaw = .Sum(raw.Range("AE:AE"))
        kgSub = .Sum(subm.Range("I:I"))
    End With

    ' Decide pass/fail once, use it twice
    passed = (nUnmapped + nNoKg + nHdc + nUnitW = 0) And (Abs(kgRaw - kgSub) < 500)

    ' Write headers the first time only
    If logWs.Range("A1").Value = "" Then
        logWs.Range("A1:H1").Value = Array("run_time", "unmapped", "no_kg", _
            "hdc_problems", "unit_weight_flags", "kg_raw", "kg_submission", "result")
    End If

    ' Find the first empty row and write this run
    r = logWs.Cells(logWs.Rows.Count, "A").End(xlUp).Row + 1
    logWs.Cells(r, 1).Value = Now
    logWs.Cells(r, 2).Value = nUnmapped
    logWs.Cells(r, 3).Value = nNoKg
    logWs.Cells(r, 4).Value = nHdc
    logWs.Cells(r, 5).Value = nUnitW
    logWs.Cells(r, 6).Value = kgRaw
    logWs.Cells(r, 7).Value = kgSub
    logWs.Cells(r, 8).Value = IIf(passed, "PASS", "FAIL")

    msg = "UNMAPPED codes: " & nUnmapped & vbNewLine & _
          "Included rows with no kg: " & nNoKg & vbNewLine & _
          "HDC problems: " & nHdc & vbNewLine & _
          "Unfixed unit-weight flags: " & nUnitW & vbNewLine & _
          "Raw_ALL kg: " & Format(kgRaw, "#,##0") & vbNewLine & _
          "Submission kg: " & Format(kgSub, "#,##0") & vbNewLine & vbNewLine & _
          IIf(passed, "ALL CHECKS PASSED", "SOMETHING NEEDS A LOOK")

    MsgBox msg, vbInformation, "EPR checks"
    DoChecks = passed
End Function


Sub RunChecks()
    DoChecks
End Sub


Sub CheckAndExport()
    If DoChecks() Then
        ExportSubmission
    Else
        MsgBox "Export stopped. Fix the failed checks first.", vbExclamation, "EPR export"
    End If
End Sub


Sub ExportSubmission()
    Const OUT_FOLDER As String = "C:\Users\saika\OneDrive\Training\valpak_epr_practice\VBA\"
    Dim subm As Worksheet, newWb As Workbook
    Dim lastRow As Long, fileName As String

    Set subm = ThisWorkbook.Worksheets("Submission")
    lastRow = subm.Cells(subm.Rows.Count, "A").End(xlUp).Row

    ' Copy VALUES only (no formulas) into a brand-new workbook
    Set newWb = Workbooks.Add
    newWb.Worksheets(1).Range("A1:J" & lastRow).Value = _
        subm.Range("A1:J" & lastRow).Value

    ' Save as CSV with today's date and time in the name
    fileName = OUT_FOLDER & "submission_2025_" & Format(Now, "yyyymmdd_hhmm") & ".csv"
    Application.DisplayAlerts = False
    newWb.SaveAs fileName:=fileName, FileFormat:=xlCSVUTF8
    newWb.Close SaveChanges:=False
    Application.DisplayAlerts = True

    MsgBox "Exported " & (lastRow - 1) & " rows to:" & vbNewLine & fileName, vbInformation
End Sub

