# The VBA code, explained

## Why VBA, and why last

Valpak asked about VBA on the phone screen, and it was my weakest answer. But VBA was left to the very end of the project on purpose: **you can only automate a process you have already done by hand and understood.** Every macro below replaces something I did manually in Modules 4 to 6.

## Setting up

1. Desktop Excel (Microsoft 365 Personal). VBA does not run in Excel for the web.
2. **File > Options > Customize Ribbon**, tick **Developer**.
3. Save a macro-enabled copy: **File > Save As**, type **Excel Macro-Enabled Workbook (.xlsm)**, name `master_2025_vba.xlsm`. A normal .xlsx cannot store macros. The original file stays as a backup.
4. Open the editor with **Alt+F11**. The code lives in **Modules > Module1**, exported to [vba/EPR_Automation.bas](../vba/EPR_Automation.bas).

## What the macro recorder wrote

My first macro was recorded, not written. I formatted the Submission header while Excel watched. This is what it produced (slightly shortened, with the long file address cut):

```vba
Sub FormatSubmission()
    ChDir "C:\Users\saika\OneDrive\Training\valpak_epr_practice\VBA"
    ActiveWorkbook.SaveAs Filename:= _
        "https://d.docs.live.net/.../master_2025_vba.xlsm" _
        , FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
    Sheets("Submission").Select
    ActiveWindow.SmallScroll Down:=-40
    Range("A1:L1").Select
    Selection.Font.Bold = True
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1:L1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.249977111117893
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.599993896298105
    End With
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    Columns("A:J").Select
    Selection.Columns.AutoFit
End Sub
```

What reading it taught me:

| Line | Problem |
|---|---|
| ChDir and ActiveWorkbook.SaveAs | I started recording before saving, so it recorded my Save As. Every run would re-save over the file. **Dangerous.** |
| SmallScroll | I scrolled. Pure noise. |
| Select, then Selection | The recorder clicks everything first. Real VBA works on objects directly: faster, and the screen doesn't jump around. |
| Cells.Select and AutoFit | An extra AutoFit on the whole sheet that I didn't need. |
| Two Interior blocks | I picked a colour, then changed my mind. Only the second one matters. |

**Lesson: the recorder is a great way to discover the right object names, but the code always needs rewriting.**

## The complete module

This is the final code in the module, exactly as it runs.

```vba
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
    newWb.SaveAs Filename:=fileName, FileFormat:=xlCSVUTF8
    newWb.Close SaveChanges:=False
    Application.DisplayAlerts = True

    MsgBox "Exported " & (lastRow - 1) & " rows to:" & vbNewLine & fileName, vbInformation
End Sub
```

## How the pieces fit together

| Procedure | Type | What it does | How I run it |
|---|---|---|---|
| FormatSubmission | Sub | Bold, coloured header, column widths, frozen top row | Alt+F8 |
| DoChecks | Function | Runs 5 checks, logs the run, shows the result, **returns** True or False | Called by the two below |
| RunChecks | Sub | Just calls DoChecks | "Run EPR checks" button on the Checks sheet |
| CheckAndExport | Sub | Runs DoChecks; exports only if it returned True | Alt+F8 |
| ExportSubmission | Sub | Saves the submission as a dated CSV, values only | Called by CheckAndExport |

## Macro by macro

### Option Explicit

```vba
Option Explicit
```

**What it does:** VBA refuses to run if I use a variable I did not declare with Dim.

**Why:** it catches typos. Without it, if I typed `nUnmaped` (one p), VBA would silently create a new empty variable that is always 0, and my checks would say PASSED when they should not. Professional VBA always starts with this line. After any edit, **Debug > Compile VBAProject** checks the whole module.

### FormatSubmission

```vba
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets("Submission")
```

**Dim** declares a variable that holds a worksheet. **Set** points it at the Submission sheet. After that, `ws.` means "on the Submission sheet", with no clicking. **ThisWorkbook** means the file the code lives in, even if another workbook is active.

```vba
With ws.Range("A1:L1")
    .Font.Bold = True
    .Interior.ThemeColor = xlThemeColorAccent1
    .Interior.TintAndShade = 0.6
End With
```

**With ... End With** saves repeating `ws.Range("A1:L1")` on every line. Each line starting with a dot belongs to that range.

```vba
ws.Activate
ActiveWindow.FreezePanes = False
ActiveWindow.SplitRow = 1
ActiveWindow.FreezePanes = True
```

Freeze panes belongs to the **window**, not the sheet, so the sheet has to be active first. Turning freeze off before turning it on means running the macro twice does not cause problems.

### DoChecks: the five checks

```vba
With Application.WorksheetFunction
    nUnmapped = .CountIf(raw.Range("U:U"), "UNMAPPED")
    ...
End With
```

**Application.WorksheetFunction** lets VBA call the same COUNTIF, COUNTIFS and SUM I already used in the sheet. So the macro checks are exactly the checks I trusted by hand.

The five checks:

| Check | Question it answers | Should be |
|---|---|---|
| nUnmapped | Did any raw value fail to map? | 0 |
| nNoKg | Is any included row (no exclude_reason) missing final_kg? | 0 |
| nHdc | Any HDC row with a wrong material or no units? | 0 |
| nUnitW | Any unit-weight flag not fixed by an override and not excluded? | 0 |
| kgRaw vs kgSub | Does the submission add up to the clean table? | within 500 kg |

```vba
nHdc = .CountIf(raw.Range("V:V"), "INVALID HDC MATERIAL") _
     + .CountIf(raw.Range("V:V"), "HDC NO UNITS")
```

The **underscore** at the end of a line means "this statement continues on the next line".

**Variable types:** `Long` is a whole number (counts). `Double` is a decimal number (kg). `Boolean` is True or False. `String` is text.

### DoChecks: pass or fail

```vba
passed = (nUnmapped + nNoKg + nHdc + nUnitW = 0) And (Abs(kgRaw - kgSub) < 500)
```

All four counts must be 0 **and** the totals must be close. The decision is made once and stored, then used twice (in the log and in the message), so the two can never disagree.

### DoChecks: the audit log

```vba
If logWs.Range("A1").Value = "" Then
    logWs.Range("A1:H1").Value = Array("run_time", "unmapped", ...)
End If
```

Writes the headers only the first time. **Array(...)** writes all 8 headers in one go.

```vba
r = logWs.Cells(logWs.Rows.Count, "A").End(xlUp).Row + 1
```

**The most common line in VBA**, and a classic interview question. It means: go to the very last row of column A (row 1,048,576), press Ctrl+Up to jump to the last filled cell, take its row number, and add 1. That is the first empty row, however many runs are already logged.

```vba
logWs.Cells(r, 1).Value = Now
...
logWs.Cells(r, 8).Value = IIf(passed, "PASS", "FAIL")
```

`Cells(row, column)` uses numbers for both, which is easy to use with a variable row. **Now** is the current date and time. **IIf(condition, a, b)** is VBA's version of Excel's IF.

**Why log at all:** in compliance work you need to show when checks were run and what they found. A pop-up disappears; a log is evidence.

### DoChecks: the message and the return value

```vba
msg = "UNMAPPED codes: " & nUnmapped & vbNewLine & ...
MsgBox msg, vbInformation, "EPR checks"
DoChecks = passed
```

`&` joins text, `vbNewLine` is a line break, and `Format(kgRaw, "#,##0")` shows 9131216 as 9,131,216.

`DoChecks = passed` is how a **Function** returns its answer: you assign the value to the function's own name.

### Sub vs Function

| | Sub | Function |
|---|---|---|
| Does something | Yes | Yes |
| Gives back a value | No | Yes (like SUM gives back a number) |
| Shows in Alt+F8 and can go on a button | Yes | No |

That is why there are two: **DoChecks** (a Function, so CheckAndExport can use its answer) and **RunChecks** (a one-line Sub, so the button still works).

### CheckAndExport: the gate

```vba
If DoChecks() Then
    ExportSubmission
Else
    MsgBox "Export stopped. Fix the failed checks first.", vbExclamation, "EPR export"
End If
```

**What it does:** runs the checks; if they return True, it exports; if not, it blocks the export.

**Why:** without the gate, nothing stops a bad file being submitted. This is how a manual process becomes a safe process.

### ExportSubmission

```vba
Const OUT_FOLDER As String = "C:\Users\saika\OneDrive\Training\valpak_epr_practice\VBA\"
```

**Const** is a fixed value set once at the top. If the folder ever moves, one line changes.

```vba
lastRow = subm.Cells(subm.Rows.Count, "A").End(xlUp).Row
```

The same last-row trick, so it works however many rows the submission has.

```vba
Set newWb = Workbooks.Add
newWb.Worksheets(1).Range("A1:J" & lastRow).Value = _
    subm.Range("A1:J" & lastRow).Value
```

Creates a new workbook and copies **values only**, in one assignment. No clipboard, no copy-paste. The upload file must not contain formulas pointing at sheets that do not exist in it. Only columns A to J are copied, so my helper key column (L) stays out.

```vba
fileName = OUT_FOLDER & "submission_2025_" & Format(Now, "yyyymmdd_hhmm") & ".csv"
```

A name like `submission_2025_20261007_2020.csv`. A new file each run, so an older export is never overwritten. That gives a history of what was sent.

```vba
Application.DisplayAlerts = False
newWb.SaveAs Filename:=fileName, FileFormat:=xlCSVUTF8
newWb.Close SaveChanges:=False
Application.DisplayAlerts = True
```

Turns off Excel's "are you sure?" prompts during the save, saves as UTF-8 CSV, closes the temporary workbook, and **always turns alerts back on**.

## How I tested it

| Test | What I did | Result |
|---|---|---|
| Pass | Ran RunChecks on the real data | All counts 0, totals 9,131,216 vs 9,131,217, ALL CHECKS PASSED |
| Fail | Changed Raw_ALL D2 from AL to XYZ and ran again | UNMAPPED 1 and HDC problems 1, SOMETHING NEEDS A LOOK |
| Gate, pass | Ran CheckAndExport on clean data | Exported 446 rows to a dated CSV |
| Gate, fail | Broke D2 again and ran CheckAndExport | "Export stopped", no file created, FAIL row in Check_Log |
| Output | Opened the CSV | 447 lines (1 header + 446 rows), units only on AVD HDC rows |

Why the fail test gave **two** alarms: XYZ became UNMAPPED, and because D2 is an HDC aluminium can, "UNMAPPED" also broke the HDC material rule. **One bad cell, two alarms.** When several flags fire, look for the single root cause first. The kg totals did not move, which is exactly why codes are checked separately from totals.

## VBA facts I learned the hard way

- **Running a macro wipes the Ctrl+Z history.** Note any value before testing on it.
- **F5 means "run" only inside the VBA editor.** In the Excel window, F5 opens Go To. From Excel, use **Alt+F8**.
- **A message box shows in the Excel window**, so it can hide behind the VBA editor and look like nothing happened.
- **When assigning a macro to a button**, click the macro name in the list. The name box is pre-filled with a suggested new name (Button1_Click), which does not exist.
- **Automate > Record Actions is not VBA.** That records Office Scripts (TypeScript). VBA is **Developer > Record Macro**.
- **Export the code** (right-click the module, Export File) as a .bas text file, so it can be shared or put on GitHub without the workbook.

## The result

| | Time |
|---|---|
| By hand: 5 checks, decide, copy values to a new workbook, save as CSV | about 6 minutes |
| Macro: CheckAndExport | about 3 seconds |

And the macro does more than the manual run: it logs every run and refuses to export a failing file.

> "By hand it took me about 6 minutes per run. The macro takes about 3 seconds, logs every run for an audit trail, and won't export if any check fails. I tested it both ways by breaking a value on purpose."
