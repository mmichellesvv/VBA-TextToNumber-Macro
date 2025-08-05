Attribute VB_Name = "Module1"
Sub TextNumberConverterWizard()
    Dim ws As Worksheet
    Dim cell As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long, j As Long
    Dim val As String
    Dim convertedCount As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' STEP 1: Highlight cells with text-formatted numbers (light red)
    For Each ws In ThisWorkbook.Worksheets
        Application.StatusBar = "Highlighting on sheet: " & ws.Name
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

        For i = 1 To lastCol
            For j = 1 To lastRow
                Set cell = ws.Cells(j, i)
                val = Trim(cell.Value)

                If val <> "" Then
                    If IsNumeric(val) And VarType(cell.Value) = vbString Then
                        cell.Interior.Color = RGB(255, 200, 200)
                    End If
                End If
            Next j
        Next i
    Next ws

    Application.StatusBar = False
    Application.ScreenUpdating = True

    If MsgBox("Text-formatted numbers were highlighted (light RED). Do you want to convert them to numbers?", vbYesNo + vbQuestion) = vbNo Then
        MsgBox "Operation canceled. No changes were made."
        Exit Sub
    End If

    ' STEP 2: Create a backup copy before making changes
    If MsgBox("Would you like to create a backup before converting?", vbYesNo + vbQuestion) = vbYes Then
        On Error Resume Next
        ThisWorkbook.SaveCopyAs ThisWorkbook.Path & "\" & "Backup_" & Format(Now, "yyyymmdd_HHMMSS") & "_" & ThisWorkbook.Name
        On Error GoTo 0
    End If

    ' STEP 3: Convert text-formatted numbers to actual numbers
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    convertedCount = 0

    For Each ws In ThisWorkbook.Worksheets
        Application.StatusBar = "Converting on sheet: " & ws.Name
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

        For i = 1 To lastCol
            For j = 1 To lastRow
                Set cell = ws.Cells(j, i)
                val = Trim(cell.Value)

                If val <> "" Then
                    If IsNumeric(val) And VarType(cell.Value) = vbString Then
                        On Error Resume Next
                        cell.Value = CDbl(val) ' Convert to number, preserve original format
                        On Error GoTo 0

                        cell.Interior.ColorIndex = xlNone ' Clear highlight
                        convertedCount = convertedCount + 1
                    End If
                End If
            Next j
        Next i
    Next ws

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Done! Converted " & convertedCount & " cells.", vbInformation
End Sub

