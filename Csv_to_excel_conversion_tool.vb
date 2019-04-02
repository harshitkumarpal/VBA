Sub ClosedCheck()
Application.ScreenUpdating = False
Dim directory As String
Dim fileName As String
Dim wbcsv As Workbook
Dim wb As Workbook
Dim sheet As Worksheet
Dim i As Integer
Dim j As Integer

directory = Application.ActiveWorkbook.Path & "\Input\"
outDirectory = Application.ActiveWorkbook.Path & "\Output\"
fileName = Dir(directory & "ClosedCheckDeduction_Deduction.csv")
Workbooks.Open (directory & fileName)
'ActiveWorkbook.Select
Set wbcsv = ActiveWorkbook
ActiveSheet.Select
Cells.Select
Selection.Copy

Workbooks.Add

Sheets("Sheet1").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Set wb = ActiveWorkbook
wbcsv.Close savechanges:=False
wb.Activate
Range("A1: AF1").Select
'ActiveCell.Select
Selection.Font.Bold = True
Range("A2").Select
ActiveWindow.FreezePanes = True
    
    Range("A1048576").Select
    Selection.End(xlUp).Select
    x = ActiveCell.Row
    Range("XFD1").Select
    Selection.End(xlToLeft).Select
    y = ActiveCell.Column
    Cells(x, y).Select
    Selection.End(xlToLeft).Select
    Range(Cells(x, y), Cells(1, 1)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range(Cells(1, 1), Cells(1, y)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Range(Cells(x, y), Cells(1, 1)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Columns.AutoFit
    Range("A2").Select
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "ClosedDeduction"

'ClosedCheckDeduction_CheckPayment--------------------------------------

fileName = Dir(directory & "ClosedCheckDeduction_CheckPayment.csv")
Workbooks.Open (directory & fileName)
'ActiveWorkbook.Select
Set wbcsv = ActiveWorkbook

ActiveSheet.Select
Cells.Select
Selection.Copy
wb.Activate
Sheets.Add After:=ActiveSheet
Sheets("Sheet2").Name = "ClosedCheck"
Sheets("ClosedCheck").Select
    ActiveSheet.Paste
    Selection.Columns.AutoFit
    Range("A2").Select
    Application.CutCopyMode = False
    wbcsv.Close savechanges:=False
wb.Activate

Range("A1: AF1").Select
'ActiveCell.Select
Selection.Font.Bold = True
Range("A2").Select
ActiveWindow.FreezePanes = True
    
    Range("A1048576").Select
    Selection.End(xlUp).Select
    x = ActiveCell.Row
    Range("XFD1").Select
    Selection.End(xlToLeft).Select
    y = ActiveCell.Column
    Cells(x, y).Select
    Selection.End(xlToLeft).Select
    Range(Cells(x, y), Cells(1, 1)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range(Cells(1, 1), Cells(1, y)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Range(Cells(x, y), Cells(1, 1)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Columns.AutoFit
    Range("A2").Select
    
ActiveWorkbook.SaveAs outDirectory & "Closed Check and Deduction Payments Canada sun to sat MMM '17.xlsx"
wb.Close
Application.ScreenUpdating = True
MsgBox ("Closed Check Payment Report Saved")

End Sub
