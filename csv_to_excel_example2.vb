Sub InboundReport()
Application.ScreenUpdating = False
Dim directory As String
Dim fileName As String
Dim wbcsv As Workbook
Dim wb As Workbook
Dim sheet As Worksheet
'   Report Name CR-15 Data 2017-MM-DD - Promotion  Tactic Detail - Prod_USA
'   Tab1    Source PrmProduct
'   Tab2    Source BplData

directory = Application.ActiveWorkbook.Path & "\Input\"
outDirectory = Application.ActiveWorkbook.Path & "\Output\"
fileName = Dir(directory & "inboundSuccessful.csv")
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
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Date"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "=TEXT(RC[-2],""MM/DD/YYYY"")"
    Range("A1048576").Select
    Selection.End(xlUp).Select
    x = ActiveCell.Row
    Range("XFD1").Select
    Selection.End(xlToLeft).Select
    y = ActiveCell.Column
    Cells(x, y).Select
    
    Range(Selection, Selection.End(xlUp)).Select
    Selection.FillDown
    
    Columns("A:G").Select
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Sheet1!R1C1:R1048576C7", Version:=6).CreatePivotTable TableDestination:= _
        "Sheet2!R3C1", TableName:="PivotTable2", DefaultVersion:=6
    Sheets("Sheet2").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Date")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("XmlObjectType")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("ObjectCount"), "Count of ObjectCount", xlCount
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Count of ObjectCount")
        .Caption = "Sum of ObjectCount"
        .Function = xlSum
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("BusinessTemplate")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("WorkState")
        .Orientation = xlRowField
        .Position = 4
    End With
    Range("B9").Select
    Range("A14").Select
    ActiveSheet.PivotTables("PivotTable2").PivotFields("BusinessTemplate"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    Range("B21").Select
    With ActiveSheet.PivotTables("PivotTable2")
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
    End With
    Range("A18").Select
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Date").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    Range("B10").Select
    ActiveSheet.PivotTables("PivotTable2").PivotFields("XmlObjectType").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    '----------------------------------------------------------------------------
    
    
    'ActiveSheet.PivotTables("PivotTable2").PivotSelect "'03/14/2017'", _
        'xlDataAndLabel, True
    ActiveSheet.PivotTables("PivotTable2").RepeatAllLabels xlRepeatLabels
    Columns("A:E").Select
    Selection.Copy
    
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Successful Inbound"
     
    Range("A1").Select
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
'    Selection.End(xlToLeft).Select
'    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Rows("1:3").Select
    
    Selection.Delete Shift:=xlUp
    Columns("A:A").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    Range("A1").Select
    ActiveSheet.Range("$A$1:$E$37").AutoFilter Field:=1, Criteria1:="(blank)"
    
    ActiveSheet.Range("$A$1:$I$" & 1000).Offset(1, 0).SpecialCells _
    (xlCellTypeVisible).EntireRow.Delete  '1000 is the lines you want to delete
    
    Selection.AutoFilter
    Range("A1").Select
    ActiveSheet.Range("$A$1:$E$11").AutoFilter Field:=1, Criteria1:= _
        "Grand Total"
    
    ActiveSheet.Range("$A$1:$I$" & 1000).Offset(1, 0).SpecialCells _
    (xlCellTypeVisible).EntireRow.Delete  '1000 is the lines you want to delete
    Range("A1").Select
    Selection.AutoFilter
    
    Range("A1").Select
    
    Columns("C:C").Select
    Selection.Replace What:="(blank)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
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
    Selection.Columns.AutoFit
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
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
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.499984740745262
        .PatternTintAndShade = 0
    End With
'    Range("A11:E11").Select
    'With Selection.Interior
        '.Pattern = xlSolid
        '.PatternColorIndex = xlAutomatic
        '.ThemeColor = xlThemeColorLight1
        '.TintAndShade = 0.499984740745262
        '.PatternTintAndShade = 0
    'End With
    'With Selection.Font
        '.ThemeColor = xlThemeColorDark1
        '.TintAndShade = 0
    'End With
    Range("A1:E1").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Range("G9").Select
    Selection.End(xlUp).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    
    Range("A1").Select
    
        Range("A1").Select
    Columns("B:B").Select
    Selection.Cut
    Selection.End(xlToLeft).Select
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Columns("D:D").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Columns("E:E").Select
    Selection.Cut
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Integration Object Type"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Object Count"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Comments"
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Columns.AutoFit
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
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
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
    Range("A1").Select
    Application.CutCopyMode = False
    
    ActiveWorkbook.SaveAs outDirectory & "Inbound_ReportMM-DD-YY.xlsx"
ActiveWorkbook.Close
Application.ScreenUpdating = True
MsgBox ("Inbound Report Saved")
    
End Sub
