Sub Bigfive()
Call CreatePivot
Call CopyData
Call FormatData
Call MoveSheets
ActiveWorkbook.Save
MsgBox ("Data Generated")
End Sub

Sub CopyData()
    
    Application.ScreenUpdating = False
    big1 = Sheets("Pivot").Range("A4").Value
    big2 = Sheets("Pivot").Range("A5").Value
    Big3 = Sheets("Pivot").Range("A6").Value
    Big4 = Sheets("Pivot").Range("A7").Value
    Big5 = Sheets("Pivot").Range("A8").Value
    
'creating a sheet
    Sheets("Dump").Select
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Final"
    
    Sheets("Dump").Select
    
    ActiveSheet.Range("$A$1:$E$611546").AutoFilter Field:=2, Criteria1:= _
        big1
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("A1:D1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Final").Select
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Sheets("Dump").Select
    Range("A1").Select
    If ActiveSheet.AutoFilterMode Then ActiveSheet.ShowAllData
    
    'ActiveCell.FormulaR1C1 = "0012X00001lrKksQAE"
    'Sheets("Dump").Select
    'ActiveWindow.SmallScroll Down:=-216
    'ActiveWindow.ScrollRow = 2606
    'ActiveWindow.ScrollRow = 1
    
    
    ActiveSheet.Range("$A$1:$E$611546").AutoFilter Field:=2, Criteria1:= _
        big2
        Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("A1:D1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Final").Select
    Range("F1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Sheets("Dump").Select
    Range("A1").Select
    If ActiveSheet.AutoFilterMode Then ActiveSheet.ShowAllData
    
    
    ActiveSheet.Range("$A$1:$E$611546").AutoFilter Field:=2, Criteria1:= _
        Big3
        Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("A1:D1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Final").Select
    Range("K1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Sheets("Dump").Select
    Range("A1").Select
    If ActiveSheet.AutoFilterMode Then ActiveSheet.ShowAllData
    
    ActiveSheet.Range("$A$1:$E$611546").AutoFilter Field:=2, Criteria1:= _
        Big4
        Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("A1:D1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Final").Select
    Range("P1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Sheets("Dump").Select
    Range("A1").Select
    If ActiveSheet.AutoFilterMode Then ActiveSheet.ShowAllData
    
    ActiveSheet.Range("$A$1:$E$611546").AutoFilter Field:=2, Criteria1:= _
        Big5
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("A1:D1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Final").Select
    Range("U1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Sheets("Dump").Select
    Range("A1").Select
    If ActiveSheet.AutoFilterMode Then ActiveSheet.ShowAllData
    
    Application.ScreenUpdating = True
    
End Sub
    
Sub CreatePivot()

    Application.ScreenUpdating = False
    Sheets("Dump").Select
    Columns("A:D").Select
    Sheets.Add
    ActiveSheet.Name = "Pivot"
    'end point of data(x,y)
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Dump!R1C1:R1048576C4", Version:=6).CreatePivotTable TableDestination:= _
        "Pivot!R3C1", TableName:="PivotTable1", DefaultVersion:=6
    Sheets("Pivot").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable1")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("account_sfid")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("product_sfid"), "Count of product_sfid", xlCount
    Range("B6").Select
    ActiveSheet.PivotTables("PivotTable1").PivotFields("account_sfid").AutoSort _
        xlDescending, "Count of product_sfid", ActiveSheet.PivotTables("PivotTable1"). _
        PivotColumnAxis.PivotLines(1), 1
    ActiveSheet.PivotTables("PivotTable1").PivotFields("account_sfid").PivotFilters _
        .Add2 Type:=xlTopCount, DataField:=ActiveSheet.PivotTables("PivotTable1"). _
        PivotFields("Count of product_sfid"), Value1:=5
    
    Range("A3").Select
    
    Application.ScreenUpdating = True
    
End Sub


Sub MoveSheets()

Dim Path As String
Path = ActiveWorkbook.Path & "\Top 5 Accounts with Productid.xlsx"
Sheets(Array("Pivot", "Final")).Move
    ActiveWorkbook.SaveAs Filename:=Path, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Sheets("Final").Select
    Range("A1").Select
    Sheets("Pivot").Select
    Range("A1").Select
    ActiveWindow.Close
End Sub


Sub FormatData()
Application.ScreenUpdating = False
Sheets("Final").Select

    Range("D1048576").Select
    Selection.End(xlUp).Select
    x = ActiveCell.Row
    Range("D1").Select
    y = ActiveCell.Column
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


    Range("I1048576").Select
    Selection.End(xlUp).Select
    x1 = ActiveCell.Row
    Range("I1").Select
    y1 = ActiveCell.Column
    Range(Cells(x1, y1), Cells(1, 6)).Select
 
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

    
    Range("N1048576").Select
    Selection.End(xlUp).Select
    x2 = ActiveCell.Row
    Range("N1").Select
    y2 = ActiveCell.Column
    Range(Cells(x2, y2), Cells(1, 11)).Select
 
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

    
    Range("S1048576").Select
    Selection.End(xlUp).Select
    x3 = ActiveCell.Row
    Range("S1").Select
    y3 = ActiveCell.Column
    Range(Cells(x3, y3), Cells(1, 16)).Select

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

    
    Range("X1048576").Select
    Selection.End(xlUp).Select
    x4 = ActiveCell.Row
    Range("X1").Select
    y4 = ActiveCell.Column
    Range(Cells(x4, y4), Cells(1, 21)).Select

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

    Range("A1:D1").Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Range("F1:I1").Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    
    Range("K1:N1").Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
        Range("P1:S1").Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
        Range("U1:X1").Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Range(Cells(x, y4), Cells(1, 1)).Select
    Selection.Columns.AutoFit
    Range("A1").Select
    
    Sheets("Main").Select
    Range("A1").Select
    
    Sheets("Dump").Select
    Range("A1").Select
    Application.ScreenUpdating = True
    
End Sub
