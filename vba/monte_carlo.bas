Dim t0 As Single
Dim t1 As Single

'############### 0 - utils ###############
Function ConvertToLetter(iCol As Integer) As String
   Dim iAlpha As Integer
   Dim iRemainder As Integer
   iAlpha = Int(iCol / 27)
   iRemainder = iCol - (iAlpha * 26)
   If iAlpha > 0 Then
      ConvertToLetter = Chr(iAlpha + 64)
   End If
   If iRemainder > 0 Then
      ConvertToLetter = ConvertToLetter & Chr(iRemainder + 64)
   End If
End Function

' Add Tabs - summary, data
Private Sub add_tabs()
    Sheets("oracle_mc").Select
    Sheets("oracle_mc").Name = "data"
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "summary"
End Sub

' Create Pivot Table - Range(A:G)
Private Sub create_pivot_table()
    Range("A:G").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
        SourceData:="data!A:G", _
        Version:=xlPivotTableVersion14). _
        CreatePivotTable TableDestination:="summary!R1C1", _
        TableName:="PivotTable1", _
        DefaultVersion:=xlPivotTableVersion14
    Sheets("summary").Select
End Sub

' Populate Pivot Table
Private Sub populate_pivot_table()
    ' y - Region, Super Category
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Region")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Super Category")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Region")
        .PivotItems("(blank)").Visible = False
    End With
        
    ' x - Rev, FM, FM
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Rev"), _
        "Rev ", xlSum
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("FM"), _
        "FM ", xlSum
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Year")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").CalculatedFields.Add "FM %", _
        "=IFERROR('FM' /'Rev',0)", _
        True
    
    ' Number Format Rev, FM, FM %
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_(* #,##0.0_);_(* (#,##0.0);_(* ""-""??_);_(@_)"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("FM %"), "FM % ", xlSum
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("FM % ")
        .NumberFormat = "0.0%"
        .Orientation = xlDataField
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Year")
        .PivotItems("(blank)").Visible = False
    End With
    
    ' Keep Pivot Table Position Fixed, Remove Field List, Bottom Subtotal
    ActiveSheet.PivotTables("PivotTable1").HasAutoFormat = False
    ActiveWorkbook.ShowPivotTableFieldList = False
    ActiveSheet.PivotTables("PivotTable1").RowGrand = False
    ActiveSheet.PivotTables("PivotTable1").ColumnGrand = False
    ActiveSheet.PivotTables("PivotTable1").SubtotalLocation xlAtBottom
    
    ' Rename y / x axis
    ActiveSheet.PivotTables("PivotTable1").CompactLayoutColumnHeader = "Year "
    ActiveSheet.PivotTables("PivotTable1").CompactLayoutRowHeader = _
        "Region / Super Category"
End Sub

' Convert Pivot Table to txt
Private Sub pivot_to_txt()
    With ActiveSheet.PivotTables("PivotTable1")
        .PivotSelect "", x1DataAndLabel, True
        .TableStyle2 = "PivotStyleLight1"
    End With
    Selection.Copy
    Sheets.Add After:=Sheets(Sheets.Count)
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
End Sub

' Delete data / summary Tabs
Private Sub del_old_data()
    Sheets("summary").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("data").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets(1).Name = "historic"
End Sub

' Insert Empty Rows for EBITDA, Rebate, etc.
Private Sub insert_empty_rows()
    Dim row As Integer
    For row = 2 To 200
        cell = Cells(row, 1)
        cell_below = Cells(row + 1, 1)
        If (cell Like "*Total") And Not IsEmpty(cell_below) Then
            Rows(row + 1 & ":" & row + 15).Select
            Selection.Insert shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        End If
    Next row
End Sub

' Copy Latest Financial Data
Private Sub copy_latest_data()
    Dim last_col As Integer
    Dim start_col As String
    Dim end_col As String
    
    last_col = ActiveSheet.UsedRange.Columns.Count
    start_col = ConvertToLetter(last_col - 2)
    end_col = ConvertToLetter(last_col)
    
    Columns(start_col & ":" & end_col).Select
    Selection.Copy
End Sub

' Open mc_template.xlsx -> insert B1
Private Sub open_template()
    Workbooks.Open Filename:="P:\_HHOS\MC Model\mc_template.xlsx"
    Range("B1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub


'############### monte_carlo_main() ###############

Sub monte_carlo_main()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    t0 = Timer
    
    add_tabs
    create_pivot_table
    populate_pivot_table
    pivot_to_txt
    del_old_data
    insert_empty_rows
    copy_latest_data
    open_template
    
    t1 = Timer
    MsgBox "monte_carlo_main() completed. " + Format(t1 - t0, "Fixed") + "s"
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
