' Add Tabs - summary, data
Private Sub add_tabs()
    Sheets("ms summary").Select
    Sheets("ms summary").Name = "data"
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "summary"
End Sub

' Create Pivot Table - Range(A:V)
Private Sub create_pivot_table()
    Range("A:V").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
        SourceData:="data!A:V", _
        Version:=xlPivotTableVersion14). _
        CreatePivotTable TableDestination:="summary!R1C1", _
        TableName:="PivotTable1", _
        DefaultVersion:=xlPivotTableVersion14
    Sheets("summary").Select
End Sub

' Populate Pivot Table
Private Sub populate_pivot_table()
    ' y - Category A, B, C
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Category A")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Category B")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Category C")
        .Orientation = xlRowField
        .Position = 3
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Category A").ShowDetail = False
    
    ' x - Imputed Revenue, GP, GP %
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Virtually Adjusted Imputed Revenue"), _
        "Imputed Revenue ", xlSum
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Virtually Adjusted GP"), _
        "GP ", xlSum
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Calendar Year")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").CalculatedFields.Add "GP %", _
        "=IFERROR('Virtually Adjusted GP' /'Virtually Adjusted Imputed Revenue',0)", _
        True
    
    ' Number Format Imputed, GP, GP %
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_(* #,##0.0_);_(* (#,##0.0);_(* ""-""??_);_(@_)"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("GP %"), "GP % ", xlSum
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("GP % ")
        .NumberFormat = "0.0%"
        .Orientation = xlDataField
    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Calendar Year")
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable1").RowGrand = False
    
    ' Keep Pivot Table Position Fixed, Remove Field List
    ActiveSheet.PivotTables("PivotTable1").HasAutoFormat = False
    ActiveWorkbook.ShowPivotTableFieldList = False
End Sub

' Add Slicers - District, Fiscal Period, OB or TSR, Region, Solution Group
Private Sub add_slicers()
    ActiveWorkbook.SlicerCaches.Add(ActiveSheet.PivotTables("PivotTable1"), _
        "Fiscal Period").Slicers.Add ActiveSheet, , "Fiscal Period", "Fiscal Period", _
        0, 500, 144, 95
    ActiveWorkbook.SlicerCaches("Slicer_Fiscal_Period").Slicers("Fiscal Period"). _
        NumberOfColumns = 3
    ActiveWorkbook.SlicerCaches.Add(ActiveSheet.PivotTables("PivotTable1"), _
        "Region").Slicers.Add ActiveSheet, , "Region", "Region", _
        96, 500, 144, 180
    ActiveWorkbook.SlicerCaches.Add(ActiveSheet.PivotTables("PivotTable1"), _
        "OB or TSR").Slicers.Add ActiveSheet, , "OB or TSR", "OB or TSR", _
        0, 645, 144, 140
    ActiveWorkbook.SlicerCaches.Add(ActiveSheet.PivotTables("PivotTable1"), _
        "District").Slicers.Add ActiveSheet, , "District", "District", _
        141, 645, 144, 236
    ActiveWorkbook.SlicerCaches.Add(ActiveSheet.PivotTables("PivotTable1"), _
        "Solution Group").Slicers.Add ActiveSheet, , "Solution Group", "Solution Group", _
        277, 500, 144, 100
End Sub

Private Sub save_file()
    Dim todays_date
    todays_date = Format(Date$, "yyyy-mm-dd")
    ActiveWorkbook.SaveAs Filename:= _
        "P:\_HHOS\MS Summary Report\MS Summary - " & todays_date & ".xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
    Range("A1").Select
End Sub

'############### ms_summary_main() ###############

Sub ms_summary_main()
    Application.ScreenUpdating = False
    add_tabs
    create_pivot_table
    populate_pivot_table
    add_slicers
    save_file
  
    MsgBox "ms_summary_main() completed."
    Application.ScreenUpdating = True
End Sub

