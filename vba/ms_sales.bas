
' Add Tabs - summary, data
Private Sub add_tabs()
    Sheets("ms_sales").Select
    Sheets("ms_sales").Name = "data"
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "summary"
End Sub

' Create Pivot Table - Range(A:X)
Private Sub create_pivot_table()
    Sheets("data").Select
    Columns("A:Y").Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A:$Y"), , xlYes).Name = _
        "Table1"
    ActiveSheet.ListObjects("Table1").TableStyle = "TableStyleLight1"
    
    Columns("A:Y").Select
    Range("Table1[[#Headers],[Virtually Adjusted Revenue]]").Activate
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Table1", Version:=xlPivotTableVersion14).CreatePivotTable TableDestination _
        :="summary!R1C1", TableName:="PivotTable1", DefaultVersion:= _
        xlPivotTableVersion14
    Sheets("summary").Select
End Sub

' Populate Pivot Table
Private Sub populate_pivot_table()
    ' y - Category Sale or Referral, A, B, C
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Category A")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Sale or Referral")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Category B")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Category C")
        .Orientation = xlRowField
        .Position = 4
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Category B").ShowDetail = False
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Sale or Referral").ShowDetail = False
    
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
    
    ' Keep Pivot Table Position Fixed, Light Black Pivot Layout, Remove Field List,
    ' Expand Category A
    ActiveSheet.PivotTables("PivotTable1").HasAutoFormat = False
    ActiveSheet.PivotTables("PivotTable1").TableStyle2 = "PivotStyleLight1"
    ActiveWorkbook.ShowPivotTableFieldList = False
End Sub

' Add Slicers - District, Calendar Month, OB or TSR, Region, Solution Group
Private Sub add_slicers()

    ' Calendar Month
    ActiveWorkbook.SlicerCaches.Add(ActiveSheet.PivotTables("PivotTable1"), _
        "Calendar Month").Slicers.Add ActiveSheet, , "Calendar Month", "Calendar Month", _
        0, 700, 144, 120
    ActiveWorkbook.SlicerCaches("Slicer_Calendar_Month").Slicers("Calendar Month"). _
        NumberOfColumns = 3
        
    ' Region
    ActiveWorkbook.SlicerCaches.Add(ActiveSheet.PivotTables("PivotTable1"), _
        "Region").Slicers.Add ActiveSheet, , "Region", "Region", _
        121, 700, 144, 180
    ActiveWorkbook.SlicerCaches("Slicer_Region").Slicers("Region").Style = _
        "SlicerStyleLight2"
        
    ' Solution Group
    ActiveWorkbook.SlicerCaches.Add(ActiveSheet.PivotTables("PivotTable1"), _
        "Solution Group").Slicers.Add ActiveSheet, , "Solution Group", "Solution Group", _
        302, 700, 144, 100
    ActiveWorkbook.SlicerCaches("Slicer_Solution_Group").Slicers("Solution Group"). _
        Style = "SlicerStyleLight6"
    
    ' OB / TSR
    ActiveWorkbook.SlicerCaches.Add(ActiveSheet.PivotTables("PivotTable1"), _
        "OB or TSR").Slicers.Add ActiveSheet, , "OB or TSR", "OB or TSR", _
        0, 845, 144, 120
    ActiveWorkbook.SlicerCaches("Slicer_OB_or_TSR").Slicers("OB or TSR").Style = _
        "SlicerStyleLight4"

    ' District
    ActiveWorkbook.SlicerCaches.Add(ActiveSheet.PivotTables("PivotTable1"), _
        "District").Slicers.Add ActiveSheet, , "District", "District", _
        121, 845, 144, 401
    ActiveWorkbook.SlicerCaches("Slicer_District").Slicers("District").Style = _
        "SlicerStyleLight3"
        
    ' Level
    ActiveWorkbook.SlicerCaches.Add(ActiveSheet.PivotTables("PivotTable1"), _
        "Level").Slicers.Add ActiveSheet, , "Level", "Level", _
        403, 700, 144, 120
    ActiveWorkbook.SlicerCaches("Slicer_Level").Slicers("Level").Style = _
        "SlicerStyleLight3"
        
    ' Level
    ActiveWorkbook.SlicerCaches.Add(ActiveSheet.PivotTables("PivotTable1"), _
        "OB/TSR (Rep Team Name)").Slicers.Add ActiveSheet, , "OB/TSR (Rep Team Name)", "OB/TSR (Rep Team Name)", _
        0, 990, 144, 120
    
End Sub

Private Sub save_file()
    Dim todays_date
    todays_date = Format(Date$, "yyyy-mm-dd")
    ActiveWorkbook.SaveAs Filename:= _
        "P:\_HHOS\Microsoft\MS Sales - " & todays_date & ".xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
    Range("A1").Select
End Sub

'############### ms_sales_main() ###############

Sub ms_sales_main()
    Dim t0 As Single
    Dim t1 As Single
    Application.ScreenUpdating = False
    t0 = Timer
    
    add_tabs
    create_pivot_table
    populate_pivot_table
    add_slicers
    save_file
  
    t1 = Timer
    MsgBox "ms_sales_main() completed. " + Format(t1 - t0, "Fixed") + "s"
    Application.ScreenUpdating = True
End Sub



