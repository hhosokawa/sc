' Add Tabs - summary, data
Private Sub add_tabs()
    Sheets("NNEA COCP").Select
    Sheets("NNEA COCP").Name = "data"
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "summary"
End Sub

' Create Pivot Table - Range(A:AS)
Private Sub create_pivot_table()
    Range("A:AS").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
        SourceData:="data!A:AS", _
        Version:=xlPivotTableVersion14). _
        CreatePivotTable TableDestination:="summary!R1C1", _
        TableName:="PivotTable1", _
        DefaultVersion:=xlPivotTableVersion14
    Sheets("summary").Select
End Sub

' Populate Pivot Table
Private Sub populate_pivot_table()
    ' y - Region, Branch, Rep
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Region")
        .Orientation = xlRowField
        .Position = 1
        .PivotItems("(blank)").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Branch")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Rep")
        .Orientation = xlRowField
        .Position = 3
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Branch").ShowDetail = False
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Region").ShowDetail = False
    
    ' x - NNEA, Renewal, COCP Win, COCP Loss
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("NNEA"), "NNEA ", xlCount
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Renewal"), "Renewal ", xlCount
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("COCP Win"), "COCP Win ", xlCount
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("COCP Loss"), "COCP Loss ", xlCount

    ' Center Fields, Keep Pivot Table Position Fixed, Remove Field List
    Columns("B:F").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
    ActiveSheet.PivotTables("PivotTable1").HasAutoFormat = False
    ActiveWorkbook.ShowPivotTableFieldList = False
End Sub

' Add Slicers - Effective Year, Month / Acct, Rep Type
Private Sub add_slicers()
    ActiveWorkbook.SlicerCaches.Add(ActiveSheet.PivotTables("PivotTable1"), _
        "Effective Year").Slicers.Add ActiveSheet, , "Effective Year", "Effective Year", _
        0, 365, 144, 95
    ActiveWorkbook.SlicerCaches.Add(ActiveSheet.PivotTables("PivotTable1"), _
        "Effective Month").Slicers.Add ActiveSheet, , "Effective Month", "Effective Month", _
        96, 365, 144, 115
    ActiveWorkbook.SlicerCaches("Slicer_Effective_Month").Slicers("Effective Month"). _
        NumberOfColumns = 3
    ActiveWorkbook.SlicerCaches.Add(ActiveSheet.PivotTables("PivotTable1"), _
        "Acct Type").Slicers.Add ActiveSheet, , "Acct Type", "Acct Type", _
        0, 510, 144, 95
    ActiveWorkbook.SlicerCaches.Add(ActiveSheet.PivotTables("PivotTable1"), _
        "Rep Type").Slicers.Add ActiveSheet, , "Rep Type", "Rep Type", _
        96, 510, 144, 95
        
    ' Select 2014, Enrollment, Rep Type
    With ActiveWorkbook.SlicerCaches("Slicer_Effective_Year")
        .SlicerItems("2014").Selected = True
        .SlicerItems("2013").Selected = False
        .SlicerItems("(blank)").Selected = False
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Acct_Type")
        .SlicerItems("Enrollment").Selected = True
        .SlicerItems("Acct").Selected = False
        .SlicerItems("(blank)").Selected = False
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Rep_Type")
        .SlicerItems("OB").Selected = True
        .SlicerItems("TSR").Selected = False
        .SlicerItems("(blank)").Selected = False
    End With
End Sub

' Add Formula "=IF(SUM(B3:D3)-E3=0,"",SUM(B3:D3)-E3)"
Private Sub add_formulas()
    Range("E1").Select
    Selection.Copy
    Range("F1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

    Range("F1").Value = "NNEA + Renewal + COCP Win - COCP Loss"
    Range("F2").formula = "=IF(SUM(B2:D2)-E2=0, """" ,SUM(B2:D2)-E2)"
    Range("F2").Copy
    Range("F2:F100").PasteSpecial (xlPasteAll)
End Sub

Private Sub save_file()
    Dim todays_date
    todays_date = Format(Date$, "yyyy-mm-dd")
    ActiveWorkbook.SaveAs Filename:= _
        "P:\_HHOS\MS Summary Report\NNEA COCP - " & todays_date & ".xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
    Range("A1").Select
End Sub

'############### ms_cocp_main() ###############

Sub ms_cocp_main()
    Application.ScreenUpdating = False
    add_tabs
    create_pivot_table
    populate_pivot_table
    add_formulas
    add_slicers
    save_file
  
    MsgBox "ms_cocp_main() completed."
    Application.ScreenUpdating = True
End Sub
