' Add Tabs - cust, enrol
Private Sub add_tabs()
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "cust"
    Sheets("Sheet2").Select
    Sheets("Sheet2").Name = "enrol"
End Sub

' Create Pivot Table - Range(A:S)
Private Sub create_pivot_table()
    Sheets("cust_data").Select
    Range("A:S").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
        SourceData:="cust_data!A:S", _
        Version:=xlPivotTableVersion14). _
        CreatePivotTable TableDestination:="cust!R1C1", _
        TableName:="cust_table", _
        DefaultVersion:=xlPivotTableVersion14
    
    Sheets("enrol_data").Select
    Range("A:R").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
        SourceData:="enrol_data!A:R", _
        Version:=xlPivotTableVersion14). _
        CreatePivotTable TableDestination:="enrol!R1C1", _
        TableName:="enrol_table", _
        DefaultVersion:=xlPivotTableVersion14
        
    Sheets("enrol").Select
End Sub

' Populate Pivot Table
Private Sub populate_pivot_table()

    ' _______________ Enrol Pivot Table _______________
    ' y - OB / TSR, Region, District
    With ActiveSheet.PivotTables("enrol_table").PivotFields("OB / TSR")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("enrol_table").PivotFields("Region")
        .Orientation = xlRowField
        .Position = 2
        .PivotItems("(blank)").Visible = False
    End With
    With ActiveSheet.PivotTables("enrol_table").PivotFields("District")
        .Orientation = xlRowField
        .Position = 3
    End With
    ActiveSheet.PivotTables("enrol_table").PivotFields("Region").ShowDetail = False
    
    ' x - Licensing ProgramName
    With ActiveSheet.PivotTables("enrol_table").PivotFields("Licensing ProgramName")
        .Orientation = xlColumnField
        .Position = 1
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("enrol_table").AddDataField ActiveSheet.PivotTables( _
        "enrol_table").PivotFields("Contract Number"), _
        "Contract Number ", xlCount
    
    ' Keep Pivot Table Position Fixed, Light Black Pivot Layout, Remove Field List, Center
    ActiveSheet.PivotTables("enrol_table").HasAutoFormat = False
    ActiveSheet.PivotTables("enrol_table").TableStyle2 = "PivotStyleLight2"
    ActiveWorkbook.ShowPivotTableFieldList = False
    Columns("B:I").HorizontalAlignment = xlCenter
    
    ' _______________ Cust Pivot Table _______________
    ' y - OB / TSR, Region, District
    Sheets("cust").Select
    With ActiveSheet.PivotTables("cust_table").PivotFields("OB / TSR")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("cust_table").PivotFields("Region")
        .Orientation = xlRowField
        .Position = 2
        .PivotItems("(blank)").Visible = False
    End With
    With ActiveSheet.PivotTables("cust_table").PivotFields("District")
        .Orientation = xlRowField
        .Position = 3
    End With
    ActiveSheet.PivotTables("cust_table").PivotFields("Region").ShowDetail = False
    
    ' x - Licensing ProgramName
    With ActiveSheet.PivotTables("cust_table").PivotFields("Licensing ProgramName")
        .Orientation = xlColumnField
        .Position = 1
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("cust_table").AddDataField ActiveSheet.PivotTables( _
        "cust_table").PivotFields("Contract Number"), _
        "Contract Number ", xlCount
    
    ' Keep Pivot Table Position Fixed, Light Black Pivot Layout, Remove Field List, Center
    ActiveSheet.PivotTables("cust_table").HasAutoFormat = False
    ActiveSheet.PivotTables("cust_table").TableStyle2 = "PivotStyleLight3"
    ActiveWorkbook.ShowPivotTableFieldList = False
    Columns("B:I").HorizontalAlignment = xlCenter
   
End Sub

Private Sub save_file()
    Dim todays_date
    todays_date = Format(Date$, "yyyy-mm-dd")
    ActiveWorkbook.SaveAs Filename:= _
        "P:\_HHOS\Microsoft\Enrol Cust - " & todays_date & ".xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
    Range("A1").Select
End Sub

'############### enrol_cust_main() ###############

Sub enrol_cust_main()
    Application.ScreenUpdating = False
    add_tabs
    create_pivot_table
    populate_pivot_table
    save_file
 
    MsgBox "enrol_cust_main() completed."
    Application.ScreenUpdating = True
End Sub
