Attribute VB_Name = "Module6"
'=============================================================================
'|Consoidation
'|
'|ref to https://msdn.microsoft.com/en-us/library/cc793964(v=office.12).aspx
'=============================================================================

Function createConsolidationQTY()
    Dim sh As Worksheet
    Dim DestSh As Worksheet
    Dim Last As Long
    Dim shLast As Long
    Dim CopyRng As Range
    Dim StartRow As Long

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    ' Delete the summary sheet if it exists.
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Worksheets("MergedQTY").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Add a new summary worksheet.
    Set DestSh = ActiveWorkbook.Worksheets.Add
    DestSh.name = "MergedQTY"

    ' Fill in the start row.
    StartRow = 2

    ' Loop through all worksheets and copy the data to the
    ' summary worksheet.
    For Each sh In ActiveWorkbook.Worksheets
       If Right(sh.name, 3) = "QTY" Then
        Debug.Print sh.name

            ' Find the last row with data on the summary
            ' and source worksheets.
            Last = LastRow(DestSh)
            shLast = LastRow(sh)

            ' If source worksheet is not empty and if the last
            ' row >= StartRow, copy the range.
            If shLast > 0 And shLast >= StartRow Then
                'Set the range that you want to copy
                Set CopyRng = sh.Range(sh.Rows(StartRow), sh.Rows(shLast))

               ' Test to see whether there are enough rows in the summary
               ' worksheet to copy all the data.
                If Last + CopyRng.Rows.Count > DestSh.Rows.Count Then
                   MsgBox "There are not enough rows in the " & _
                   "summary worksheet to place the data."
                   GoTo ExitTheSub
                End If

                ' This statement copies values and formats.
                CopyRng.Copy
                With DestSh.Cells(Last + 1, "A")
                    .PasteSpecial xlPasteValues
                    .PasteSpecial xlPasteFormats
                    Application.CutCopyMode = False
                End With

            End If

        End If
    Next
    DestSh.Rows(1).Insert
    
    DestSh.Cells(1, 1) = "PART#"
    DestSh.Cells(1, 2) = "DESCRIPTION"
    DestSh.Cells(1, 3) = "MRP TYPE"
    DestSh.Cells(1, 4) = "PLANNED"
    DestSh.Cells(1, 5) = "ORDERED"
    DestSh.Cells(1, 6) = "TO ORDER"
    DestSh.Cells(1, 7) = "DELIVERED"
    DestSh.Cells(1, 8) = "OPEN QTY"
    DestSh.Cells(1, 9) = "BOM STATUS"
               
    lrow = DestSh.Cells(Rows.Count, 1).End(xlUp).Row
    
    DestSh.Range("A1:I" & lrow).AutoFilter _
        field:=9, _
        Criteria1:="<>deleted in BOM", _
        VisibleDropDown:=True

ExitTheSub:

    Application.Goto DestSh.Cells(1)

    ' AutoFit the column width in the summary sheet.
    DestSh.Columns.AutoFit

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Function


Function CreateConsolidationPivotTable()

    Dim sht As Worksheet
    Dim pvtCache As PivotCache
    Dim pvt As PivotTable
    Dim StartPvt As String
    Dim SrcData As String
    Dim SheetName As String
    
    projName = "Merged"
     
    deleteSheet (projName & "PIVOT")
     
    
    lcol = Sheets(projName & "QTY").Cells(1, Columns.Count).End(xlToLeft).Column
    lrow = Sheets(projName & "QTY").Cells(Rows.Count, 1).End(xlUp).Row
    SrcData = Sheets(projName & "QTY").name & "!" & Range("A1:H" & lrow).Address(ReferenceStyle:=xlR1C1)
      
       
    Set sht = Sheets.Add
    ActiveSheet.name = projName & "PIVOT"
    
    
    '----------------------------------------------------------------
    StartPvt = sht.name & "!" & sht.Range("A1").Address(ReferenceStyle:=xlR1C1)
    Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)
    Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt)
        
        pvt.PivotFields("MRP TYPE").Orientation = xlRowField
        pvt.AddDataField pvt.PivotFields("PLANNED"), "Sum of Planned", xlSum
        pvt.AddDataField pvt.PivotFields("ORDERED"), "Sum of Ordered", xlSum
        pvt.ManualUpdate = False
    Set pf = Sheets(projName & "PIVOT").PivotTables(1).PivotFields("MRP TYPE")
    
    ActiveSheet.Range(Cells(2, 1), Cells(lrow, lcol)).Select
        Set myChart = ActiveSheet.Shapes.AddChart(xlColumnClustered, 300, 10, 600, 400).Chart
        With myChart
            .PlotBy = xlColumns
            .ChartArea.Format.TextFrame2.TextRange.Font.Size = 8
            .HasTitle = True
            .ChartTitle.Text = projName & "_Planned and Order (" & Now & ")"
            .ApplyLayout (2)
            .ChartColor = 11
            .FullSeriesCollection(2).ChartType = xlLine
            .FullSeriesCollection(2).AxisGroup = 1
            
        End With
                
        
        
        
    '----------------------------------------------------------------
    lrow = Sheets(projName & "PIVOT").Cells(Rows.Count, 1).End(xlUp).Row
    
    StartPvt2 = sht.name & "!" & sht.Range("J1").Address(ReferenceStyle:=xlR1C1)
    Set pvtCache2 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)
    Set pvt2 = pvtCache.CreatePivotTable(TableDestination:=StartPvt2)
    
        pvt2.PivotFields("MRP TYPE").Orientation = xlRowField
        pvt2.AddDataField pvt.PivotFields("DELIVERED"), "Sum of Delivered", xlSum
        pvt2.AddDataField pvt.PivotFields("OPEN QTY"), "Sum of Open Qty", xlSum
    
       pvt.ManualUpdate = False
    Set pf2 = Sheets(projName & "PIVOT").PivotTables(2).PivotFields("MRP TYPE")
    
    plow = lrow + 2
    lrow = Sheets(projName & "PIVOT").Cells(Rows.Count, 1).End(xlUp).Row
    
    '=============Chart from Pivot Table ======'
    ActiveSheet.Range(Cells(plow, 10), Cells(lrow, lcol)).Select
        Set myChart = ActiveSheet.Shapes.AddChart(xlBarStacked100, 300, 410, 736, 400).Chart
    
        With myChart
            .PlotBy = xlColumns
            .ChartArea.Format.TextFrame2.TextRange.Font.Size = 8
            .HasTitle = True
            .ChartTitle.Text = projName & "_Delivered and Order (" & Now & ")"
            .ApplyLayout (2)
            .ChartColor = 11
            
        End With
        
        
        
        
        
        
End Function

Function copyConPivot()
    deleteSheet ("MergedPIVOT_Major")
    Sheets("MergedPIVOT").Copy After:=Sheets("MergedPIVOT")
    ActiveSheet.name = "MergedPIVOT_Major"
    
    lrow = Cells(Rows.Count, 1).End(xlUp).Row
    For Each pt In ActiveSheet.PivotTables
        Debug.Print (pt.name)
        With ActiveSheet.PivotTables(pt.name).PivotFields("MRP TYPE")
            .PivotItems("_TBD_").Visible = False
            .PivotItems("Bracket").Visible = False
            .PivotItems("Cable").Visible = False
            .PivotItems("Connector").Visible = False
            .PivotItems("ETC").Visible = False
            .PivotItems("Module").Visible = False
            .PivotItems("Plug").Visible = False
            .PivotItems("Regulator").Visible = False
            .PivotItems("Seal").Visible = False
            .PivotItems("Sensor").Visible = False
            .PivotItems("Silencer").Visible = False
        End With
    Next
End Function
