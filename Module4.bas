Attribute VB_Name = "Module4"
Function CreatePivotTable(projName)

Dim sht As Worksheet
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim StartPvt As String
Dim SrcData As String
Dim sheetName As String

 
deleteSheet (projName & "PIVOT")
 
lrow = Sheets(projName & "QTY").Cells(Rows.Count, 1).End(xlUp).Row
lcol = Sheets(projName & "QTY").Cells(1, Columns.Count).End(xlToLeft).Column

SrcData = Sheets(projName & "QTY").Name & "!" & Range("A1:H" & lrow).Address(ReferenceStyle:=xlR1C1)
  
   
Set sht = Sheets.Add
ActiveSheet.Name = projName & "PIVOT"


'----------------------------------------------------------------
StartPvt = sht.Name & "!" & sht.Range("A1").Address(ReferenceStyle:=xlR1C1)
Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)
Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt)

    'pvt.PivotFields("LO number").Orientation = xlPageField
    'pvt.PivotFields("LOT project status").Orientation = xlColumnField
    pvt.PivotFields("MRP TYPE").Orientation = xlRowField
    pvt.AddDataField pvt.PivotFields("PLANNED"), "Sum of Planned", xlSum
    pvt.AddDataField pvt.PivotFields("ORDERED"), "Sum of Ordered", xlSum

   pvt.ManualUpdate = False
Set pf = Sheets(projName & "PIVOT").PivotTables(1).PivotFields("MRP TYPE")

ActiveSheet.Range(Cells(2, 1), Cells(lrow, lcol)).Select
    Set myChart = ActiveSheet.Shapes.AddChart(xlColumnClustered, 300, 10, , 175).Chart

    With myChart
        .PlotBy = xlColumns
        .ChartArea.Format.TextFrame2.TextRange.Font.Size = 8
        .HasTitle = True
        .ChartTitle.Text = "Planned and Order (" & Now & ")"
        .ApplyLayout (5)
    End With
    
    
'----------------------------------------------------------------
lrow = Sheets(projName & "PIVOT").Cells(Rows.Count, 1).End(xlUp).Row

StartPvt2 = sht.Name & "!" & sht.Range("J1").Address(ReferenceStyle:=xlR1C1)
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
    Set myChart = ActiveSheet.Shapes.AddChart(xlBarStacked100, 300, 210, , 175).Chart

    With myChart
        .PlotBy = xlColumns
        .ChartArea.Format.TextFrame2.TextRange.Font.Size = 8
        .HasTitle = True
        .ChartTitle.Text = "Delivered and Order (" & Now & ")"
        .ApplyLayout (5)
    End With


End Function
