Attribute VB_Name = "Module2"

'===================================================
' (!) TimeSeries to be developed
'===================================================


Public Function timeSeries(projNo)

Dim rng As Range, cell As Range

Set rng = Sheets(projNo & "HQ").Range("A6:RO6")
Dim i As Integer

deleteSheet (projNo & "Time")
addSheet (projNo & "Time")

'last row of the partnumber in HQ worksheet
lrow = Sheets(projNo & "HQ").Cells(Rows.Count, 10).End(xlUp).Row
Debug.Print "row count of HQ", lrow


Sheets(projNo & "HQ").Range("J6:M" & lrow).Copy
Sheets(projNo & "Time").Cells(6, 1).PasteSpecial xlPasteValues

 'C_MRP TYPE
 lrowPartlist = Sheets("MajorParts").Cells(Rows.Count, 3).End(xlUp).Row
 ActiveSheet.Range(Cells(2, 5), Cells(lrow, 5)).FormulaR1C1 = _
 "=VLOOKUP(RC1,'MajorParts'!R1C1:R" & lrowPartlist & "C3,3,FALSE)"
              

lcol = Cells(6, Columns.Count).End(xlToLeft).Column
i = lcol + 1

For Each cell In rng
    If cell.Value Like "*SUM*" Then
        Debug.Print "LOOP", cell.Column
        cell.Offset(-1, 0).FormulaR1C1 = "=sum(R7C" & cell.Column & ":R" & lrow & "C" & cell.Column & ")"
        
        Sheets(projNo & "HQ").Cells(1, cell.Column).EntireColumn.Copy
        Sheets(projNo & "Time").Cells(1, i).PasteSpecial xlPasteValues
        i = i + 1
        With ActiveSheet
            .Range(Cells(7, i), Cells(lrow, i)).FormulaR1C1 = _
            "=sumif(R7C1:R" & lrow & "C1, RC1, R7C" & i - 1 & ":R" & lrow & "C" & i - 1 & ")"
            
        End With
        
        Cells(5, i).FormulaR1C1 = _
            "=sum(R7C" & i & ":R" & lrow & "C" & i & ")"
        
        
        
        i = i + 1
        
        
        Debug.Print "COPIED", i

            
    End If
Next cell
Debug.Print lrow + 1

       

lrow2 = Cells(Rows.Count, 20).End(xlUp).Row
Rows(lrow + 1 & ":" & lrow2).EntireRow.Delete




End Function


Public Function tidyUpTimeSeries(projNo)
Debug.Print "tidyupTimeseries() called"
ActiveSheet.UsedRange.Value = ActiveSheet.UsedRange.Value


    lcol = Cells(6, Columns.Count).End(xlToLeft).Column
    For i = lcol To 6 Step -1
        Debug.Print i
        If Cells(5, i).Value = 0 Then
            Cells(5, i).EntireColumn.Delete
        End If
    Next i

Rows("1:5").EntireRow.Delete

lrow = Cells(Rows.Count, 1).End(xlUp).Row
lcol = Cells(2, Columns.Count).End(xlToLeft).Column
Set MyRange = ActiveSheet.Range(Cells(1, 1), Cells(lrow, lcol))
MyRange.removeDuplicates Columns:=1, Header:=xlYes


End Function


