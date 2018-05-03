Attribute VB_Name = "Module1"
Public Function CountPivotTables() As Integer
    Dim ws As Worksheet
    Dim i As Integer
    i = 0
    For Each ws In ActiveWorkbook.Worksheets
        For Each PivotTable In ws.PivotTables
            i = i + 1
        Next PivotTable
    Next ws
    Debug.Print i, "pivot table"
    
    CountPivotTables = i
    

End Function
Public Function addSheet(mySheetName)
On Error Resume Next
    mySheetNameTest = Worksheets(mySheetName).Name
    If Err.Number = 0 Then
        'MsgBox "The sheet named ''" & mySheetName & "'' DOES exist in this workbook."
    Else
        Err.Clear
        Worksheets.Add.Name = mySheetName
        'MsgBox "The sheet named ''" & mySheetName & "'' did not exist in this workbook but it has been created now."
    End If
End Function
Public Function deleteSheet(mySheetName)
Application.DisplayAlerts = False
    On Error Resume Next
    Sheets(mySheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Function


Public Function readLOT(link)
       
    projName = Mid(link, 80, 3)
       
    Dim wkbMyWorkbook As Workbook
    Dim wkbWebWorkbook As Workbook
    Dim wksWebWorkSheet As Worksheet
    
    Set wkbMyWorkbook = ActiveWorkbook
    Workbooks.Open (link)
    Set wkbWebWorkbook = ActiveWorkbook
    Set wksWebWorkSheet = ActiveSheet
    
    wkbWebWorkbook.Sheets("Overview purchase order").Copy after:=wkbMyWorkbook.Sheets("MAIN")
    ActiveSheet.UsedRange.Copy
    addSheet (projName & "SO")
    Cells(1, 1).PasteSpecial xlPasteValues
    deleteSheet ("Overview purchase order")
           
    wkbWebWorkbook.Sheets("BOM set (inner comp. transfer)").Copy after:=wkbMyWorkbook.Sheets("MAIN")
    ActiveSheet.UsedRange.Copy
    addSheet (projName & "HQ")
    Cells(1, 1).PasteSpecial xlPasteValues
    deleteSheet ("BOM set (inner comp. transfer)")

    wkbWebWorkbook.Close savechanges:=False
End Function

Public Function removeDuplicates(projNo)
    Dim MyRange As Range
    Dim LastRow As Long
    
    addSheet (projNo & "QTY")
    lrow = Sheets(projNo & "HQ").Cells(Rows.Count, 10).End(xlUp).Row
    lrow2 = Sheets(projNo & "SO").Cells(Rows.Count, 5).End(xlUp).Row
    lrowPartList = Sheets("MajorParts").Cells(Rows.Count, 1).End(xlUp).Row
    
    Sheets(projNo & "QTY").Activate
    With ActiveSheet
       
       'A_PART#
       .Range(.Cells(2, 1), .Cells(lrow, 1)).FormulaR1C1 = _
       "='" & projNo & "HQ'!RC10"
       
       'B_DESCRIPTION
       .Range(.Cells(2, 2), .Cells(lrow, 2)).FormulaR1C1 = _
       "=VLOOKUP(RC1,'" & projNo & "HQ'!R7C10:R" & lrow & "C11,2,FALSE)"
       
       'C_MRP TYPE
       .Range(.Cells(2, 3), .Cells(lrow, 3)).FormulaR1C1 = _
       "=iferror(VLOOKUP(RC1,'MajorParts'!R1C1:R" & lrowPartList & "C3,3,FALSE),"""")"
              
       'D_PLANNED
       .Range(.Cells(2, 4), .Cells(lrow, 4)).FormulaR1C1 = _
       "=INT(SUMIF('" & projNo & "HQ'!R7C10:R" & lrow & "C10, RC1,'" & projNo & "HQ'!R7C18:R" & lrow & "C18))"
       
       'E_ORDERED
       .Range(.Cells(2, 5), .Cells(lrow, 5)).FormulaR1C1 = _
       "=INT(SUMIF('" & projNo & "SO'!R5C7:R" & lrow2 & "C7, RC1,'" & projNo & "SO'!R5C10:R" & lrow2 & "C10))"
       
       'F_TO ORDER
       .Range(.Cells(2, 6), .Cells(lrow, 6)).FormulaR1C1 = _
       "=RC4-RC5"
       
       'G_DELIEVERED
       .Range(.Cells(2, 7), .Cells(lrow, 7)).FormulaR1C1 = _
       "=RC5-RC8"
       
       'H_OPEN QTY
       .Range(.Cells(2, 8), .Cells(lrow, 8)).FormulaR1C1 = _
       "=INT(SUMIF('" & projNo & "SO'!R5C7:R" & lrow2 & "C7, RC1,'" & projNo & "SO'!R5C11:R" & lrow2 & "C11))"
       
    End With
    
       
    
'
'
'    Dim rng As Range, cell As Range
'    Set rng = Sheets(projNo & "HQ").Range("KB6:MN6")
'    Dim i As Integer
'
'    lcol = Sheets(projNo & "QTY").Cells(1, Columns.Count).End(xlToLeft).Column
'
'    i = lcol + 1
'
'    For Each cell In rng
'
'    If cell.Value Like "*SUM*" Then
'        Debug.Print "LOOP", cell.Column
'
'
'        Sheets(projNo & "HQ").Cells(1, cell.Column).EntireColumn.Copy
'        Sheets(projNo & "QTY").Cells(1, i).PasteSpecial xlPasteValues
'        i = i + 1
'        Debug.Print "COPIED", i
'
'
'        End If
'    Next cell
'
    
    ActiveSheet.Columns.AutoFit
    LastRow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
    
    Set MyRange = ActiveSheet.Range("A1:H" & LastRow)
    MyRange.removeDuplicates Columns:=1, Header:=xlYes
    
    
    Rows(lrow + 1 & ":100").EntireRow.Delete
    
    
    
    Rows("1:5").EntireRow.Delete
    
        Cells(1, 1) = "PART#"
        Cells(1, 2) = "DESCRIPTION"
        Cells(1, 3) = "MRP TYPE"
        Cells(1, 4) = "PLANNED"
        Cells(1, 5) = "ORDERED"
        Cells(1, 6) = "TO ORDER"
        Cells(1, 7) = "DELIVERED"
        Cells(1, 8) = "OPEN QTY"
    'Sheets("MAIN").Activate
End Function


Function deleteSheets()



Dim s As Worksheet, t As String
    Dim i As Long, K As Long
    K = Sheets.Count
    
    For i = K To 1 Step -1
        t = Sheets(i).Name
        'Or InStr(1, t, "pivot", 1)
        If InStr(1, t, "Sheet", 1) _
            Or InStr(1, t, "KrCon", 1) _
            Or InStr(1, t, "Copy", 1) _
            Or InStr(1, t, "QTY", 1) _
            Or InStr(1, t, "HQ", 1) _
            Or InStr(1, t, "SO", 1) _
            Or InStr(1, t, "time", 1) _
        Then
        
            
            Application.DisplayAlerts = False
                Sheets(i).Delete
            Application.DisplayAlerts = True
        End If
    Next i
    
End Function




Public Function assignLink(projNo) As String
Dim link As String
Select Case projNo
    Case 482
        assignLink = "http://sps.emea.festo.net/sites/LOrders/inquiries/Shared%20Documents/LO_482/LO_482.xlsx"
    Case 480
        assignLink = "http://sps.emea.festo.net/sites/LOrders/inquiries/Shared%20Documents/LO_480/LO_480.xlsx"
    Case 477
        assignLink = "http://sps.emea.festo.net/sites/LOrders/inquiries/Shared%20Documents/LO_477/LO_477_HYUNDAI_Dymos_Q-Project.xlsx"
    Case 460
        assignLink = "http://sps.emea.festo.net/sites/LOrders/inquiries/Shared%20Documents/LO_460/LO_460_CWA%208.xlsx"
    Case 459
        assignLink = "http://sps.emea.festo.net/sites/LOrders/inquiries/Shared%20Documents/LO_459/LO_459_CNA%205.xlsx"
    Case 458
        assignLink = "http://sps.emea.festo.net/sites/LOrders/inquiries/Shared%20Documents/LO_458/LO_458_CWA%205_6_7.xlsx"
    Case 456
        assignLink = "http://sps.emea.festo.net/sites/LOrders/inquiries/Shared%20Documents/LO_456/LO_456_Hot%20Press.xlsx"
    Case 455
        assignLink = "http://sps.emea.festo.net/sites/LOrders/inquiries/Shared%20Documents/LO_455/LO_455_LGD%20GP3.xlsx"
    Case 440
        assignLink = "http://sps.emea.festo.net/sites/LOrders/inquiries/Shared%20Documents/LO_440/LO_440_LGC-Poland_MEB%20Phase%201.xlsx"
    Case 413
        assignLink = "http://sps.emea.festo.net/sites/LOrders/inquiries/Shared%20Documents/LO_413/LO_413%20BOE%20B7%20Phase3.xlsx"
    Case 391
        assignLink = "http://sps.emea.festo.net/sites/LOrders/inquiries/Shared%20Documents/LO_391/LO_391_CTS%20S%20LD.xlsx"
    Case 298
        assignLink = "http://sps.emea.festo.net/sites/LOrders/inquiries/Shared%20Documents/LO_298/LO_298_HKC%20Phase2.xlsx"
    Case 285
        assignLink = "http://sps.emea.festo.net/sites/LOrders/inquiries/Shared%20Documents/LO_285/LO_285_CVC%20Door%20open%20system.xlsx"

   End Select
   
   
   
End Function
