Attribute VB_Name = "Module1"
'=========================================================
' Utility function prepared here
'=========================================================



Public Function addSheet(mySheetName)
On Error Resume Next
    mySheetNameTest = Worksheets(mySheetName).name
    If Err.Number = 0 Then
        'MsgBox "The sheet named ''" & mySheetName & "'' DOES exist in this workbook."
    Else
        Err.Clear
        Worksheets.Add.name = mySheetName
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


Function deleteSheetsWithName(name)
    Dim s As Worksheet, t As String
    Dim i As Long, K As Long
    K = Sheets.Count
    
    For i = K To 1 Step -1
        t = Sheets(i).name
        
        If InStr(1, t, name, 1) _
        Then
            Application.DisplayAlerts = False
                Sheets(i).Delete
            Application.DisplayAlerts = True
        End If
    Next i
    
End Function

Function deleteSheets()
    Dim s As Worksheet, t As String
    Dim i As Long, K As Long
    K = Sheets.Count
    
    For i = K To 1 Step -1
        t = Sheets(i).name
        
        If InStr(1, t, "Sheet", 1) _
            Or InStr(1, t, "KrCon", 1) _
            Or InStr(1, t, "Copy", 1) _
            Or InStr(1, t, "QTY", 1) _
            Or InStr(1, t, "HQ", 1) _
            Or InStr(1, t, "SO", 1) _
            Or InStr(1, t, "time", 1) _
            Or InStr(1, t, "pivot", 1) _
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

'=========================================
' Read file from Sharepoint
''=========================================

Public Function readLOT(link)
    projName = Mid(link, 80, 3)
    Dim wkbMyWorkbook As Workbook
    Dim wkbWebWorkbook As Workbook
    Dim wksWebWorkSheet As Worksheet
    
    Set wkbMyWorkbook = ActiveWorkbook
    Workbooks.Open (link)
    Set wkbWebWorkbook = ActiveWorkbook
    Set wksWebWorkSheet = ActiveSheet
    
    'projName SO
    wkbWebWorkbook.Sheets("Overview purchase order").Copy After:=wkbMyWorkbook.Sheets("MAIN")
    ActiveSheet.UsedRange.Copy
    addSheet (projName & "SO")
    Cells(1, 1).PasteSpecial xlPasteValues
    deleteSheet ("Overview purchase order")
           
    'projName HQ
    wkbWebWorkbook.Sheets("BOM set (inner comp. transfer)").Copy After:=wkbMyWorkbook.Sheets("MAIN")
    ActiveSheet.UsedRange.Copy
    addSheet (projName & "HQ")
    Cells(1, 1).PasteSpecial xlPasteValues
    deleteSheet ("BOM set (inner comp. transfer)")

    wkbWebWorkbook.Close savechanges:=False
End Function



Function initializeProjectProgress(SheetName)
    If Sheets(SheetName).AutoFilterMode Then
    If Sheets(SheetName).FilterMode Then
        Sheets(SheetName).ShowAllData
    End If

End Function



Function LastRow(sh As Worksheet)
    On Error Resume Next
    LastRow = sh.Cells.Find(What:="*", _
                            After:=sh.Range("A1"), _
                            Lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row
    On Error GoTo 0
End Function

Function LastCol(sh As Worksheet)
    On Error Resume Next
    LastCol = sh.Cells.Find(What:="*", _
                            After:=sh.Range("A1"), _
                            Lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
    On Error GoTo 0
End Function




