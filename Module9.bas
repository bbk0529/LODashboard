Attribute VB_Name = "Module9"
'=============================================
' Drive File to run the program
'=============================================

Sub readfile(projnum)

    'check for filter, turn on if none exists
    Sheets("MajorParts").AutoFilterMode = False

    link = assignLink(projnum) 'mapping function from project number to sharepointlink
    readLOT (link) 'read large order template from sharepoint
    removeDuplicates (projnum)
    CreatePivotTable (projnum)
    
    '==========================
    'to be developed further
    'timeSeries (projnum)
    'tidyUpTimeSeries (projnum)
    '==========================

End Sub

Sub read_all()
    
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    StartTime = Timer
    
    deleteSheets
    
    lrow = Sheets("MAIN").Cells(Rows.Count, 1).End(xlUp).Row
       
    For i = 1 To lrow
            projnum = Sheets("MAIN").Cells(i, 1).Value
            readfile (projnum)
    Next i
    
    SecondsElapsed = Round(Timer - StartTime, 2)
    Debug.Print "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation
    deleteSheetsWithName ("SO")
    deleteSheetsWithName ("HQ")
    'deleteSheetsWithName ("QTY")
    Sheets("MAIN").Activate
    
End Sub

Sub runConPivot()
    createConsolidationQTY
    CreateConsolidationPivotTable
    copyConPivot
End Sub


Sub deleteAll()
    deleteSheets
End Sub

