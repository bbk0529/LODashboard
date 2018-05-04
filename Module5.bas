Attribute VB_Name = "Module5"
'=============================================
' Drive File to run the program
'=============================================

Sub readfile(projnum)
    'filterOff ("MajorParts") 'filter off if so.
    link = assignLink(projnum) 'mapping function from project number to sharepointlink
    readLOT (link) 'read large order template from sharepoint
    removeDuplicates (projnum)
    CreatePivotTable (projnum)
    timeSeries (projnum)
    'tidyUpTimeSeries (projnum)

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
    MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation
    
End Sub


Sub blocktest()
    deleteSheets
    read_all
    'deleteSheetsWithName ("Time")
    'deleteSheetsWithName ("HQ")
    'deleteSheetsWithName ("SO")
    'deleteSheetsWithName ("QTY")
    
End Sub

