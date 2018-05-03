Attribute VB_Name = "Module5"
''''''''''''''''''''''''''''''''''''''''''''''
' Drive File to run theprogram

''''''''''''''''''''''''''''''''''''''''''''''



Sub readfile(projnum)
    link = assignLink(projnum)
    readLOT (link) 'read large order template
    removeDuplicates (projnum)
    CreatePivotTable (projnum)
    timeSeries (projnum)
    tidyUpTimeSeries (projnum)

End Sub


Sub read_all()
    deleteSheets
    projnum = Mid(ThisWorkbook.Name, 3, 3)
    readfile (projnum)
    Dim projlist(1 To 5) As Integer
    projlist(1) = 482
    projlist(2) = 480
    projlist(3) = 477
    projlist(4) = 460
    projlist(5) = 459
    
    For projnum = 1 To UBound(projlist):
            readfile (projlist(projnum))
    Next projnum
    
End Sub


Sub blocktest()
    deleteSheets
    readfile (482)
    
End Sub

