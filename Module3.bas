Attribute VB_Name = "Module3"
'=======================================
' PROJECT QTY sheets made
'=======================================
Public Function a()
    removeDuplicates (459)
End Function

Public Function removeDuplicates(projNo)
    Dim MyRange As Range
    Dim LastRow As Long
    
    addSheet (projNo & "QTY")
    lrow = Sheets(projNo & "HQ").Cells(Rows.Count, 10).End(xlUp).Row
    lrow2 = Sheets(projNo & "SO").Cells(Rows.Count, 5).End(xlUp).Row
    lrowPartlist = Sheets("MajorParts").Cells(Rows.Count, 1).End(xlUp).Row
    
    Debug.Print "lrow : ", lrow, "lrow2 : ", lrow2, "lrowpartlist : ", lrowPartlist
    
    
    Sheets(projNo & "QTY").Activate

    lrow = Sheets(projNo & "HQ").Cells(Rows.Count, 10).End(xlUp).Row


    With ActiveSheet
       
        'A_PART# from HQ
        .Range(.Cells(2, 1), .Cells(lrow, 1)).FormulaR1C1 = _
        "='" & projNo & "HQ'!RC10"
       
        'A_PART# from SO
        If lrow2 > 4 Then
            .Range(.Cells(lrow + 1, 1), .Cells((lrow + 1) + (lrow2 - 4) - 1, 1)).FormulaR1C1 = _
            "='" & projNo & "SO'!R[" & -lrow - 1 + 5 & "]C7"
            Debug.Print "============!!!", (-lrow + 5)
        End If
        
        
        'J_PART# from HQ
        .Range(.Cells(2, 9), .Cells(lrow, 9)).FormulaR1C1 = _
        "='" & projNo & "HQ'!RC8"
        
        qtyrow = Sheets(projNo & "QTY").Cells(Rows.Count, 1).End(xlUp).Row
               
       
       'B_DESCRIPTION
       .Range(.Cells(2, 2), .Cells(qtyrow, 2)).FormulaR1C1 = _
       "=iferror(VLOOKUP(RC1,'MajorParts'!R1C1:R" & lrowPartlist & "C2,2,FALSE),"""")"
       '"=VLOOKUP(RC1,'" & projNo & "HQ'!R7C10:R" & lrow & "C11,2,FALSE)"
       
       'C_MRP TYPE
       .Range(.Cells(2, 3), .Cells(qtyrow, 3)).FormulaR1C1 = _
       "=iferror(VLOOKUP(RC1,'MajorParts'!R1C1:R" & lrowPartlist & "C3,3,FALSE),"""")"
              
       'D_PLANNED
       .Range(.Cells(2, 4), .Cells(qtyrow, 4)).FormulaR1C1 = _
       "=INT(SUMIF('" & projNo & "HQ'!R7C10:R" & lrow & "C10, RC1,'" & projNo & "HQ'!R7C18:R" & lrow & "C18))"
       
       'E_ORDERED
       .Range(.Cells(2, 5), .Cells(qtyrow, 5)).FormulaR1C1 = _
       "=INT(SUMIF('" & projNo & "SO'!R5C7:R" & lrow2 & "C7, RC1,'" & projNo & "SO'!R5C10:R" & lrow2 & "C10))"
       
       'F_TO ORDER
       .Range(.Cells(2, 6), .Cells(qtyrow, 6)).FormulaR1C1 = _
       "=RC4-RC5"
       
       'G_DELIEVERED
       .Range(.Cells(2, 7), .Cells(qtyrow, 7)).FormulaR1C1 = _
       "=RC5-RC8"
       
       'H_OPEN QTY
       .Range(.Cells(2, 8), .Cells(qtyrow, 8)).FormulaR1C1 = _
       "=INT(SUMIF('" & projNo & "SO'!R5C7:R" & lrow2 & "C7, RC1,'" & projNo & "SO'!R5C11:R" & lrow2 & "C11))"
       
    End With
    
     
   
    
    
'========================
'Remove duplicate
'========================
    LastRow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
    Set MyRange = ActiveSheet.Range("A1:I" & LastRow)
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

    ActiveSheet.Columns.AutoFit
    Set CopyRng = ActiveSheet.UsedRange
    CopyRng.Copy
    With ActiveSheet.Cells(Last + 1, "A")
    .PasteSpecial xlPasteValues
    .PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
End With
End Function

