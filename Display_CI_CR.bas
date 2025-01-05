Attribute VB_Name = "Display_CI_CR"
Sub Display_CI_CR()


    Dim ws As Worksheet
    Dim selectedValue As Variant
    
    ' Get the value from cell J4 in "Home" sheet
    selectedValue = ThisWorkbook.Sheets("Home").Range("J4").Value

    ' Select the appropriate sheet and range
    
    Select Case selectedValue
        Case 3
            Set ws = ThisWorkbook.Sheets("NumberOfCriteria-3")
            If WorksheetIsEmpty(ws.Range("O1:O2")) Then
                MsgBox "No weights found", vbExclamation
                Exit Sub
            End If
            MsgBox "The Consistency Ratio is " & Round(ws.Range("O2").Value * 100, 2) & "%"

            
        Case 4
            Set ws = ThisWorkbook.Sheets("NumberOfCriteria-4")
            If WorksheetIsEmpty(ws.Range("O1:O2")) Then
                MsgBox "No weights found", vbExclamation
                Exit Sub
            End If
            MsgBox "The Consistency Ratio is " & Round(ws.Range("O2").Value * 100, 2) & "%"
            
        Case 5
            Set ws = ThisWorkbook.Sheets("NumberOfCriteria-5")
            If WorksheetIsEmpty(ws.Range("O1:O2")) Then
                MsgBox "No weights found", vbExclamation
                Exit Sub
            End If
            MsgBox "The Consistency Ratio is " & Round(ws.Range("O2").Value * 100, 2) & "%"
            
        Case Else
            MsgBox "Error. Please check your input.", vbCritical
            Exit Sub
    End Select

End Sub


