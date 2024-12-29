Attribute VB_Name = "Define_Number_of_Criteria"
Sub Define_Number_of_Criteria()
    Dim criteriaNumber As Variant
    Dim ws As Worksheet
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Home")
    
    Do
        ' Prompt the user for the number of criteria
        criteriaNumber = Application.InputBox("What is the number of criteria? (Choose 3, 4, or 5)", "Criteria Input", Type:=1)
        
        ' Check if the user clicked Cancel
        If criteriaNumber = False Then
            MsgBox "Please enter number of criteria", vbExclamation, "Canceled"
            Exit Sub
        End If
        
        ' Check if the input is valid (3, 4, or 5)
        If criteriaNumber = 3 Or criteriaNumber = 4 Or criteriaNumber = 5 Then
            Exit Do ' Exit the loop if valid
        Else
            MsgBox "Invalid input. Please enter 3, 4, or 5.", vbCritical, "Invalid Input"
        End If
    Loop
    
    ' Assign the input value to cell J4
    ws.Range("J4").Value = criteriaNumber
    
    MsgBox "The number of criteria has been set to " & criteriaNumber, vbInformation, "Number of Criteria"
    
End Sub

