Attribute VB_Name = "Input_Criteria_Name"
Sub Input_Criteria_Name()

    Dim homeSheet As Worksheet
    Dim criteriaSheet As Worksheet
    Dim NumberOfCriteria As Integer
    Dim i As Integer
    
    ' Set the Home sheet
    On Error Resume Next
    Set homeSheet = ThisWorkbook.Sheets("Home")
    If homeSheet Is Nothing Then
        MsgBox "Home sheet not found!", vbCritical
        Exit Sub
    End If
    
    ' Get the number from cell J4
    NumberOfCriteria = homeSheet.Range("J4").Value
    On Error GoTo 0
    
    ' Validate the number of criteria
    If NumberOfCriteria < 3 Or NumberOfCriteria > 5 Then
        MsgBox "Please enter a valid number of criteria (3 to 5) in cell J4.", vbExclamation
        Exit Sub
    End If
    
    ' Set the appropriate sheet
    On Error Resume Next
    Set criteriaSheet = ThisWorkbook.Sheets("NumberOfCriteria-" & NumberOfCriteria)
    On Error GoTo 0
    If criteriaSheet Is Nothing Then
        MsgBox "Worksheet 'NumberOfCriteria-" & NumberOfCriteria & "' not found!", vbCritical
        Exit Sub
    End If
    
    ' Prompt user for criteria names
    For i = 1 To NumberOfCriteria
        criteriaSheet.Cells(1, i + 1).Value = InputBox("Enter the Name of Criteria " & i, "Add Criteria Name")
    Next i
    
    MsgBox "Criteria names have been updated successfully!", vbInformation

End Sub

