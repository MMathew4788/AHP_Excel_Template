Attribute VB_Name = "Criteria_Importance"
Sub Criteria_Importance()
    Dim valueInJ4 As Integer
    Dim ws As Worksheet
    Dim questionRange As Range
    Dim resultRange As Range
    Dim questionCount As Integer
    Dim i As Integer
    Dim userAnswer As String
    Dim questionParts As Variant
    Dim criteria1 As String, criteria2 As String
    
    'check J4 in Home
    If IsEmpty(ThisWorkbook.Sheets("Home").Range("J4").Value) Then
        MsgBox "Please Select Number of Criteria ", vbExclamation
        Exit Sub
    End If
    
    
    ' Determine which sheet and ranges to use
    Select Case ThisWorkbook.Sheets("Home").Range("J4").Value
        Case 3
            Set ws = ThisWorkbook.Sheets("NumberOfCriteria-3")
            If WorksheetIsEmpty(ws.Range("A7:A10")) Then
                MsgBox "Please Generate Questionnaire", vbExclamation
                Exit Sub
            End If
            Set questionRange = ws.Range("A7:A10") ' Adjust range based on the max questions
            Set resultRange = ws.Range("E7:E10")
        Case 4
            Set ws = ThisWorkbook.Sheets("NumberOfCriteria-4")
            If WorksheetIsEmpty(ws.Range("A8:A13")) Then
                MsgBox "Please Generate Questionnaire", vbExclamation
                Exit Sub
            End If
            Set questionRange = ws.Range("A8:A13") ' Adjust range based on the max questions
            Set resultRange = ws.Range("E8:E13")
        Case 5
            Set ws = ThisWorkbook.Sheets("NumberOfCriteria-5")
            If WorksheetIsEmpty(ws.Range("A9:A18")) Then
                MsgBox "Please Generate Questionnaire", vbExclamation
                Exit Sub
            End If
            Set questionRange = ws.Range("A9:A18") ' Adjust range based on the max questions
            Set resultRange = ws.Range("E9:E18")
        Case Else
            MsgBox "Error. Please check your input.", vbCritical
            Exit Sub ' Exit if something goes wrong
    End Select
    
    ' Clear previous results
    resultRange.ClearContents
    
    ' Get the number of questions
    questionCount = Application.WorksheetFunction.CountA(questionRange)
    
    ' Loop through questions
    For i = 1 To questionCount
        ' Extract criteria from the current question
        questionParts = Split(questionRange.Cells(i, 1).Value, ":")
        criteria1 = Trim(Split(questionParts(1), " or ")(0)) ' Extract the first criterion
        criteria2 = Trim(Split(questionParts(1), " or ")(1)) ' Extract the second criterion
        
        ' Remove any trailing "?" from the second criterion
        If Right(criteria2, 1) = "?" Then
            criteria2 = Left(criteria2, Len(criteria2) - 1)
        End If
        
        ' Load the question into the UserForm
        With UserForm1
            .lblQuestion.Caption = questionRange.Cells(i, 1).Value ' Set the question text
            .cmbOptions.Clear ' Clear any existing items in the dropdown
            
            ' Add the two criteria to the dropdown
            .cmbOptions.AddItem criteria1
            .cmbOptions.AddItem criteria2
            
            .cmbOptions.ListIndex = -1 ' Clear selection
            .Show ' Display the UserForm
            
            ' Get the user's answer
            userAnswer = .cmbOptions.Value
        End With
        
        ' Save the answer in the result range
        resultRange.Cells(i, 1).Value = userAnswer
    Next i
    
    MsgBox "Criteria Importance Saved Successfully", vbInformation
End Sub
