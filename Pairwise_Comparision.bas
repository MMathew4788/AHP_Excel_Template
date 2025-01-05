Attribute VB_Name = "Pairwise_Comparision"

Sub Pairwise_Comparision()
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
            If WorksheetIsEmpty(ws.Range("A12:A14")) Then
                MsgBox "Please Generate Questionnaire", vbExclamation
                Exit Sub
            End If
            Set questionRange = ws.Range("A12:A14") ' Adjust range based on the max questions
            Set resultRange = ws.Range("E12:E14")
        Case 4
            Set ws = ThisWorkbook.Sheets("NumberOfCriteria-4")
            If WorksheetIsEmpty(ws.Range("A16:A21")) Then
                MsgBox "Please Generate Questionnaire", vbExclamation
                Exit Sub
            End If
            Set questionRange = ws.Range("A16:A21") ' Adjust range based on the max questions
            Set resultRange = ws.Range("E16:E21")
        Case 5
            Set ws = ThisWorkbook.Sheets("NumberOfCriteria-5")
            If WorksheetIsEmpty(ws.Range("A21:A30")) Then
                MsgBox "Please Generate Questionnaire", vbExclamation
                Exit Sub
            End If
            Set questionRange = ws.Range("A21:A30") ' Adjust range based on the max questions
            Set resultRange = ws.Range("E21:E30")
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
        
        ' Load the question into the UserForm
        With UserForm1
            .lblQuestion.Caption = questionRange.Cells(i, 1).Value ' Set the question text
            .cmbOptions.Clear ' Clear any existing items in the dropdown
            
            ' Add the two criteria to the dropdown
            .cmbOptions.AddItem "Equal Importance"
            .cmbOptions.AddItem "Equal to Moderate Importance"
            .cmbOptions.AddItem "Moderate Importance"
            .cmbOptions.AddItem "Moderate to Strong Importance"
            .cmbOptions.AddItem "Strong Importance"
            .cmbOptions.AddItem "Strong to Very Strong Importance"
            .cmbOptions.AddItem "Very Strong Importance"
            .cmbOptions.AddItem "Very Strong  to Extreme Importance"
            .cmbOptions.AddItem "Extreme Importance"
            
            .cmbOptions.ListIndex = -1 ' Clear selection
            .Show ' Display the UserForm
            
            ' Get the user's answer
            userAnswer = .cmbOptions.Value
        End With
        
        ' Save the answer in the result range
        resultRange.Cells(i, 1).Value = userAnswer
    Next i
    
    MsgBox "Pairwise comparison Saved Successfully", vbInformation
End Sub

