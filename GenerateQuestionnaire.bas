Attribute VB_Name = "GenerateQuestionnaire"
Sub GenerateQuestionnaire()
    Dim ws As Worksheet
    Dim homeSheet As Worksheet
    Dim criteriaRange As Range
    Dim criteriaList As Variant
    Dim resultRange As Range
    Dim numCriteria As Integer
    Dim questionCount As Integer
    Dim i As Integer, j As Integer
    Dim questionIndex As Integer
    
    ' Set the "Home" sheet and check J4 value
    Set homeSheet = ThisWorkbook.Sheets("Home")
    If IsEmpty(homeSheet.Range("J4").Value) Then
        MsgBox "Please Select Number of Criteria ", vbExclamation
        Exit Sub
    End If

    ' Determine the number of criteria from J4
    Dim numCriteriaSheet As String
    Select Case homeSheet.Range("J4").Value
        Case 3
            numCriteriaSheet = "NumberOfCriteria-3"
        Case 4
            numCriteriaSheet = "NumberOfCriteria-4"
        Case 5
            numCriteriaSheet = "NumberOfCriteria-5"
        Case Else
            MsgBox "Please Select Number of Criteria", vbExclamation
            Exit Sub
    End Select

    ' Set the worksheet and criteria/result ranges
    Set ws = ThisWorkbook.Sheets(numCriteriaSheet)
    Select Case numCriteriaSheet
        Case "NumberOfCriteria-3"
            If WorksheetIsEmpty(ws.Range("A2:A4")) Then
                MsgBox "Please Input Criteria Name", vbExclamation
                Exit Sub
            End If
            Set criteriaRange = ws.Range("A2:A4")
            Set resultRange = ws.Range("A7")
        Case "NumberOfCriteria-4"
            If WorksheetIsEmpty(ws.Range("A2:A5")) Then
                MsgBox "Please Input Criteria Name", vbExclamation
                Exit Sub
            End If
            Set criteriaRange = ws.Range("A2:A5")
            Set resultRange = ws.Range("A8")
        Case "NumberOfCriteria-5"
            If WorksheetIsEmpty(ws.Range("A2:A6")) Then
                MsgBox "Please Input Criteria Name", vbExclamation
                Exit Sub
            End If
            Set criteriaRange = ws.Range("A2:A6")
            Set resultRange = ws.Range("A9")
    End Select

    ' Get the criteria into an array
    criteriaList = criteriaRange.Value
    numCriteria = criteriaRange.Rows.Count
    questionCount = numCriteria * (numCriteria - 1) \ 2
    
    ' Clear previous results
    resultRange.Resize(questionCount, 1).ClearContents
    
    ' Generate questionnaire
    questionIndex = 0
    For i = 1 To numCriteria
        For j = i + 1 To numCriteria
            questionIndex = questionIndex + 1
            resultRange.Cells(questionIndex, 1).Value = _
                "Which is more important: " & criteriaList(i, 1) & " or " & criteriaList(j, 1) & "?"
        Next j
    Next i
    
    MsgBox "Questionnaire Generated successfully"
End Sub

Function WorksheetIsEmpty(rng As Range) As Boolean
    Dim cell As Range
    For Each cell In rng
        If Not IsEmpty(cell.Value) Then
            WorksheetIsEmpty = False
            Exit Function
        End If
    Next cell
    WorksheetIsEmpty = True
End Function


