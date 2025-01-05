Attribute VB_Name = "Calculate_AHP"
Sub Calculate_AHP()
    Dim wsHome As Worksheet
    Dim criteriaSheet As Worksheet
    Dim numCriteria As Integer
    Dim tempRI As Variant
    Dim RI(1 To 15) As Double
    Dim X() As Double
    Dim Xn() As Double
    Dim W() As Double
    Dim WSV() As Double
    Dim Lamda() As Double
    Dim Lamdamax As Double
    Dim CI As Double
    Dim CR As Double
    Dim i As Integer, j As Integer
    Dim totalSum As Double
    
    ' Set references
    Set wsHome = ThisWorkbook.Sheets("Home")
    
    'check J4 in Home
    If IsEmpty(ThisWorkbook.Sheets("Home").Range("J4").Value) Then
        MsgBox "Please Select Number of Criteria ", vbExclamation
        Exit Sub
    End If
    
    ' Read the number of criteria from Sheet("Home"), cell J4
    numCriteria = wsHome.Range("J4").Value
    
    ' Load the appropriate sheet based on the number of criteria
    On Error Resume Next
    Set criteriaSheet = ThisWorkbook.Sheets("NumberOfCriteria-" & numCriteria)
    On Error GoTo 0
    If criteriaSheet Is Nothing Then
        MsgBox "No sheet found for NumberOfCriteria-" & numCriteria, vbExclamation
        Exit Sub
    End If
    
    ' Temporary array with RI values
    tempRI = Array(0, 0, 0.58, 0.9, 1.12, 1.24, 1.32, 1.41, 1.45, 1.49, 1.51, 1.54, 1.56, 1.57, 1.58)

    ' Populate RI array
    For i = LBound(RI) To UBound(RI)
        RI(i) = tempRI(i - 1) ' Adjust for 0-based Array index in tempRI
    Next i

    ' Read matrix X from the criteriaSheet
    ReDim X(1 To numCriteria, 1 To numCriteria)
    For i = 1 To numCriteria
        For j = 1 To numCriteria
            X(i, j) = criteriaSheet.Cells(i + 1, j + 1).Value ' Assuming data starts from B2
        Next j
    Next i
    
    ' Normalize the matrix X (Xn)
    ReDim Xn(1 To numCriteria, 1 To numCriteria)
    For j = 1 To numCriteria
        totalSum = 0
        For i = 1 To numCriteria
            totalSum = totalSum + X(i, j)
        Next i
        For i = 1 To numCriteria
            Xn(i, j) = X(i, j) / totalSum
        Next i
    Next j
    
    ' Calculate weights (W)
    ReDim W(1 To numCriteria)
    For i = 1 To numCriteria
        W(i) = 0
        For j = 1 To numCriteria
            W(i) = W(i) + Xn(i, j)
        Next j
        W(i) = W(i) / numCriteria
    Next i
    
    ' Calculate Weighted Sum Vector (WSV)
    ReDim WSV(1 To numCriteria)
    For i = 1 To numCriteria
        WSV(i) = 0
        For j = 1 To numCriteria
            WSV(i) = WSV(i) + X(i, j) * W(j)
        Next j
    Next i
    
    ' Calculate Lambda (Lamda)
    ReDim Lamda(1 To numCriteria)
    For i = 1 To numCriteria
        Lamda(i) = WSV(i) / W(i)
    Next i
    
    ' Calculate Lambda Max
    Lamdamax = 0
    For i = 1 To numCriteria
        Lamdamax = Lamdamax + Lamda(i)
    Next i
    Lamdamax = Lamdamax / numCriteria
    
    ' Calculate Consistency Index (CI) and Consistency Ratio (CR)
    CI = (Lamdamax - numCriteria) / (numCriteria - 1)
    CR = CI / RI(numCriteria)
    
    ' Output results to Home sheet

    For i = 1 To numCriteria
        criteriaSheet.Cells(1 + i, 12).Value = W(i) ' Output weights in column L
    Next i
    criteriaSheet.Range("O1").Value = CI
    criteriaSheet.Range("O2").Value = CR
    
    ' Notify user
    MsgBox "AHP Calculation completed successfully!", vbInformation
End Sub

