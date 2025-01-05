Attribute VB_Name = "Delete"
Sub Delete()
    Dim ws As Worksheet
    Dim sheetData As Variant
    Dim i As Integer

    ' Array containing sheet names and ranges to clear
    sheetData = Array( _
        Array("Home", "J4"), _
        Array("NumberOfCriteria-3", "A1:A4", "A1:D1", "A7:A9", "E7:E10", "E12:E14", "L2:L4", "O1:O2"), _
        Array("NumberOfCriteria-4", "A1:E1", "A1:A5", "A8:A13", "E8:E13", "E16:E21", "L2:L5", "O1:O2"), _
        Array("NumberOfCriteria-5", "A1:F1", "A1:A6", "A9:A18", "E9:E18", "E21:E30", "L2:L6", "O1:O2") _
    )

    On Error Resume Next ' Prevent macro from halting on missing sheets/ranges
    For i = LBound(sheetData) To UBound(sheetData)
        Set ws = ThisWorkbook.Sheets(sheetData(i)(0)) ' Get sheet by name
        If Not ws Is Nothing Then
            Dim j As Integer
            For j = 1 To UBound(sheetData(i))
                ws.Range(sheetData(i)(j)).ClearContents
            Next j
        End If
    Next i
    On Error GoTo 0 ' Reset error handling

    ' Call the external subroutine
    Call close_chart.close_chart
End Sub


