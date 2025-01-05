Attribute VB_Name = "Create_Chart"
Sub Create_Chart()

    Dim chrt As ChartObject
    Dim ws As Worksheet
    Dim selectedValue As Variant
    
    ' Add chart to Sheet1
    Set chrt = Sheets("Home").ChartObjects.Add(Left:=550, Width:=300, Top:=20, Height:=200)
    
    ' Get the value from cell J4 in "Home" sheet
    selectedValue = ThisWorkbook.Sheets("Home").Range("J4").Value

    ' Select the appropriate sheet and range
    On Error GoTo ErrorHandler
    Select Case selectedValue
        Case 3
            Set ws = ThisWorkbook.Sheets("NumberOfCriteria-3")
            If WorksheetIsEmpty(ws.Range("L2:L4")) Then
                MsgBox "No weights found", vbExclamation
                Exit Sub
            End If
            chrt.Chart.SetSourceData Source:=ws.Range("K2:L4")
            
        Case 4
            Set ws = ThisWorkbook.Sheets("NumberOfCriteria-4")
            If WorksheetIsEmpty(ws.Range("L2:L5")) Then
                MsgBox "No weights found", vbExclamation
                Exit Sub
            End If
            chrt.Chart.SetSourceData Source:=ws.Range("K2:L5")
            
        Case 5
            Set ws = ThisWorkbook.Sheets("NumberOfCriteria-5")
            If WorksheetIsEmpty(ws.Range("L2:L6")) Then
                MsgBox "No weights found", vbExclamation
                Exit Sub
            End If
            chrt.Chart.SetSourceData Source:=ws.Range("K2:L6")
            
        Case Else
            MsgBox "Error. Please check your input.", vbCritical
            Exit Sub
    End Select

    ' Customize the chart
    With chrt.Chart
        .ChartType = xl3DPie
        .ChartArea.Interior.ColorIndex = 40
        .SetElement msoElementDataLabelInsideEnd
        
        With .SeriesCollection(1).DataLabels
            .NumberFormat = "0.00%"
        End With

        With .ChartArea.Format.TextFrame2.TextRange.Font
            .Name = "Times New Roman"
            .Bold = True
            .Size = 12
        End With

        .SetElement msoElementChartTitleAboveChart
        .ChartTitle.Text = "Weights of the Criteria"
        .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 14
    End With

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

