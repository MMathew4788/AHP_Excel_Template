Attribute VB_Name = "close_chart"
Sub close_chart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    
    ' Reference the "Home" sheet
    Set ws = Sheets("Home")
    
    ' Loop through all chart objects in the sheet and delete them
    For Each chartObj In ws.ChartObjects
        chartObj.Delete
    Next chartObj
End Sub

