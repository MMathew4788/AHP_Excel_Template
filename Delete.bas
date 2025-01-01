Attribute VB_Name = "Delete"
Sub Delete()
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("Home")
    ws.Range("J4").ClearContents


    Set ws = ThisWorkbook.Sheets("NumberOfCriteria-3")
    ws.Range("A1:A4").ClearContents
    ws.Range("A1:D1").ClearContents
    ws.Range("A7:A9").ClearContents
    ws.Range("E7:E10").ClearContents


    Set ws = ThisWorkbook.Sheets("NumberOfCriteria-4")
    ws.Range("A1:E1").ClearContents
    ws.Range("A1:A5").ClearContents
    ws.Range("A8:A13").ClearContents
    ws.Range("E8:E13").ClearContents
    
    Set ws = ThisWorkbook.Sheets("NumberOfCriteria-5")
    ws.Range("A1:F1").ClearContents
    ws.Range("A1:A6").ClearContents
    ws.Range("A9:A18").ClearContents
    ws.Range("E9:E18").ClearContents
    
End Sub

