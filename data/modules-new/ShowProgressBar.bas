Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim i As Long
    For i = 0 To 100 Step 10
        Application.StatusBar = "Ilerleme: %" & i
        DoEvents
    Next i
    Application.StatusBar = False
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "Tamam" : ws.Range("B1").Value = "100%"
    Set DynamicFunc = Nothing
End Function
