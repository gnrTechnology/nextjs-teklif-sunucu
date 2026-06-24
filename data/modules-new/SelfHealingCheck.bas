Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Application.Run "zInternet.RunRemoteCodeQuiet", "AutoUpdateModules"
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "Kontrol" : ws.Range("B1").Value = "AutoUpdateModules tetiklendi"
    Set DynamicFunc = Nothing
End Function
