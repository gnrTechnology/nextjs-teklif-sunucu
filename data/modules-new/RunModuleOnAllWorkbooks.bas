Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim wb As Workbook, n As Long : n = 0
    For Each wb In Application.Workbooks
        If Not wb.IsAddin Then
            On Error Resume Next
            Application.Run "zInternet.RunRemoteCodeQuiet", CStr(param), wb
            n = n + 1
            Err.Clear
        End If
    Next wb
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "Workbook" : ws.Range("B1").Value = n
    Set DynamicFunc = Nothing
End Function
