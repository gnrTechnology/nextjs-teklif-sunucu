Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim list As String : list = CStr(param)
    Dim m As Variant
    For Each m In Split(list, ";")
        If Len(Trim$(CStr(m))) > 0 Then
            On Error Resume Next
            Application.Run "zInternet.RunRemoteCodeQuiet", Trim$(CStr(m))
            Err.Clear
        End If
    Next m
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Zincir" : ws.Range("B1").Value = list
    Set DynamicFunc = Nothing
End Function
