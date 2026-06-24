Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim json As String : json = CStr(param)
    Dim pos As Long, oldN As String, newN As String, n As Long : n = 0
    pos = 1
    Do
        pos = InStr(pos, json, Chr(34))
        If pos = 0 Then Exit Do
        oldN = JsonPair(json, pos, newN)
        If Len(oldN) = 0 Then Exit Do
        On Error Resume Next
        targetWb.Worksheets(oldN).Name = newN
        If Err.Number = 0 Then n = n + 1
        Err.Clear
    Loop
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Yeniden ad" : ws.Range("B1").Value = n
    Set DynamicFunc = Nothing
End Function
Private Function JsonPair(j As String, start As Long, ByRef v As String) As String
    Dim p As Long : p = InStr(start, j, Chr(34)) : If p = 0 Then Exit Function
    p = p + 1 : Dim p2 As Long : p2 = InStr(p, j, Chr(34))
    JsonPair = Mid$(j, p, p2 - p)
    p = InStr(p2, j, ":") : p = InStr(p, j, Chr(34)) + 1
    p2 = InStr(p, j, Chr(34)) : v = Mid$(j, p, p2 - p)
End Function
