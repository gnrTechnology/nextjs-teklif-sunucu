Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim col As Long : col = CLng(Val(CStr(param))) : If col < 1 Then col = 1
    Dim ur As Range : Set ur = targetWb.ActiveSheet.UsedRange
    Dim r As Long, n As Long : n = 0
    For r = 1 To ur.Rows.Count
        Dim v As String : v = Replace(Replace(Replace(CStr(ur.Cells(r, col).Value), " ", ""), "-", ""), "(", "")
        v = Replace(v, ")", "")
        If Len(v) >= 10 Then
            If Left$(v, 1) = "0" Then v = Mid$(v, 2)
            If Left$(v, 2) <> "90" Then v = "90" & v
            ur.Cells(r, col).Value = "+90 " & Mid$(v, 3, 3) & " " & Mid$(v, 6, 3) & " " & Mid$(v, 9, 2) & " " & Mid$(v, 11, 2)
            n = n + 1
        End If
    Next r
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Normalize" : ws.Range("B1").Value = n
    Set DynamicFunc = Nothing
End Function
