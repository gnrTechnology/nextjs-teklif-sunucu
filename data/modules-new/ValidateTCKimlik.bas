Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim tc As String : tc = Replace(Trim$(CStr(param)), " ", "")
    Dim ok As Boolean : ok = False
    If Len(tc) = 11 And IsNumeric(tc) Then
        If Left$(tc, 1) <> "0" Then
            Dim d(1 To 11) As Long, i As Long
            For i = 1 To 11 : d(i) = CLng(Mid$(tc, i, 1)) : Next
            Dim s1 As Long, s2 As Long
            For i = 1 To 9 Step 2 : s1 = s1 + d(i) : Next
            For i = 2 To 8 Step 2 : s2 = s2 + d(i) : Next
            ok = ((s1 * 7 - s2) Mod 10 = d(10)) And (((s1 + s2 + d(10)) Mod 10) = d(11))
        End If
    End If
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "TC" : ws.Range("B1").Value = tc
    ws.Range("A2").Value = "Gecerli" : ws.Range("B2").Value = ok
    Set DynamicFunc = Nothing
End Function
