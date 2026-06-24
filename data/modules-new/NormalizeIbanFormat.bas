Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim iban As String : iban = UCase$(Replace(Replace(CStr(param), " ", ""), "-", ""))
    Dim out As String : Dim i As Long
    For i = 1 To Len(iban) Step 4
        If Len(out) > 0 Then out = out & " "
        out = out & Mid$(iban, i, 4)
    Next i
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "IBAN" : ws.Range("B1").Value = out
    Set DynamicFunc = Nothing
End Function
