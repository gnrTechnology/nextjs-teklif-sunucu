Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim rng As Range : Set rng = targetWb.ActiveSheet.Range(CStr(param))
    Dim c As Range
    For Each c In rng.Cells
        c.Value = EncryptXor(CStr(c.Value), "teklif")
    Next c
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "Sifrelendi" : ws.Range("B1").Value = rng.Address
    Set DynamicFunc = Nothing
End Function
Private Function EncryptXor(s As String, k As String) As String
    Dim i As Long, o As String
    For i = 1 To Len(s) : o = o & Chr(Asc(Mid$(s, i, 1)) Xor Asc(Mid$(k, ((i - 1) Mod Len(k)) + 1, 1))) : Next
    EncryptXor = o
End Function
