Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    Dim plain As String : plain = parts(0)
    Dim key As String : key = IIf(UBound(parts) >= 1, parts(1), "teklif")
    Dim i As Long, out As String
    For i = 1 To Len(plain)
        out = out & Chr(Asc(Mid$(plain, i, 1)) Xor Asc(Mid$(key, ((i - 1) Mod Len(key)) + 1, 1)))
    Next i
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Sifreli" : ws.Range("B1").Value = Base64Encode(out)
    Set DynamicFunc = Nothing
End Function
Private Function Base64Encode(s As String) As String
    Dim dm As Object : Set dm = CreateObject("MSXML2.DOMDocument")
    Dim el As Object : Set el = dm.createElement("b64")
    el.DataType = "bin.base64" : el.nodeTypedValue = StrConv(s, vbFromUnicode)
    Base64Encode = Replace(Replace(el.Text, vbCr, ""), vbLf, "")
End Function
