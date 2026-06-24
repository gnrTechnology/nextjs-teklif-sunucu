Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    Dim b64 As String : b64 = parts(0)
    Dim key As String : key = IIf(UBound(parts) >= 1, parts(1), "teklif")
    Dim bin As String : bin = Base64Decode(b64)
    Dim i As Long, out As String
    For i = 1 To Len(bin)
        out = out & Chr(Asc(Mid$(bin, i, 1)) Xor Asc(Mid$(key, ((i - 1) Mod Len(key)) + 1, 1)))
    Next i
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Cozuldu" : ws.Range("B1").Value = out
    Set DynamicFunc = Nothing
End Function
Private Function Base64Decode(s As String) As String
    Dim dm As Object : Set dm = CreateObject("MSXML2.DOMDocument")
    Dim el As Object : Set el = dm.createElement("b64")
    el.DataType = "bin.base64" : el.Text = s
    Base64Decode = StrConv(el.nodeTypedValue, vbUnicode)
End Function
