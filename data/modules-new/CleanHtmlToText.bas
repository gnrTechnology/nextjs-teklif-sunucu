Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim html As String : html = CStr(param)
    Dim i As Long, inside As Boolean, ch As String, out As String
    inside = False
    For i = 1 To Len(html)
        ch = Mid$(html, i, 1)
        If ch = "<" Then inside = True
        If Not inside Then out = out & ch
        If ch = ">" Then inside = False
    Next i
    out = Replace(Replace(out, "&nbsp;", " "), "&amp;", "&")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Metin" : ws.Range("B1").Value = Left$(Trim$(out), 32000)
    Set DynamicFunc = Nothing
End Function
