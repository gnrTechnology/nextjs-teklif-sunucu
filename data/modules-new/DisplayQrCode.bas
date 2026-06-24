Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim url As String : url = "https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=" & EncodeUrl(CStr(param))
    targetWb.ActiveSheet.Shapes.AddPicture url, False, True, 50, 50, 150, 150
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "QR" : ws.Range("B1").Value = CStr(param)
    Set DynamicFunc = Nothing
End Function
Private Function EncodeUrl(s As String) As String
    EncodeUrl = Replace(Replace(s, " ", "%20"), "#", "%23")
End Function
