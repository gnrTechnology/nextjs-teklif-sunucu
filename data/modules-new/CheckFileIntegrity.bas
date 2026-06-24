Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim fp As String, expected As String
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    fp = Trim$(parts(0)) : expected = IIf(UBound(parts) >= 1, LCase$(Trim$(parts(1))), "")
    Application.Run "GetFileHashMd5", targetWb, fp
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    Dim actual As String : actual = LCase$(Trim$(CStr(ws.Range("B2").Value)))
    ws.Range("A4").Value = "Beklenen" : ws.Range("B4").Value = expected
    ws.Range("A5").Value = "Eslesme" : ws.Range("B5").Value = (actual = expected)
    Set DynamicFunc = Nothing
End Function
