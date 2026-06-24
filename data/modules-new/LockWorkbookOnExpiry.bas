Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim expiry As Date : expiry = CDate(param)
    If Now > expiry Then
        Dim pwd As String : pwd = "LOCKED"
        Dim sh As Worksheet
        For Each sh In targetWb.Worksheets : sh.Protect pwd : Next sh
    End If
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Bitis" : ws.Range("B1").Value = expiry
    ws.Range("A2").Value = "Kilitli" : ws.Range("B2").Value = (Now > expiry)
    Set DynamicFunc = Nothing
End Function
