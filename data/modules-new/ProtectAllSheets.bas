Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim pwd As String : pwd = CStr(param)
    Dim sh As Worksheet
    For Each sh In targetWb.Worksheets
        sh.Protect Password:=pwd
    Next sh
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Korunan" : ws.Range("B1").Value = targetWb.Worksheets.Count
    Set DynamicFunc = Nothing
End Function
