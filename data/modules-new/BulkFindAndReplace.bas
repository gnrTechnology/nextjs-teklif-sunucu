Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim oldT As String, newT As String
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    oldT = parts(0) : newT = IIf(UBound(parts) >= 1, parts(1), "")
    Dim sh As Worksheet, c As Range, n As Long : n = 0
    For Each sh In targetWb.Worksheets
        For Each c In sh.UsedRange
            If InStr(1, CStr(c.Value), oldT, vbTextCompare) > 0 Then
                c.Value = Replace(CStr(c.Value), oldT, newT, , , vbTextCompare) : n = n + 1
            End If
        Next c
    Next sh
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Degisen hucre" : ws.Range("B1").Value = n
    Set DynamicFunc = Nothing
End Function
