Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim c As Range, r As Long : r = 1
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Cells.ClearContents : ws.Range("A1").Value = "E-posta"
    For Each c In targetWb.ActiveSheet.UsedRange
        If InStr(c.Value, "@") > 0 And InStr(c.Value, ".") > 0 Then r = r + 1 : ws.Cells(r, 1).Value = c.Value
    Next c
    Set DynamicFunc = Nothing
End Function
