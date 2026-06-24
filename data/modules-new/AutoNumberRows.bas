Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim col As Long : col = CLng(Val(CStr(param))) : If col < 1 Then col = 1
    Dim ur As Range : Set ur = targetWb.ActiveSheet.UsedRange
    Dim r As Long
    For r = 1 To ur.Rows.Count
        targetWb.ActiveSheet.Cells(ur.Row + r - 1, col).Value = r
    Next r
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Satir" : ws.Range("B1").Value = ur.Rows.Count
    Set DynamicFunc = Nothing
End Function
