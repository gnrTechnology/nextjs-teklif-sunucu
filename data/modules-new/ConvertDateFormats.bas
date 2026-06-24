Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim col As Long : col = CLng(Val(CStr(param))) : If col < 1 Then col = 1
    Dim ur As Range : Set ur = targetWb.ActiveSheet.UsedRange
    Dim r As Long
    For r = 2 To ur.Rows.Count
        If IsDate(ur.Cells(r, col).Value) Then ur.Cells(r, col).Value = Format(CDate(ur.Cells(r, col).Value), "yyyy-mm-dd")
    Next r
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "Sutun" : ws.Range("B1").Value = col
    Set DynamicFunc = Nothing
End Function
