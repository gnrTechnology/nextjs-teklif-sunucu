Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim col As Long : col = CLng(Val(CStr(param)))
    If col < 1 Then col = 1
    Dim ur As Range : Set ur = targetWb.ActiveSheet.UsedRange
    Dim before As Long : before = ur.Rows.Count
    ur.RemoveDuplicates Columns:=col, Header:=xlYes
    Dim after As Long : after = targetWb.ActiveSheet.UsedRange.Rows.Count
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Silinen" : ws.Range("B1").Value = before - after
    Set DynamicFunc = Nothing
End Function
