Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim col As Long : col = CLng(Val(CStr(param))) : If col < 1 Then col = 1
    targetWb.ActiveSheet.Columns(col).NumberFormat = "#,##0.00 ""₺"""
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Sutun" : ws.Range("B1").Value = col
    Set DynamicFunc = Nothing
End Function
