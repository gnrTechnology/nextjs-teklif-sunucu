Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim idx As Long : idx = CLng(Val(CStr(param)))
    If idx = 0 Then idx = 0
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Metric" : ws.Range("B1").Value = GetSystemMetrics(idx)
    Set DynamicFunc = Nothing
End Function
