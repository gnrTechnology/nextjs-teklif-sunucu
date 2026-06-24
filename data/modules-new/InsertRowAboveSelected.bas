Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim n As Long : n = CLng(Val(CStr(param))) : If n < 1 Then n = 1
    Dim r As Long
    For r = 1 To n
        targetWb.ActiveSheet.Rows(targetWb.ActiveSheet.Selection.Row).Insert
    Next r
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Eklenen" : ws.Range("B1").Value = n
    Set DynamicFunc = Nothing
End Function
