Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim parts() As String : parts = Split(CStr(param), "|")
    Dim sep As String : sep = IIf(UBound(parts) >= 0, parts(0), " ")
    Dim c1 As Long : c1 = IIf(UBound(parts) >= 1, CLng(Val(parts(1))), 1)
    Dim c2 As Long : c2 = IIf(UBound(parts) >= 2, CLng(Val(parts(2))), 2)
    Dim r As Long
    For r = 1 To targetWb.ActiveSheet.UsedRange.Rows.Count
        targetWb.ActiveSheet.Cells(r, c2 + 1).Value = targetWb.ActiveSheet.Cells(r, c1).Value & sep & targetWb.ActiveSheet.Cells(r, c2).Value
    Next r
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "Birlestirildi"
    Set DynamicFunc = Nothing
End Function
