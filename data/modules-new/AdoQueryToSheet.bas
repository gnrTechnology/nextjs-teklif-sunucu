Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    On Error GoTo Fail
    Dim cn As Object : Set cn = CreateObject("ADODB.Connection")
    cn.Open Trim$(parts(0))
    Dim rs As Object : Set rs = cn.Execute(IIf(UBound(parts) >= 1, parts(1), "SELECT 1"))
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Cells.ClearContents
    Dim r As Long, c As Long : r = 1
    For c = 0 To rs.Fields.Count - 1 : ws.Cells(1, c + 1).Value = rs.Fields(c).Name : Next
    Do While Not rs.EOF
        r = r + 1
        For c = 0 To rs.Fields.Count - 1 : ws.Cells(r, c + 1).Value = rs.Fields(c).Value : Next
        rs.MoveNext
    Loop
    GoTo Done
Fail: ws.Range("A1").Value = "Hata" : ws.Range("B1").Value = Err.Description
Done: Set DynamicFunc = Nothing
End Function
