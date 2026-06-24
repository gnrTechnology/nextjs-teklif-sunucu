Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    On Error GoTo Fail
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    Dim progId As String : progId = Trim$(parts(0))
    Dim method As String : method = IIf(UBound(parts) >= 1, Trim$(parts(1)), "")
    Dim obj As Object : Set obj = CreateObject(progId)
    Dim result As Variant
    If Len(method) > 0 Then result = CallByName(obj, method, VbMethod)
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = progId : ws.Range("B1").Value = CStr(result)
    GoTo Done
Fail:
    ws.Range("A1").Value = "Hata" : ws.Range("B1").Value = Err.Description
Done:
    Set DynamicFunc = Nothing
End Function
