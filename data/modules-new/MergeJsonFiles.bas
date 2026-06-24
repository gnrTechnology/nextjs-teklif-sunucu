Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim list As String : list = CStr(param)
    Dim merged As String : merged = "["
    Dim f As Variant, first As Boolean : first = True
    For Each f In Split(list, ";")
        Dim ts As Object
        Set ts = CreateObject("Scripting.FileSystemObject").OpenTextFile(Trim$(CStr(f)), 1)
        Dim txt As String : txt = Trim$(ts.ReadAll) : ts.Close
        txt = Mid$(txt, 2, Len(txt) - 2)
        If Len(txt) > 0 Then
            If Not first Then merged = merged & ","
            merged = merged & txt : first = False
        End If
    Next f
    merged = merged & "]"
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "JSON" : ws.Range("B1").Value = Left$(merged, 32000)
    Set DynamicFunc = Nothing
End Function
