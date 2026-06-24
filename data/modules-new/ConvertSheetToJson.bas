Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim ur As Range : Set ur = targetWb.ActiveSheet.UsedRange
    Dim json As String : json = "["
    Dim r As Long, c As Long
    For r = 1 To ur.Rows.Count
        If r > 1 Then json = json & ","
        json = json & "{"
        For c = 1 To ur.Columns.Count
            If c > 1 Then json = json & ","
            json = json & Chr(34) & "c" & c & Chr(34) & ":" & Chr(34) & Replace(CStr(ur.Cells(r, c).Value), Chr(34), "'") & Chr(34)
        Next c
        json = json & "}"
    Next r
    json = json & "]"
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "JSON" : ws.Range("B1").Value = Left$(json, 32000)
    Set DynamicFunc = Nothing
End Function
