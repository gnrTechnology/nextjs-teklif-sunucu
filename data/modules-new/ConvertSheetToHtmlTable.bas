Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim outPath As String : outPath = Trim$(CStr(param))
    If Len(outPath) = 0 Then outPath = Environ("TEMP") & "\sheet.html"
    targetWb.ActiveSheet.SaveAs outPath, xlHtml
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "HTML" : ws.Range("B1").Value = outPath
    Set DynamicFunc = Nothing
End Function
