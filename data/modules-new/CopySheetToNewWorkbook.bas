Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim newPath As String : newPath = Trim$(CStr(param))
    If Len(newPath) = 0 Then newPath = Environ("TEMP") & "\sheet-copy.xlsx"
    targetWb.ActiveSheet.Copy
    ActiveWorkbook.SaveAs Filename:=newPath
    ActiveWorkbook.Close False
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Dosya" : ws.Range("B1").Value = newPath
    Set DynamicFunc = Nothing
End Function
