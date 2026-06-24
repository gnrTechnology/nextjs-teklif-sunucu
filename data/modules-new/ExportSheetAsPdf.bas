Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim outPdf As String : outPdf = Trim$(CStr(param))
    If Len(outPdf) = 0 Then outPdf = Environ("TEMP") & "\export.pdf"
    targetWb.ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=outPdf
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "PDF" : ws.Range("B1").Value = outPdf
    Set DynamicFunc = Nothing
End Function
