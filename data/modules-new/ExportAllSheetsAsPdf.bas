Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim outDir As String : outDir = Trim$(CStr(param))
    If Len(outDir) = 0 Then outDir = Environ("TEMP") & "\pdf-export\"
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(outDir) Then fso.CreateFolder outDir
    Dim sh As Worksheet, n As Long : n = 0
    For Each sh In targetWb.Worksheets
        sh.ExportAsFixedFormat xlTypePDF, outDir & sh.Name & ".pdf"
        n = n + 1
    Next sh
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "PDF sayisi" : ws.Range("B1").Value = n
    ws.Range("A2").Value = "Klasor" : ws.Range("B2").Value = outDir
    Set DynamicFunc = Nothing
End Function
