Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim outDir As String : outDir = Trim$(CStr(param))
    If Len(outDir) = 0 Then outDir = Environ("TEMP") & "\vba-export\"
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(outDir) Then fso.CreateFolder outDir
    Dim comp As Object, n As Long : n = 0
    For Each comp In targetWb.VBProject.VBComponents
        If comp.Type = 1 Then
            n = n + 1
            Dim ts As Object : Set ts = fso.CreateTextFile(outDir & comp.Name & ".bas", True)
            ts.Write comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines) : ts.Close
        End If
    Next comp
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "Modul" : ws.Range("B1").Value = n
    Set DynamicFunc = Nothing
End Function
