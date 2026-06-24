Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim inPath As String : inPath = Trim$(CStr(param))
    If Len(inPath) = 0 Then inPath = Environ("LOCALAPPDATA") & "\TeklifAgent\vba-settings-backup.json"
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(inPath) Then GoTo Done
    Dim ts As Object : Set ts = fso.OpenTextFile(inPath, 1)
    Dim json As String : json = ts.ReadAll : ts.Close
    Dim n As Long : n = 0
    Dim pos As Long : pos = 1
    Do
        pos = InStr(pos, json, Chr(34))
        If pos = 0 Then Exit Do
        Dim kEnd As Long : kEnd = InStr(pos + 1, json, Chr(34))
        If kEnd = 0 Then Exit Do
        Dim key As String : key = Mid$(json, pos + 1, kEnd - pos - 1)
        Dim vStart As Long : vStart = InStr(kEnd, json, Chr(34))
        If vStart = 0 Then Exit Do
        vStart = vStart + 1
        Dim vEnd As Long : vEnd = InStr(vStart, json, Chr(34))
        If vEnd = 0 Then Exit Do
        Dim val As String : val = Mid$(json, vStart, vEnd - vStart)
        Dim dot As Long : dot = InStr(key, ".")
        If dot > 0 Then
            SaveSetting Left$(key, dot - 1), "Settings", Mid$(key, dot + 1), val
            n = n + 1
        End If
        pos = vEnd + 1
    Loop
Done:
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Geri yuklenen" : ws.Range("B1").Value = n
    ws.Range("A2").Value = "Kaynak" : ws.Range("B2").Value = inPath
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
