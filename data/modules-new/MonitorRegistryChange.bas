Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim regPath As String : regPath = Trim$(CStr(param))
    If Len(regPath) = 0 Then regPath = "HKCU\Software\ilhan\Settings"
    Dim snapFile As String : snapFile = Environ("LOCALAPPDATA") & "\TeklifAgent\reg-snap.txt"
    Dim sh As Object : Set sh = CreateObject("WScript.Shell")
    Dim cur As String : cur = ""
    On Error Resume Next
    cur = CStr(sh.RegRead(regPath))
    On Error GoTo 0
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    Dim old As String : old = ""
    If fso.FileExists(snapFile) Then
        Dim ts As Object : Set ts = fso.OpenTextFile(snapFile, 1)
        old = ts.ReadAll : ts.Close
    End If
    Dim changed As Boolean : changed = (old <> cur)
    If changed Or Len(old) = 0 Then
        Dim tw As Object : Set tw = fso.OpenTextFile(snapFile, 2, True)
        tw.Write cur : tw.Close
    End If
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Yol" : ws.Range("B1").Value = regPath
    ws.Range("A2").Value = "Degisti" : ws.Range("B2").Value = IIf(changed And Len(old) > 0, "EVET", "HAYIR")
    ws.Range("A3").Value = "Deger" : ws.Range("B3").Value = Left$(cur, 32000)
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
