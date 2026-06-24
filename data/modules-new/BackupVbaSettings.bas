Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim outPath As String
    outPath = Trim$(CStr(param))
    If Len(outPath) = 0 Then outPath = Environ("LOCALAPPDATA") & "\TeklifAgent\vba-settings-backup.json"
    Dim sections() As String : sections = Split("ilhan,scngnr,sercan", ",")
    Dim keys() As String : keys = Split("mac,mdip,TBveren,teklifYolu,startingAddin,ihlalDosyaYolu,ihlalDosyaAdi,apiBaseUrl", ",")
    Dim json As String : json = "{"
    Dim si As Long, ki As Long, first As Boolean : first = True
    For si = LBound(sections) To UBound(sections)
        For ki = LBound(keys) To UBound(keys)
            If Not first Then json = json & ","
            json = json & Chr(34) & sections(si) & "." & keys(ki) & Chr(34) & ":"
            json = json & Chr(34) & JsonEsc(GetSetting(sections(si), "Settings", keys(ki), "")) & Chr(34)
            first = False
        Next ki
    Next si
    json = json & "}"
    WriteText outPath, json
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Yedek dosyasi" : ws.Range("B1").Value = outPath
    ws.Range("A2").Value = "Boyut" : ws.Range("B2").Value = Len(json)
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
Private Function JsonEsc(s As String) As String
    JsonEsc = Replace(Replace(CStr(s), Chr(92), Chr(92) & Chr(92)), Chr(34), Chr(92) & Chr(34))
End Function
Private Sub WriteText(p As String, t As String)
    Dim f As Integer : f = FreeFile
    Dim dir As String : dir = Left$(p, InStrRev(p, Chr(92)) - 1)
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(dir) Then fso.CreateFolder dir
    Open p For Output As #f : Print #f, t; : Close #f
End Sub
