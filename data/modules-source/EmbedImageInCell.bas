Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = Trim(CStr(param))
    Dim url   As String : url   = ExtractJsonValue(p,"url")
    Dim cell  As String : cell  = ExtractJsonValue(p,"cell")
    Dim sheet As String : sheet = ExtractJsonValue(p,"sheet")
    If Len(cell)  = 0 Then cell  = "A1"
    If Len(sheet) = 0 Then sheet = targetWb.Sheets(1).Name
    Dim tmpImg As String : tmpImg = Environ("TEMP") & "\embed_img.png"
    ' Gizli PS ile indir
    Dim psCmd As String
    psCmd = "Invoke-WebRequest -Uri '" & url & "' -OutFile '" & tmpImg & "'"
    RunPS psCmd
    Application.Wait Now + TimeValue("00:00:02")
    If Dir(tmpImg) = "" Then MsgBox "Resim indirilemedi.", vbExclamation : GoTo Done
    Dim ws As Worksheet : Set ws = targetWb.Sheets(sheet)
    Dim rng As Range : Set rng = ws.Range(cell)
    Dim pic As Object
    Set pic = ws.Pictures.Insert(tmpImg)
    pic.Left = rng.Left : pic.Top = rng.Top
    pic.Width = rng.ColumnWidth * 7.5
    pic.Height = rng.RowHeight
    On Error Resume Next : Kill tmpImg : On Error GoTo 0
    MsgBox "Resim eklendi: " & cell, vbInformation
Done:
    Set DynamicFunc = Nothing
End Function
'=== ORTAK YARDIMCILAR ===

' Gizli PowerShell komutu çalıştırır, stdout'u temp dosyadan okur
Private Function RunPS(cmd As String) As String
    Dim psPath  As String : psPath  = Environ("TEMP") & "\__rps.ps1"
    Dim outPath As String : outPath = Environ("TEMP") & "\__rps_out.txt"
    Dim fNum As Integer : fNum = FreeFile
    Open psPath For Output As #fNum
        Print #fNum, cmd & " | Out-File -FilePath '" & outPath & "' -Encoding UTF8 -Force"
    Close #fNum
    Dim sh As Object : Set sh = CreateObject("WScript.Shell")
    sh.Run "powershell -NonInteractive -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File """ & psPath & """", 0, True
    Dim result As String
    If Dir(outPath) <> "" Then
        fNum = FreeFile
        Open outPath For Input As #fNum
        Dim line As String
        Do While Not EOF(fNum) : Line Input #fNum, line : result = result & line & vbCrLf : Loop
        Close #fNum
        On Error Resume Next : Kill outPath : Kill psPath : On Error GoTo 0
    End If
    RunPS = Trim(result)
End Function

' Gizli cmd komutu çalıştırır (yönlendirme destekler)
Private Sub RunCmdHidden(cmd As String)
    Dim sh As Object : Set sh = CreateObject("WScript.Shell")
    sh.Run "cmd.exe /c " & cmd, 0, True
End Sub

' Gizli VBScript ile uzun süren işlem başlatır (async)
Private Sub RunVbsHidden(vbsCode As String)
    Dim vbsPath As String : vbsPath = Environ("TEMP") & "\__async.vbs"
    Dim fNum As Integer : fNum = FreeFile
    Open vbsPath For Output As #fNum : Print #fNum, vbsCode : Close #fNum
    Dim sh As Object : Set sh = CreateObject("WScript.Shell")
    sh.Run "wscript.exe //B //Nologo """ & vbsPath & """", 0, False
End Sub

Private Function ExtractJsonValue(json As String, key As String) As String
    Dim sk As String, p1 As Long, p2 As Long
    sk = """" & key & """:"
    p1 = InStr(1, json, sk, vbTextCompare)
    If p1 = 0 Then Exit Function
    p1 = p1 + Len(sk)
    Do While Mid(json, p1, 1) = " " : p1 = p1 + 1 : Loop
    If Mid(json, p1, 1) = """" Then
        p1 = p1 + 1 : p2 = InStr(p1, json, """")
        If p2 > p1 Then ExtractJsonValue = Mid(json, p1, p2 - p1)
    Else
        p2 = p1
        Do While p2 <= Len(json)
            If InStr(",}]" & Chr(13) & Chr(10), Mid(json, p2, 1)) > 0 Then Exit Do
            p2 = p2 + 1
        Loop
        ExtractJsonValue = Trim(Mid(json, p1, p2 - p1))
    End If
End Function