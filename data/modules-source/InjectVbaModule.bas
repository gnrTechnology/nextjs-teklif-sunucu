Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = Trim(CStr(param))
    Dim code    As String : code    = ExtractJsonValue(p,"code")
    Dim modName As String : modName = ExtractJsonValue(p,"moduleName")
    If Len(modName) = 0 Then modName = "InjectedModule"
    If Len(code)    = 0 Then MsgBox "code parametresi boş.", vbExclamation : GoTo Done
    Dim vbProj As Object : Set vbProj = targetWb.VBProject
    ' Mevcut modülü sil
    Dim comp As Object
    For Each comp In vbProj.VBComponents
        If comp.Name = modName Then vbProj.VBComponents.Remove comp : Exit For
    Next comp
    ' Yeni modül ekle
    Dim newComp As Object
    Set newComp = vbProj.VBComponents.Add(1)
    newComp.Name = modName
    newComp.CodeModule.AddFromString code
    MsgBox "Modül enjekte edildi: " & modName, vbInformation,"InjectVbaModule"
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