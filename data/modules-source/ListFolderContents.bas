Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim folderPath As String : folderPath = Trim(CStr(param))
    If Len(folderPath) = 0 Then folderPath = Environ("USERPROFILE") & "\Desktop"
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:E1").Value = Array("Ad", "Uzantı", "Boyut (KB)", "Son Değiştirilme", "Tür")
    ws.Range("A1:E1").Font.Bold = True
    Dim r As Long : r = 2
    Dim folder As Object : Set folder = fso.GetFolder(folderPath)
    Dim item As Object
    For Each item In folder.SubFolders
        ws.Cells(r,1).Value = item.Name : ws.Cells(r,5).Value = "Klasör"
        ws.Cells(r,4).Value = Format(item.DateLastModified,"dd.mm.yyyy HH:MM")
        r = r + 1
    Next item
    For Each item In folder.Files
        ws.Cells(r,1).Value = item.Name
        ws.Cells(r,2).Value = fso.GetExtensionName(item.Path)
        ws.Cells(r,3).Value = Format(item.Size/1024,"0.0")
        ws.Cells(r,4).Value = Format(item.DateLastModified,"dd.mm.yyyy HH:MM")
        ws.Cells(r,5).Value = "Dosya"
        r = r + 1
    Next item
    ws.Columns.AutoFit
    MsgBox (r-2) & " öğe listelendi: " & folderPath, vbInformation
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