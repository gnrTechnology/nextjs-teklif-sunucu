Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ' param: {"taskName":"GorevAdi","command":"script.bat","startTime":"08:00"}
    Dim p As String : p = CStr(param)
    Dim taskName  As String : taskName  = ExtractJsonValue(p, "taskName")
    Dim command   As String : command   = ExtractJsonValue(p, "command")
    Dim startTime As String : startTime = ExtractJsonValue(p, "startTime")

    If Len(taskName)  = 0 Then taskName  = "VbaGunluk_" & Format(Now, "yyyymmdd")
    If Len(command)   = 0 Then command   = "notepad.exe"
    If Len(startTime) = 0 Then startTime = "08:00"

    Dim cmd As String
    cmd = "schtasks /create /f /tn """ & taskName & """ /sc DAILY " & _
          "/st """ & startTime & """ /tr """ & command & """"

    Debug.Print "[ScheduleTaskDaily] " & cmd

    Dim wsh As Object : Set wsh = CreateObject("WScript.Shell")
    Dim ret As Long : ret = wsh.Run("cmd.exe /c " & cmd, 0, True)

    If ret = 0 Then
        MsgBox "Günlük görev oluşturuldu: " & taskName & vbCrLf & "Saat: " & startTime, vbInformation
    Else
        MsgBox "Görev oluşturulamadı (hata kodu: " & ret & ")", vbExclamation
    End If
    Set DynamicFunc = Nothing
End Function

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
            If InStr(",}] " & Chr(13) & Chr(10), Mid(json, p2, 1)) > 0 Then Exit Do
            p2 = p2 + 1
        Loop
        ExtractJsonValue = Trim(Mid(json, p1, p2 - p1))
    End If
End Function
