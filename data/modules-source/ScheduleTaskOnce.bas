Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ' param: {"taskName":"GorevAdi","command":"notepad.exe","runAt":"2026-06-23T09:00:00"}
    Dim p As String : p = CStr(param)
    Dim taskName As String : taskName = ExtractJsonValue(p, "taskName")
    Dim command  As String : command  = ExtractJsonValue(p, "command")
    Dim runAt    As String : runAt    = ExtractJsonValue(p, "runAt")

    If Len(taskName) = 0 Then taskName = "VbaGorev_" & Format(Now, "yyyymmddHHMMSS")
    If Len(command)  = 0 Then command  = "notepad.exe"
    If Len(runAt)    = 0 Then runAt    = Format(Now + TimeSerial(0, 5, 0), "yyyy-MM-ddTHH:mm:ss")

    Dim runDate As String : runDate = Left(runAt, 10)
    Dim runTime As String : runTime = Mid(runAt, 12, 5)

    Dim cmd As String
    cmd = "schtasks /create /f /tn """ & taskName & """ /sc ONCE " & _
          "/sd """ & runDate & """ /st """ & runTime & """ " & _
          "/tr """ & command & """"

    Debug.Print "[ScheduleTaskOnce] " & cmd

    Dim wsh As Object : Set wsh = CreateObject("WScript.Shell")
    Dim ret As Long : ret = wsh.Run("cmd.exe /c " & cmd, 0, True)

    If ret = 0 Then
        MsgBox "Görev oluşturuldu: " & taskName & vbCrLf & "Çalışma: " & runAt, vbInformation
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
        p1 = p1 + 1
        p2 = InStr(p1, json, """")
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
