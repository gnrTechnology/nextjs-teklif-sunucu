Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ' param: görev adı (string) veya {"taskName":"GorevAdi"}
    Dim taskName As String
    Dim p As String : p = Trim(CStr(param))

    If Left(p, 1) = "{" Then
        taskName = ExtractJsonValue(p, "taskName")
    Else
        taskName = p
    End If

    If Len(taskName) = 0 Then
        MsgBox "Görev adı belirtilmedi.", vbExclamation
        Set DynamicFunc = Nothing
        Exit Function
    End If

    Dim cmd As String
    cmd = "schtasks /delete /f /tn """ & taskName & """"

    Dim wsh As Object : Set wsh = CreateObject("WScript.Shell")
    Dim ret As Long : ret = wsh.Run("cmd.exe /c " & cmd, 0, True)

    If ret = 0 Then
        MsgBox "Görev silindi: " & taskName, vbInformation
    Else
        MsgBox "Görev silinemedi (hata: " & ret & "). Var mı kontrol edin.", vbExclamation
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
    End If
End Function
