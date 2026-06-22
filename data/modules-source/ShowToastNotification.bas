Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ' param: {"title":"Başlık","message":"Mesaj"}  veya düz metin
    Dim p As String : p = CStr(param)
    Dim title As String : title = ExtractJsonVal(p,"title")
    Dim msg   As String : msg   = ExtractJsonVal(p,"message")
    If title = "" Then title = "Excel Bildirimi"
    If msg   = "" Then msg   = p
    Dim out As String
    out = RunPS("$xml=[Windows.UI.Notifications.ToastNotificationManager,Windows.UI.Notifications,ContentType=WindowsRuntime];" & _
        "[void]$xml;" & _
        "$t=[Windows.UI.Notifications.ToastTemplateType,Windows.UI.Notifications,ContentType=WindowsRuntime]::ToastText02;" & _
        "$c=[Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent($t);" & _
        "$c.GetElementsByTagName('text')[0].AppendChild($c.CreateTextNode('" & Replace(title,"'","`'") & "'));" & _
        "$c.GetElementsByTagName('text')[1].AppendChild($c.CreateTextNode('" & Replace(msg,"'","`'") & "'));" & _
        "$n=[Windows.UI.Notifications.ToastNotification,Windows.UI.Notifications,ContentType=WindowsRuntime]::new($c);" & _
        "[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('Excel').Show($n);" & _
        "'OK'")
    If Trim(out) <> "OK" Then
        ' Fallback: WScript popup
        CreateObject("WScript.Shell").Popup msg, 3, title, 64
    End If
    Set DynamicFunc = Nothing
End Function

Private Function RunPS(cmd As String) As String
    Dim tmp As String : tmp = Environ("TEMP") & "\ps_out_" & CLng(Timer*1000) & ".txt"
    Dim wsh As Object : Set wsh = CreateObject("WScript.Shell")
    wsh.Run "powershell -NonInteractive -NoProfile -Command " & Chr(34) & cmd & " | Out-File -Encoding UTF8 '" & tmp & "'" & Chr(34), 0, True
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(tmp) Then
        Dim f As Object : Set f = fso.OpenTextFile(tmp, 1, False, -1)
        RunPS = f.ReadAll : f.Close : fso.DeleteFile tmp
    End If
End Function

Private Function EscJ(s As String) As String
    s = Replace(s, "\", "\\") : s = Replace(s, Chr(34), "\""")
    s = Replace(s, Chr(10), "\n") : s = Replace(s, Chr(13), "")
    EscJ = s
End Function

Private Function ExtractJsonVal(json As String, key As String) As String
    Dim sk As String : sk = Chr(34) & key & Chr(34) & ":"
    Dim p1 As Long : p1 = InStr(1, json, sk, vbTextCompare)
    If p1 = 0 Then Exit Function
    p1 = p1 + Len(sk)
    Do While Mid(json, p1, 1) = " " : p1 = p1 + 1 : Loop
    If Mid(json, p1, 1) = Chr(34) Then
        p1 = p1 + 1 : Dim p2 As Long : p2 = InStr(p1, json, Chr(34))
        If p2 > p1 Then ExtractJsonVal = Mid(json, p1, p2 - p1)
    Else
        Dim p3 As Long : p3 = p1
        Do While p3 <= Len(json)
            If InStr(",}] " & Chr(13) & Chr(10), Mid(json, p3, 1)) > 0 Then Exit Do
            p3 = p3 + 1
        Loop
        ExtractJsonVal = Trim(Mid(json, p1, p3 - p1))
    End If
End Function

Private Sub WriteResult(ws As Worksheet, key As String, val As String, row As Long)
    ws.Cells(row, 1).Value = key
    ws.Cells(row, 2).Value = val
End Sub
