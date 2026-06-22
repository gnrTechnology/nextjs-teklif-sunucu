Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ' param: {"steps":5,"message":"İşleniyor..."}
    Dim p As String : p = CStr(param)
    Dim steps   As Long   : steps   = 5
    Dim msg     As String : msg     = "İşleniyor"
    Dim stepsS  As String : stepsS  = ExtractJsonVal(p,"steps")
    Dim msgS    As String : msgS    = ExtractJsonVal(p,"message")
    If IsNumeric(stepsS) Then steps = CLng(stepsS)
    If msgS <> "" Then msg = msgS
    Dim i As Long
    For i = 1 To steps
        Dim pct As Long : pct = CLng(i / steps * 100)
        Application.StatusBar = msg & " [" & String(pct\5, "=") & String(20-(pct\5), "-") & "] " & pct & "%"
        Application.Wait Now + TimeValue("00:00:01")
    Next i
    Application.StatusBar = False
    targetWb.Sheets(1).Range("A1").Value = "✅ Tamamlandı"
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
