Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Güncelleme" : ws.Range("B1").Value = "KB" : ws.Range("C1").Value = "Boyut"
    ws.Range("A1:C1").Font.Bold = True
    ' Get-WindowsUpdate gerektirmeden temel kontrol
    Dim out As String
    out = RunPS("(New-Object -ComObject Microsoft.Update.Searcher).Search('IsInstalled=0 and Type=Software').Updates | Select-Object Title,@{N='KB';E={($_.KBArticleIDs -join ',')}},@{N='Size';E={[math]::Round($_.MaxDownloadSize/1MB,1)}} | ConvertTo-Csv -NoTypeInformation")
    Dim lines() As String : lines = Split(out, Chr(10))
    Dim r As Long : r = 2
    Dim i As Long
    For i = 1 To UBound(lines)
        Dim ln As String : ln = Trim(Replace(lines(i), Chr(13),""))
        If Len(ln) > 2 Then
            ln = Replace(ln, Chr(34), "")
            Dim parts() As String : parts = Split(ln, ",")
            If UBound(parts) >= 2 Then
                ws.Cells(r,1).Value = parts(0) : ws.Cells(r,2).Value = parts(1) : ws.Cells(r,3).Value = parts(2) & " MB"
                r = r + 1
            End If
        End If
    Next i
    If r = 2 Then ws.Range("A2").Value = "✅ Bekleyen güncelleme yok veya yetki gerekiyor"
    ws.Columns.AutoFit
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
