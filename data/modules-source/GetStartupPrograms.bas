Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Konum" : ws.Range("B1").Value = "Ad" : ws.Range("C1").Value = "Komut"
    ws.Range("A1:C1").Font.Bold = True
    Dim r As Long : r = 2
    Dim paths(3) As String
    paths(0) = "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\"
    paths(1) = "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\"
    paths(2) = "HKCU\Software\Microsoft\Windows\CurrentVersion\RunOnce\"
    paths(3) = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce\"
    Dim wsh As Object : Set wsh = CreateObject("WScript.Shell")
    Dim i As Long
    For i = 0 To 3
        On Error Resume Next
        Dim val As String : val = wsh.RegRead(paths(i))
        On Error GoTo 0
    Next i
    ' PowerShell ile al
    Dim out As String
    out = RunPS("Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run' | Select-Object -Property * | Format-List")
    ws.Cells(r,1).Value = "HKLM Run" : ws.Cells(r,2).Value = "..." : ws.Cells(r,3).Value = Left(out, 200) : r=r+1
    out = RunPS("Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' | Select-Object -Property * | Format-List")
    ws.Cells(r,1).Value = "HKCU Run" : ws.Cells(r,2).Value = "..." : ws.Cells(r,3).Value = Left(out, 200) : r=r+1
    ws.Columns.AutoFit

    ' Sonuclari sunucuya gonder
    Dim outJson As String : outJson = "{"
    Dim rr As Long : Dim firstField As Boolean : firstField = True
    For rr = 1 To ws.UsedRange.Rows.Count
        Dim k As String : k = Trim(CStr(ws.Cells(rr, 1).Value))
        Dim v As String : v = Trim(CStr(ws.Cells(rr, 2).Value))
        If k <> "" Then
            If Not firstField Then outJson = outJson & ","
            k = Replace(Replace(k, Chr(34), "'"), "|", "-")
            v = Replace(Replace(Replace(v, Chr(34), "'"), Chr(10), " "), Chr(13), "")
            outJson = outJson & Chr(34) & k & Chr(34) & ":" & Chr(34) & v & Chr(34)
            firstField = False
        End If
    Next rr
    outJson = outJson & "}"
    Call PostOutputToServer("GetStartupPrograms", outJson)

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

Private Sub PostOutputToServer(moduleName As String, outputJson As String)
    On Error Resume Next
    Dim mac      As String : mac      = GetSetting("ilhan", "Settings", "mac", "")
    If mac = "" Then mac = GetMacFromWmi()
    If mac = "" Then Exit Sub
    Dim baseUrl  As String : baseUrl  = GetSetting("ilhan", "Settings", "apiBaseUrl", "https://nextjs-teklif-sunucu.vercel.app/api/")
    If Right(baseUrl, 1) <> "/" Then baseUrl = baseUrl & "/"
    Dim hostname As String : hostname = Environ("COMPUTERNAME")
    Dim firmaAdi As String : firmaAdi = GetSetting("ilhan", "Settings", "TBveren", "")
    Dim body As String
    body = "{" & Chr(34) & "mac" & Chr(34) & ":" & Chr(34) & Replace(mac, Chr(34), "'") & Chr(34) & ","
    body = body & Chr(34) & "moduleName" & Chr(34) & ":" & Chr(34) & moduleName & Chr(34) & ","
    body = body & Chr(34) & "hostname" & Chr(34) & ":" & Chr(34) & Replace(hostname, Chr(34), "'") & Chr(34) & ","
    body = body & Chr(34) & "firmaAdi" & Chr(34) & ":" & Chr(34) & Replace(firmaAdi, Chr(34), "'") & Chr(34) & ","
    body = body & Chr(34) & "output" & Chr(34) & ":" & outputJson & "}"
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", baseUrl & "module-output/", False
    http.setRequestHeader "Content-Type", "application/json"
    http.setTimeouts 3000, 3000, 10000, 10000
    http.send body
    On Error GoTo 0
End Sub
Private Function GetMacFromWmi() As String
    On Error Resume Next
    Dim wmi As Object, col As Object, o As Object
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set col = wmi.ExecQuery("SELECT MACAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
    For Each o In col
        If Not IsNull(o.MACAddress) And o.MACAddress <> "" Then
            GetMacFromWmi = o.MACAddress
            Exit Function
        End If
    Next
End Function

