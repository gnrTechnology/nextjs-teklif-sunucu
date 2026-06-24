Private Function EscJson(s As String) As String
    Dim i As Long, c As String, out As String, ac As Long
    out = ""
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        ac = AscW(c)
        Select Case ac
            Case 92:  out = out & Chr(92) & Chr(92)
            Case 34:  out = out & Chr(92) & Chr(34)
            Case 10:  out = out & Chr(92) & "n"
            Case 13:  ' atla
            Case 9:   out = out & Chr(92) & "t"
            Case Else: out = out & c
        End Select
    Next i
    EscJson = out
End Function

Private Function PsQuote(s As String) As String
    PsQuote = Replace(s, "'", "''")
End Function

Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim apiBase As String
    apiBase = GetSetting("ilhan", "Settings", "apiBaseUrl", "https://nextjs-teklif-sunucu.vercel.app/api/")
    If Len(Trim(CStr(param))) > 0 And InStr(LCase(CStr(param)), "http") > 0 Then
        apiBase = Trim(CStr(param))
    End If
    If Right(apiBase, 1) <> "/" Then apiBase = apiBase & "/"

    Dim mac As String
    Dim hostname As String
    Dim firmaAdi As String
    mac = GetSetting("ilhan", "Settings", "mac", "")
    hostname = Environ("COMPUTERNAME")
    firmaAdi = GetSetting("ilhan", "Settings", "mdip", "")

    If mac = "" Then
        On Error Resume Next
        Dim objWMINet As Object, colNet As Object, objNet As Object
        Set objWMINet = GetObject("winmgmts:\\.\root\cimv2")
        Set colNet = objWMINet.ExecQuery( _
            "SELECT MACAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
        For Each objNet In colNet
            If objNet.MACAddress <> "" Then mac = objNet.MACAddress : Exit For
        Next
        On Error GoTo 0
    End If

    If mac = "" Then
        MsgBox "MAC adresi bulunamadi.", vbExclamation, "CaptureScreenshot"
        Set DynamicFunc = Nothing
        Exit Function
    End If

    Dim psCmd As String
    psCmd = "$ErrorActionPreference='Stop';" & _
        "$api='" & PsQuote(apiBase) & "';" & _
        "$mac='" & PsQuote(mac) & "';" & _
        "$hn='" & PsQuote(hostname) & "';" & _
        "$firma='" & PsQuote(firmaAdi) & "';" & _
        "Add-Type -AssemblyName System.Windows.Forms;" & _
        "Add-Type -AssemblyName System.Drawing;" & _
        "$scr=[Windows.Forms.Screen]::PrimaryScreen;" & _
        "$bmp=New-Object Drawing.Bitmap $scr.Bounds.Width,$scr.Bounds.Height;" & _
        "$g=[Drawing.Graphics]::FromImage($bmp);" & _
        "$g.CopyFromScreen($scr.Bounds.Location,[Drawing.Point]::Empty,$bmp.Size);" & _
        "$g.Dispose();" & _
        "$maxW=1280;" & _
        "if($bmp.Width -gt $maxW){" & _
        "$ratio=$maxW/$bmp.Width;" & _
        "$nh=[int]($bmp.Height*$ratio);" & _
        "$resized=New-Object Drawing.Bitmap $maxW,$nh;" & _
        "$gr=[Drawing.Graphics]::FromImage($resized);" & _
        "$gr.InterpolationMode=[Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic;" & _
        "$gr.DrawImage($bmp,0,0,$maxW,$nh);" & _
        "$gr.Dispose();$bmp.Dispose();$bmp=$resized};" & _
        "$ms=New-Object IO.MemoryStream;" & _
        "$enc=[Drawing.Imaging.ImageCodecInfo]::GetImageEncoders()|?{$_.MimeType -eq 'image/jpeg'};" & _
        "$ep=New-Object Drawing.Imaging.EncoderParameters 1;" & _
        "$ep.Param[0]=New-Object Drawing.Imaging.EncoderParameter([Drawing.Imaging.Encoder]::Quality,65L);" & _
        "$bmp.Save($ms,$enc,[ref]$ep);" & _
        "$bmp.Dispose();" & _
        "$b64=[Convert]::ToBase64String($ms.ToArray());" & _
        "$ms.Dispose();" & _
        "$body=@{mac=$mac;moduleName='CaptureScreenshot';hostname=$hn;firmaAdi=$firma;" & _
        "output=@{type='screenshot';mimeType='image/jpeg';imageBase64=$b64}}|" & _
        "ConvertTo-Json -Depth 5 -Compress;" & _
        "$r=Invoke-RestMethod -Uri ($api+'module-output/') -Method POST -Body $body -ContentType 'application/json; charset=utf-8';" & _
        "Write-Output 'OK'"

    Dim result As String
    result = RunPS(psCmd)

    If InStr(1, result, "OK", vbTextCompare) > 0 Then
        MsgBox "Ekran goruntusu sunucuya gonderildi." & Chr(10) & _
               "Web: Modul Ciktilari > Ekran Goruntuleri", vbInformation, "CaptureScreenshot"
    Else
        MsgBox "Ekran goruntusu gonderilemedi." & Chr(10) & Left(result, 200), vbExclamation, "CaptureScreenshot"
    End If

    Set DynamicFunc = Nothing
End Function

'=== ORTAK YARDIMCILAR ===

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
