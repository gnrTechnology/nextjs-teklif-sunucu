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
        Set colNet = objWMINet.ExecQuery("SELECT MACAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
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

    Dim result As String
    result = RunCaptureScript(apiBase, mac, hostname, firmaAdi)

    If InStr(1, result, "OK", vbTextCompare) > 0 Then
        MsgBox "Ekran goruntusu sunucuya gonderildi." & Chr(10) & _
               "Web: Modul Ciktilari > Ekran Goruntuleri", vbInformation, "CaptureScreenshot"
    Else
        MsgBox "Ekran goruntusu gonderilemedi." & Chr(10) & Left(result, 300), vbExclamation, "CaptureScreenshot"
    End If

    Set DynamicFunc = Nothing
End Function

' PowerShell betigini satir satir dosyaya yazar (VBA 25 satir devam limiti yok)
Private Function RunCaptureScript(apiBase As String, mac As String, hostname As String, firmaAdi As String) As String
    Dim psPath As String
    Dim outPath As String
    Dim fNum As Integer
    psPath = Environ("TEMP") & "\__capture_screenshot.ps1"
    outPath = Environ("TEMP") & "\__capture_screenshot_out.txt"

    On Error Resume Next
    Kill outPath
    On Error GoTo 0

    fNum = FreeFile
    Open psPath For Output As #fNum
    Print #fNum, "$ErrorActionPreference = 'Stop'"
    Print #fNum, "$api = '" & PsQuote(apiBase) & "'"
    Print #fNum, "$mac = '" & PsQuote(mac) & "'"
    Print #fNum, "$hn = '" & PsQuote(hostname) & "'"
    Print #fNum, "$firma = '" & PsQuote(firmaAdi) & "'"
    Print #fNum, "$outFile = '" & PsQuote(outPath) & "'"
    Print #fNum, "try {"
    Print #fNum, "  Add-Type -AssemblyName System.Windows.Forms"
    Print #fNum, "  Add-Type -AssemblyName System.Drawing"
    Print #fNum, "  $scr = [Windows.Forms.Screen]::PrimaryScreen"
    Print #fNum, "  $bmp = New-Object Drawing.Bitmap $scr.Bounds.Width, $scr.Bounds.Height"
    Print #fNum, "  $g = [Drawing.Graphics]::FromImage($bmp)"
    Print #fNum, "  $g.CopyFromScreen($scr.Bounds.Location, [Drawing.Point]::Empty, $bmp.Size)"
    Print #fNum, "  $g.Dispose()"
    Print #fNum, "  $maxW = 1280"
    Print #fNum, "  if ($bmp.Width -gt $maxW) {"
    Print #fNum, "    $ratio = $maxW / $bmp.Width"
    Print #fNum, "    $nh = [int]($bmp.Height * $ratio)"
    Print #fNum, "    $resized = New-Object Drawing.Bitmap $maxW, $nh"
    Print #fNum, "    $gr = [Drawing.Graphics]::FromImage($resized)"
    Print #fNum, "    $gr.InterpolationMode = [Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic"
    Print #fNum, "    $gr.DrawImage($bmp, 0, 0, $maxW, $nh)"
    Print #fNum, "    $gr.Dispose()"
    Print #fNum, "    $bmp.Dispose()"
    Print #fNum, "    $bmp = $resized"
    Print #fNum, "  }"
    Print #fNum, "  $ms = New-Object IO.MemoryStream"
    Print #fNum, "  $enc = [Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' }"
    Print #fNum, "  $ep = New-Object Drawing.Imaging.EncoderParameters 1"
    Print #fNum, "  $ep.Param[0] = New-Object Drawing.Imaging.EncoderParameter([Drawing.Imaging.Encoder]::Quality, 65L)"
    Print #fNum, "  $bmp.Save($ms, $enc, [ref]$ep)"
    Print #fNum, "  $bmp.Dispose()"
    Print #fNum, "  $b64 = [Convert]::ToBase64String($ms.ToArray())"
    Print #fNum, "  $ms.Dispose()"
    Print #fNum, "  $body = @{"
    Print #fNum, "    mac = $mac"
    Print #fNum, "    moduleName = 'CaptureScreenshot'"
    Print #fNum, "    hostname = $hn"
    Print #fNum, "    firmaAdi = $firma"
    Print #fNum, "    output = @{"
    Print #fNum, "      type = 'screenshot'"
    Print #fNum, "      mimeType = 'image/jpeg'"
    Print #fNum, "      imageBase64 = $b64"
    Print #fNum, "    }"
    Print #fNum, "  } | ConvertTo-Json -Depth 5 -Compress"
    Print #fNum, "  Invoke-RestMethod -Uri ($api + 'module-output/') -Method POST -Body $body -ContentType 'application/json; charset=utf-8' | Out-Null"
    Print #fNum, "  'OK' | Out-File -FilePath $outFile -Encoding UTF8 -Force"
    Print #fNum, "} catch {"
    Print #fNum, "  $_.Exception.Message | Out-File -FilePath $outFile -Encoding UTF8 -Force"
    Print #fNum, "}"
    Close #fNum

    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    sh.Run "powershell.exe -NonInteractive -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File """ & psPath & """", 0, True

    Dim result As String
    result = ""
    If Dir(outPath) <> "" Then
        fNum = FreeFile
        Open outPath For Input As #fNum
        Dim line As String
        Do While Not EOF(fNum)
            Line Input #fNum, line
            result = result & line & vbCrLf
        Loop
        Close #fNum
    End If

    On Error Resume Next
    Kill outPath
    Kill psPath
    On Error GoTo 0

    RunCaptureScript = Trim(result)
End Function
