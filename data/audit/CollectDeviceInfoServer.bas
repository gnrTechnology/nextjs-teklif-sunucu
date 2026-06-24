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
    mac      = GetSetting("ilhan", "Settings", "mac", "")
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

    If mac = "" Then Set DynamicFunc = Nothing : Exit Function

    Dim json As String
    json = "{"

    ' ── Temel sistem bilgileri ────────────────────────
    json = json & """computerName"":""" & EscJson(hostname) & ""","
    json = json & """loggedInUser"":""" & EscJson(Environ("USERNAME")) & ""","

    Dim wmi As Object
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Dim col As Object, obj As Object

    ' Windows sürümü
    Set col = wmi.ExecQuery("SELECT Caption, BuildNumber, Version FROM Win32_OperatingSystem")
    For Each obj In col
        json = json & """windowsVersion"":""" & EscJson(obj.Caption) & " (Build " & obj.BuildNumber & ")" & ""","
    Next

    ' CPU
    Set col = wmi.ExecQuery("SELECT Name, NumberOfCores, MaxClockSpeed FROM Win32_Processor")
    Dim cpuStr As String : cpuStr = ""
    For Each obj In col
        cpuStr = EscJson(obj.Name) & " (" & obj.NumberOfCores & " core, " & Format(obj.MaxClockSpeed / 1000, "0.0") & " GHz)"
        json = json & """cpu"":""" & cpuStr & ""","
    Next

    ' RAM
    Set col = wmi.ExecQuery("SELECT TotalPhysicalMemory FROM Win32_ComputerSystem")
    For Each obj In col
        Dim ramGb As Double
        ramGb = CDbl(obj.TotalPhysicalMemory) / (1024 ^ 3)
        json = json & """ram"":""" & Format(ramGb, "0.0") & " GB" & ""","
    Next

    ' Ekran kartı
    Set col = wmi.ExecQuery("SELECT Name, AdapterRAM FROM Win32_VideoController")
    Dim gpuStr As String : gpuStr = ""
    For Each obj In col
        If gpuStr <> "" Then gpuStr = gpuStr & " | "
        Dim vramMb As Long
        On Error Resume Next
        vramMb = CLng(obj.AdapterRAM) / (1024 ^ 2)
        On Error GoTo 0
        gpuStr = gpuStr & EscJson(obj.Name)
        If vramMb > 0 Then gpuStr = gpuStr & " (" & vramMb & " MB)"
    Next
    If gpuStr <> "" Then json = json & """gpu"":""" & gpuStr & ""","

    ' Ekran çözünürlüğü
    Dim sw As Long : sw = 0 : Dim sh As Long : sh = 0
    Set col = wmi.ExecQuery("SELECT CurrentHorizontalResolution, CurrentVerticalResolution FROM Win32_VideoController")
    For Each obj In col
        On Error Resume Next
        sw = CLng(obj.CurrentHorizontalResolution)
        sh = CLng(obj.CurrentVerticalResolution)
        On Error GoTo 0
        If sw > 0 Then Exit For
    Next
    If sw > 0 Then json = json & """screenResolution"":""" & sw & "x" & sh & ""","

    ' BIOS
    Set col = wmi.ExecQuery("SELECT Manufacturer, Name, SMBIOSBIOSVersion, ReleaseDate FROM Win32_BIOS")
    For Each obj In col
        json = json & """bios"":""" & EscJson(obj.Manufacturer) & " - " & EscJson(obj.SMBIOSBIOSVersion) & ""","
    Next

    ' Anakart
    Set col = wmi.ExecQuery("SELECT Manufacturer, Product, SerialNumber FROM Win32_BaseBoard")
    For Each obj In col
        json = json & """motherboard"":""" & EscJson(obj.Manufacturer) & " " & EscJson(obj.Product) & ""","
    Next

    ' ── Ağ bilgileri ──────────────────────────────────
    json = json & """mac"":""" & EscJson(mac) & ""","
    Set col = wmi.ExecQuery("SELECT IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
    Dim ipStr As String : ipStr = ""
    For Each obj In col
        If Not IsNull(obj.IPAddress) Then
            Dim ipArr As Variant : ipArr = obj.IPAddress
            Dim i As Long
            For i = 0 To UBound(ipArr)
                If InStr(ipArr(i), ":") = 0 Then  ' IPv4
                    If ipStr <> "" Then ipStr = ipStr & " | "
                    ipStr = ipStr & ipArr(i)
                End If
            Next i
        End If
    Next
    If ipStr <> "" Then json = json & """ip"":""" & EscJson(ipStr) & ""","

    ' Dış IP
    On Error Resume Next
    Dim httpIp As Object
    Set httpIp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    httpIp.Open "GET", "https://api.ipify.org", False
    httpIp.setTimeouts 3000, 3000, 5000, 5000
    httpIp.send
    If httpIp.Status = 200 Then
        json = json & """publicIp"":""" & EscJson(Trim(httpIp.responseText)) & ""","
    End If
    On Error GoTo 0

    ' ── Disk bilgileri ────────────────────────────────
    Set col = wmi.ExecQuery("SELECT Caption, Size, FreeSpace, FileSystem FROM Win32_LogicalDisk WHERE DriveType=3")
    Dim diskStr As String : diskStr = ""
    For Each obj In col
        If diskStr <> "" Then diskStr = diskStr & "; "
        Dim totGb As Double : totGb = CDbl(obj.Size) / (1024 ^ 3)
        Dim freeGb As Double : freeGb = CDbl(obj.FreeSpace) / (1024 ^ 3)
        diskStr = diskStr & obj.Caption & " " & Format(totGb, "0.0") & "GB (Boş:" & Format(freeGb, "0.0") & "GB " & obj.FileSystem & ")"
    Next
    If diskStr <> "" Then json = json & """disks"":""" & EscJson(diskStr) & ""","

    ' ── Pil ───────────────────────────────────────────
    Set col = wmi.ExecQuery("SELECT EstimatedChargeRemaining, BatteryStatus FROM Win32_Battery")
    For Each obj In col
        Dim batStr As String : batStr = obj.EstimatedChargeRemaining & "%"
        Select Case obj.BatteryStatus
            Case 1: batStr = batStr & " (Boşalıyor)"
            Case 2: batStr = batStr & " (Şarj oluyor)"
            Case Else: batStr = batStr & " (Durum:" & obj.BatteryStatus & ")"
        End Select
        json = json & """battery"":""" & EscJson(batStr) & ""","
    Next

    ' ── Domain/TimeZone/Activation ────────────────────
    Set col = wmi.ExecQuery("SELECT Domain, PartOfDomain FROM Win32_ComputerSystem")
    For Each obj In col
        json = json & """domainName"":""" & EscJson(obj.Domain) & ""","
    Next

    Set col = wmi.ExecQuery("SELECT Description FROM Win32_TimeZone")
    For Each obj In col
        json = json & """timeZone"":""" & EscJson(obj.Description) & ""","
    Next

    ' Windows Aktivasyon (slmgr — max ~8 sn bekle, takilmasin)
    On Error Resume Next
    Dim wsh2 As Object : Set wsh2 = CreateObject("WScript.Shell")
    Dim tmpFile2 As String : tmpFile2 = Environ("TEMP") & "\lic_check.txt"
    On Error Resume Next
    Dim fso2 As Object : Set fso2 = CreateObject("Scripting.FileSystemObject")
    If fso2.FileExists(tmpFile2) Then fso2.DeleteFile tmpFile2
    wsh2.Run "cmd /c cscript //NoLogo %windir%\System32\slmgr.vbs /xpr > """ & tmpFile2 & """", 0, False
    Dim tStart As Single : tStart = Timer
    Do While Not fso2.FileExists(tmpFile2) And (Timer - tStart) < 8
        DoEvents
        Application.Wait Now + TimeValue("00:00:01")
    Loop
    If fso2.FileExists(tmpFile2) Then
        Dim f2 As Object : Set f2 = fso2.OpenTextFile(tmpFile2, 1)
        Dim actTxt As String : actTxt = f2.ReadAll : f2.Close
        fso2.DeleteFile tmpFile2
        actTxt = Trim(actTxt)
        If Len(actTxt) > 200 Then actTxt = Left(actTxt, 200) & "..."
        json = json & """windowsActivation"":""" & EscJson(actTxt) & ""","
    End If
    On Error GoTo 0

    ' ── Sistem çalışma süresi ─────────────────────────
    Set col = wmi.ExecQuery("SELECT LastBootUpTime FROM Win32_OperatingSystem")
    For Each obj In col
        On Error Resume Next
        Dim bootTimeRaw As String : bootTimeRaw = CStr(obj.LastBootUpTime)
        Dim bootYear As Integer : bootYear = CInt(Left(bootTimeRaw, 4))
        Dim bootMonth As Integer : bootMonth = CInt(Mid(bootTimeRaw, 5, 2))
        Dim bootDay As Integer : bootDay = CInt(Mid(bootTimeRaw, 7, 2))
        Dim bootHour As Integer : bootHour = CInt(Mid(bootTimeRaw, 9, 2))
        Dim bootMin As Integer : bootMin = CInt(Mid(bootTimeRaw, 11, 2))
        Dim bootDt As Date : bootDt = DateSerial(bootYear, bootMonth, bootDay) + TimeSerial(bootHour, bootMin, 0)
        Dim uptimeMins As Long : uptimeMins = DateDiff("n", bootDt, Now)
        Dim uptimeDays As Long : uptimeDays = uptimeMins \ 1440
        Dim uptimeHrs As Long : uptimeHrs = (uptimeMins Mod 1440) \ 60
        Dim uptimeMn As Long : uptimeMn = uptimeMins Mod 60
        json = json & """systemUptime"":""" & uptimeDays & " gün " & uptimeHrs & " sa " & uptimeMn & " dk" & ""","
        On Error GoTo 0
    Next

    ' ── JSON kapat ─────────────────────────────────────
    ' Son virgülü kaldır
    If Right(json, 1) = "," Then json = Left(json, Len(json) - 1)
    json = json & "}"

    ' ── Sunucuya POST ──────────────────────────────────
    Dim postBody As String
    postBody = "{" & _
        """mac"":""" & EscJson(mac) & """," & _
        """hostname"":""" & EscJson(hostname) & """," & _
        """firmaAdi"":""" & EscJson(firmaAdi) & """," & _
        """data"":" & json & _
        "}"

    On Error Resume Next
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", apiBase & "device-info/", False
    http.setRequestHeader "Content-Type", "application/json"
    http.setTimeouts 5000, 5000, 20000, 20000
    http.send postBody

    If http.Status = 200 Then
        ' İsteğe bağlı: başarı mesajı
        ' MsgBox "Cihaz bilgileri sunucuya gönderildi.", vbInformation
    End If
    On Error GoTo 0

    Set DynamicFunc = Nothing
End Function

Private Function EscJson(s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\" & Chr(34))
    s = Replace(s, Chr(10), "\n")
    s = Replace(s, Chr(13), "")
    s = Replace(s, Chr(9), "\t")
    EscJson = s
End Function
