Private Function EscJson(s As String) As String
    Dim i As Long, c As String, out As String, ac As Long
    out = ""
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        ac = AscW(c)
        Select Case ac
            Case 92
                out = out & Chr(92) & Chr(92)
            Case 34
                out = out & Chr(92) & Chr(34)
            Case 10
                out = out & Chr(92) & "n"
            Case 13
                ' atla
            Case 9
                out = out & Chr(92) & "t"
            Case Else
                out = out & c
        End Select
    Next i
    EscJson = out
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

    If mac = "" Then Set DynamicFunc = Nothing : Exit Function

    Dim json As String
    json = "{"

    json = json & """computerName"":""" & EscJson(hostname) & ""","
    json = json & """loggedInUser"":""" & EscJson(Environ("USERNAME")) & ""","

    Dim wmi As Object
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Dim col As Object, obj As Object

    Set col = wmi.ExecQuery("SELECT Caption, BuildNumber, Version FROM Win32_OperatingSystem")
    For Each obj In col
        json = json & """windowsVersion"":""" & EscJson(CStr(obj.Caption) & " (Build " & obj.BuildNumber & ")") & ""","
    Next

    Set col = wmi.ExecQuery("SELECT Name, NumberOfCores, MaxClockSpeed FROM Win32_Processor")
    For Each obj In col
        Dim cpuRaw As String
        cpuRaw = CStr(obj.Name) & " (" & obj.NumberOfCores & " core, " & Format(obj.MaxClockSpeed / 1000, "0.0") & " GHz)"
        json = json & """cpu"":""" & EscJson(cpuRaw) & ""","
    Next

    Set col = wmi.ExecQuery("SELECT TotalPhysicalMemory FROM Win32_ComputerSystem")
    For Each obj In col
        Dim ramGb As Double
        ramGb = CDbl(obj.TotalPhysicalMemory) / (1024 ^ 3)
        json = json & """ram"":""" & EscJson(Format(ramGb, "0.0") & " GB") & ""","
    Next

    Set col = wmi.ExecQuery("SELECT Name, AdapterRAM FROM Win32_VideoController")
    Dim gpuStr As String
    gpuStr = ""
    For Each obj In col
        If gpuStr <> "" Then gpuStr = gpuStr & " | "
        Dim vramMb As Long
        On Error Resume Next
        vramMb = CLng(obj.AdapterRAM) / (1024 ^ 2)
        On Error GoTo 0
        gpuStr = gpuStr & CStr(obj.Name)
        If vramMb > 0 Then gpuStr = gpuStr & " (" & vramMb & " MB)"
    Next
    If gpuStr <> "" Then json = json & """gpu"":""" & EscJson(gpuStr) & ""","

    Dim sw As Long, sh As Long
    sw = 0 : sh = 0
    Set col = wmi.ExecQuery("SELECT CurrentHorizontalResolution, CurrentVerticalResolution FROM Win32_VideoController")
    For Each obj In col
        On Error Resume Next
        sw = CLng(obj.CurrentHorizontalResolution)
        sh = CLng(obj.CurrentVerticalResolution)
        On Error GoTo 0
        If sw > 0 Then Exit For
    Next
    If sw > 0 Then json = json & """screenResolution"":""" & EscJson(CStr(sw) & "x" & CStr(sh)) & ""","

    Set col = wmi.ExecQuery("SELECT Manufacturer, SMBIOSBIOSVersion FROM Win32_BIOS")
    For Each obj In col
        json = json & """bios"":""" & EscJson(CStr(obj.Manufacturer) & " - " & CStr(obj.SMBIOSBIOSVersion)) & ""","
    Next

    Set col = wmi.ExecQuery("SELECT Manufacturer, Product FROM Win32_BaseBoard")
    For Each obj In col
        json = json & """motherboard"":""" & EscJson(CStr(obj.Manufacturer) & " " & CStr(obj.Product)) & ""","
    Next

    json = json & """mac"":""" & EscJson(mac) & ""","

    Set col = wmi.ExecQuery("SELECT IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
    Dim ipStr As String
    ipStr = ""
    For Each obj In col
        If Not IsNull(obj.IPAddress) Then
            Dim ipArr As Variant
            ipArr = obj.IPAddress
            Dim j As Long
            For j = 0 To UBound(ipArr)
                If InStr(CStr(ipArr(j)), ":") = 0 Then
                    If ipStr <> "" Then ipStr = ipStr & " | "
                    ipStr = ipStr & CStr(ipArr(j))
                End If
            Next j
        End If
    Next
    If ipStr <> "" Then json = json & """ip"":""" & EscJson(ipStr) & ""","

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

    Set col = wmi.ExecQuery("SELECT Caption, Size, FreeSpace, FileSystem FROM Win32_LogicalDisk WHERE DriveType=3")
    Dim diskStr As String
    diskStr = ""
    For Each obj In col
        If diskStr <> "" Then diskStr = diskStr & "; "
        Dim totGb As Double, freeGb As Double
        totGb = CDbl(obj.Size) / (1024 ^ 3)
        freeGb = CDbl(obj.FreeSpace) / (1024 ^ 3)
        diskStr = diskStr & CStr(obj.Caption) & " " & Format(totGb, "0.0") & "GB (Bos:" & Format(freeGb, "0.0") & "GB " & CStr(obj.FileSystem) & ")"
    Next
    If diskStr <> "" Then json = json & """disks"":""" & EscJson(diskStr) & ""","

    Set col = wmi.ExecQuery("SELECT EstimatedChargeRemaining, BatteryStatus FROM Win32_Battery")
    For Each obj In col
        Dim batStr As String
        batStr = CStr(obj.EstimatedChargeRemaining) & "%"
        Select Case CLng(obj.BatteryStatus)
            Case 1: batStr = batStr & " (Bosaliyor)"
            Case 2: batStr = batStr & " (Sarj oluyor)"
            Case Else: batStr = batStr & " (Durum:" & obj.BatteryStatus & ")"
        End Select
        json = json & """battery"":""" & EscJson(batStr) & ""","
    Next

    Set col = wmi.ExecQuery("SELECT Domain FROM Win32_ComputerSystem")
    For Each obj In col
        json = json & """domainName"":""" & EscJson(CStr(obj.Domain)) & ""","
    Next

    Set col = wmi.ExecQuery("SELECT Description FROM Win32_TimeZone")
    For Each obj In col
        json = json & """timeZone"":""" & EscJson(CStr(obj.Description)) & ""","
    Next

    Set col = wmi.ExecQuery("SELECT LastBootUpTime FROM Win32_OperatingSystem")
    For Each obj In col
        On Error Resume Next
        Dim bootTimeRaw As String
        bootTimeRaw = CStr(obj.LastBootUpTime)
        Dim bootYear As Integer, bootMonth As Integer, bootDay As Integer
        Dim bootHour As Integer, bootMin As Integer
        bootYear = CInt(Left(bootTimeRaw, 4))
        bootMonth = CInt(Mid(bootTimeRaw, 5, 2))
        bootDay = CInt(Mid(bootTimeRaw, 7, 2))
        bootHour = CInt(Mid(bootTimeRaw, 9, 2))
        bootMin = CInt(Mid(bootTimeRaw, 11, 2))
        Dim bootDt As Date
        bootDt = DateSerial(bootYear, bootMonth, bootDay) + TimeSerial(bootHour, bootMin, 0)
        Dim uptimeMins As Long
        uptimeMins = DateDiff("n", bootDt, Now)
        Dim uptimeDays As Long, uptimeHrs As Long, uptimeMn As Long
        uptimeDays = uptimeMins \ 1440
        uptimeHrs = (uptimeMins Mod 1440) \ 60
        uptimeMn = uptimeMins Mod 60
        json = json & """systemUptime"":""" & EscJson(CStr(uptimeDays) & " gun " & uptimeHrs & " sa " & uptimeMn & " dk") & ""","
        On Error GoTo 0
    Next

    If Right(json, 1) = "," Then json = Left(json, Len(json) - 1)
    json = json & "}"

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
    On Error GoTo 0

    Set DynamicFunc = Nothing
End Function
