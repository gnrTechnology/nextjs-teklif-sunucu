#If Win64 Then
Private Const AGENT_ARCH As String = "x64"
#Else
Private Const AGENT_ARCH As String = "x86"
#End If

Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = Trim(CStr(param))

    Dim stopFlag As Boolean : stopFlag = False
    Dim intervalMin As Long : intervalMin = 1

    If Left(p, 1) = "{" Then
        If LCase(ExtractJsonValue(p, "stop")) = "true" Then stopFlag = True
        Dim intStr As String : intStr = ExtractJsonValue(p, "intervalMin")
        If Len(intStr) > 0 Then intervalMin = CLng(intStr)
    ElseIf IsNumeric(p) Then
        intervalMin = CLng(p)
    End If

    Dim baseUrl As String
    baseUrl = GetSetting("ilhan", "Settings", "apiBaseUrl", _
                         "https://nextjs-teklif-sunucu.vercel.app/api/")
    If Right(baseUrl, 1) <> "/" Then baseUrl = baseUrl & "/"

    If stopFlag Then
        Call StopTeklifAgent
        SaveSetting "ilhan", "Heartbeat", "active", "false"
        Debug.Print "[HeartbeatPing] TeklifAgent durduruldu."
        Set DynamicFunc = Nothing
        Exit Function
    End If

    Dim mac As String : mac = GetFirstMACAddress()
    Dim hostname As String : hostname = Environ("COMPUTERNAME")
    Dim usr As String : usr = Environ("USERNAME")
    Dim excelVer As String : excelVer = Application.Version

    ' Ilk Excel oturumu: agent ping/komut icin hazir bayragi
    MarkExcelSessionReady

    Dim pingOk As Boolean
    pingOk = SendHeartbeatNow(baseUrl, mac, hostname, usr, excelVer)
    Debug.Print "[HeartbeatPing] Ilk Excel ping: " & IIf(pingOk, "OK", "HATA")

    Dim agentOk As Boolean
    agentOk = EnsureAndStartAgent(baseUrl, mac, hostname, usr, excelVer, intervalMin)

    SaveSetting "ilhan", "Heartbeat", "active", "true"
    SaveSetting "ilhan", "Heartbeat", "intervalMin", CStr(intervalMin)

    Debug.Print "[HeartbeatPing] Agent: " & IIf(agentOk, "calisiyor", "baslatilamadi") & _
                " | aralik " & intervalMin & " dk"

    Set DynamicFunc = Nothing
End Function

Private Sub MarkExcelSessionReady()
    Dim agentDir As String
    agentDir = Environ("LOCALAPPDATA") & "\TeklifAgent"
    EnsureFolder agentDir
    WriteTextFile agentDir & "\excel-session.ready", Format(Now, "yyyy-mm-ddTHH:nn:ss")
    Debug.Print "[HeartbeatPing] excel-session.ready yazildi"
End Sub

Private Function SendHeartbeatNow(baseUrl As String, mac As String, _
                                   hostname As String, usr As String, _
                                   excelVer As String) As Boolean
    On Error GoTo Fail
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim ts As String
    ts = Format(Now, "yyyy-mm-ddTHH:nn:ss")
    Dim body As String
    body = "{""mac"":""" & JsonEsc(mac) & """,""hostname"":""" & JsonEsc(hostname) & _
           """,""user"":""" & JsonEsc(usr) & """,""excelVersion"":""" & JsonEsc(excelVer) & _
           """,""timestamp"":""" & ts & """}"
    http.Open "POST", baseUrl & "heartbeat/", False
    http.setTimeouts 5000, 5000, 15000, 15000
    http.setRequestHeader "Content-Type", "application/json;charset=UTF-8"
    http.send body
    SendHeartbeatNow = (http.Status = 200)
    Set http = Nothing
    Exit Function
Fail:
    SendHeartbeatNow = False
End Function

Private Function EnsureAndStartAgent(baseUrl As String, mac As String, _
                                    hostname As String, usr As String, _
                                    excelVer As String, intervalMin As Long) As Boolean
    On Error GoTo Fail

    Dim agentDir As String
    agentDir = Environ("LOCALAPPDATA") & "\TeklifAgent"
    EnsureFolder agentDir

    If Not AgentFilesReady(agentDir) Then
        If Not DownloadAgentFiles(baseUrl, agentDir) Then GoTo Fail
    End If

    Dim cfgPath As String : cfgPath = agentDir & "\config.json"
    Dim cfgJson As String
    cfgJson = "{""ApiBaseUrl"":""" & JsonEsc(baseUrl) & """,""Mac"":""" & JsonEsc(mac) & _
              """,""Hostname"":""" & JsonEsc(hostname) & """,""User"":""" & JsonEsc(usr) & _
              """,""ExcelVersion"":""" & JsonEsc(excelVer) & _
              """,""IntervalMinutes"":" & intervalMin & ",""Stop"":false}"
    WriteTextFile cfgPath, cfgJson

    If Dir(agentDir & "\stop.flag") <> "" Then Kill agentDir & "\stop.flag"

    If TryStartViaCom(cfgJson) Then
        EnsureAndStartAgent = True
        Exit Function
    End If

    EnsureAndStartAgent = StartAgentExe(agentDir)
    Exit Function
Fail:
    EnsureAndStartAgent = False
End Function

Private Function AgentFilesReady(agentDir As String) As Boolean
    AgentFilesReady = (Dir(agentDir & "\TeklifAgent.exe") <> "") _
                   Or (Dir(agentDir & "\TeklifAgent.Com.dll") <> "")
End Function

Private Function DownloadAgentFiles(baseUrl As String, agentDir As String) As Boolean
    On Error GoTo Fail
    StopAgentForUpdate agentDir

    Dim arch As String
    arch = AGENT_ARCH

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", baseUrl & "agent/download/?arch=" & arch, False
    http.setTimeouts 10000, 10000, 60000, 60000
    http.send
    If http.Status <> 200 Then GoTo Fail

    Dim bin() As Byte
    bin = http.responseBody
    SaveBinarySafe agentDir & "\TeklifAgent.Com.dll", bin

    http.Open "GET", baseUrl & "agent/download/?arch=" & arch & "&file=exe", False
    http.send
    If http.Status = 200 Then
        bin = http.responseBody
        SaveBinarySafe agentDir & "\TeklifAgent.exe", bin
    End If

    Set http = Nothing
    DownloadAgentFiles = True
    Exit Function
Fail:
    DownloadAgentFiles = False
End Function

Private Function TryStartViaCom(cfgJson As String) As Boolean
    On Error GoTo Fail
    Dim agent As Object
    Set agent = CreateObject("TeklifAgent.Agent")
    agent.Start cfgJson
    TryStartViaCom = (LCase(agent.GetStatus()) = "running")
    Exit Function
Fail:
    TryStartViaCom = False
End Function

Private Function StartAgentExe(agentDir As String) As Boolean
    On Error GoTo Fail
    On Error Resume Next
    Dim sh0 As Object : Set sh0 = CreateObject("WScript.Shell")
    sh0.Run "taskkill /F /IM TeklifAgent.exe", 0, True
    On Error GoTo Fail
    Application.Wait Now + TimeValue("00:00:01")
    Dim exePath As String : exePath = agentDir & "\TeklifAgent.exe"
    If Dir(exePath) = "" Then Exit Function
    Dim sh As Object : Set sh = CreateObject("WScript.Shell")
    sh.Run """" & exePath & """ --worker", 0, False
    StartAgentExe = True
    Exit Function
Fail:
    StartAgentExe = False
End Function

Private Sub StopTeklifAgent()
    On Error Resume Next
    Dim agentDir As String : agentDir = Environ("LOCALAPPDATA") & "\TeklifAgent"
    Open agentDir & "\stop.flag" For Output As #1 : Close #1
    Dim readyPath As String : readyPath = agentDir & "\excel-session.ready"
    If Dir(readyPath) <> "" Then Kill readyPath
    Dim agent As Object
    Set agent = CreateObject("TeklifAgent.Agent")
    agent.Stop
    Dim sh As Object : Set sh = CreateObject("WScript.Shell")
    sh.Run "taskkill /F /IM TeklifAgent.exe", 0, True
End Sub

Private Sub EnsureFolder(path As String)
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(path) Then fso.CreateFolder path
End Sub

Private Sub WriteTextFile(path As String, content As String)
    Dim fNum As Integer : fNum = FreeFile
    Open path For Output As #fNum
    Print #fNum, content;
    Close #fNum
End Sub

Private Sub StopAgentForUpdate(agentDir As String)
    On Error Resume Next
    Open agentDir & "\stop.flag" For Output As #1 : Close #1
    Dim sh As Object : Set sh = CreateObject("WScript.Shell")
    sh.Run "taskkill /F /IM TeklifAgent.exe", 0, True
    Application.Wait Now + TimeValue("00:00:02")
    On Error GoTo 0
End Sub

Private Sub SaveBinarySafe(destPath As String, data() As Byte)
    On Error GoTo Fail
    Dim tmpPath As String
    tmpPath = Environ("TEMP") & "\TeklifAgent_" & CLng(Timer * 1000) & ".tmp"
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1
    stm.Open
    stm.Write data
    stm.SaveToFile tmpPath, 2
    stm.Close
    Set stm = Nothing
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(destPath) Then
        On Error Resume Next
        fso.DeleteFile destPath, True
        If Err.Number <> 0 Then
            Err.Clear
            StopAgentForUpdate Left(destPath, InStrRev(destPath, "\"))
            fso.DeleteFile destPath, True
        End If
        On Error GoTo Fail
    End If
    fso.MoveFile tmpPath, destPath
    Exit Sub
Fail:
    Err.Raise vbObjectError + 3004, , "Dosyaya yazilamadi: " & destPath
End Sub

Private Function JsonEsc(s As String) As String
    JsonEsc = Replace(Replace(s, "\", "\\"), """", "\""")
End Function

Private Function GetFirstMACAddress() As String
    Dim objWMI As Object, col As Object, obj As Object
    On Error GoTo Fail
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
    GetFirstMACAddress = "UNKNOWN"
    For Each obj In col
        If Not IsNull(obj.MACAddress) And obj.MACAddress <> "" Then
            GetFirstMACAddress = obj.MACAddress : Exit For
        End If
    Next
Fail:
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
    Else
        p2 = p1
        Do While p2 <= Len(json)
            If InStr(",}] " & Chr(13) & Chr(10), Mid(json, p2, 1)) > 0 Then Exit Do
            p2 = p2 + 1
        Loop
        ExtractJsonValue = Trim(Mid(json, p1, p2 - p1))
    End If
End Function
