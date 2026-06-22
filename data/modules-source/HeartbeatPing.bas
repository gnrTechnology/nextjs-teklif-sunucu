Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ' param: {"intervalMin":1,"stop":false}  veya "1" (dakika)
    Dim p As String : p = Trim(CStr(param))

    Dim stopFlag    As Boolean : stopFlag    = False
    Dim intervalMin As Long    : intervalMin = 1

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

    ' ----- DURDUR -----
    If stopFlag Then
        Call StopTeklifAgent
        SaveSetting "ilhan", "Heartbeat", "active", "false"
        MsgBox "TeklifAgent durduruldu.", vbInformation, "HeartbeatPing"
        Set DynamicFunc = Nothing
        Exit Function
    End If

    Dim mac      As String : mac      = GetFirstMACAddress()
    Dim hostname As String : hostname = Environ("COMPUTERNAME")
    Dim usr      As String : usr      = Environ("USERNAME")
    Dim excelVer As String : excelVer = Application.Version

    ' ----- 1) İLK AÇILIŞTA HEMEN PING (VBA — anında) -----
    Dim pingOk As Boolean
    pingOk = SendHeartbeatNow(baseUrl, mac, hostname, usr, excelVer)
    Debug.Print "[HeartbeatPing] Anlik ping: " & IIf(pingOk, "OK", "HATA")

    ' ----- 2) TeklifAgent DLL/exe kur ve başlat -----
    Dim agentOk As Boolean
    agentOk = EnsureAndStartAgent(baseUrl, mac, hostname, usr, excelVer, intervalMin)

    SaveSetting "ilhan", "Heartbeat", "active",      "true"
    SaveSetting "ilhan", "Heartbeat", "intervalMin",  CStr(intervalMin)

    If agentOk Then
        MsgBox "Heartbeat aktif!" & vbCrLf & _
               "Anlik ping: " & IIf(pingOk, "gonderildi", "basarisiz") & vbCrLf & _
               "TeklifAgent arka planda calisiyor (her " & intervalMin & " dk).", _
               vbInformation, "HeartbeatPing"
    Else
        MsgBox "Anlik ping: " & IIf(pingOk, "OK", "HATA") & vbCrLf & _
               "TeklifAgent baslatilamadi — InstallTeklifAgent modulunu calistirin.", _
               vbExclamation, "HeartbeatPing"
    End If

    Set DynamicFunc = Nothing
End Function

' ─── Anında heartbeat (Excel içinden, MSXML) ───────────────────────────────
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

' ─── Agent kurulum + başlatma ───────────────────────────────────────────────
Private Function EnsureAndStartAgent(baseUrl As String, mac As String, _
                                    hostname As String, usr As String, _
                                    excelVer As String, intervalMin As Long) As Boolean
    On Error GoTo Fail

    Dim agentDir As String
    agentDir = Environ("LOCALAPPDATA") & "\TeklifAgent"
    EnsureFolder agentDir

    ' Agent dosyaları yoksa sunucudan indir
    If Not AgentFilesReady(agentDir) Then
        If Not DownloadAgentFiles(baseUrl, agentDir) Then GoTo Fail
    End If

    ' config.json yaz
    Dim cfgPath As String : cfgPath = agentDir & "\config.json"
    Dim cfgJson As String
    cfgJson = "{""ApiBaseUrl"":""" & JsonEsc(baseUrl) & """,""Mac"":""" & JsonEsc(mac) & _
              """,""Hostname"":""" & JsonEsc(hostname) & """,""User"":""" & JsonEsc(usr) & _
              """,""ExcelVersion"":""" & JsonEsc(excelVer) & _
              """,""IntervalMinutes"":" & intervalMin & ",""Stop"":false}"
    WriteTextFile cfgPath, cfgJson

    ' Stop flag temizle
    Dim stopPath As String : stopPath = agentDir & "\stop.flag"
    If Dir(stopPath) <> "" Then Kill stopPath

    ' Önce COM dene
    If TryStartViaCom(cfgJson) Then
        EnsureAndStartAgent = True
        Exit Function
    End If

    ' COM yoksa exe --worker
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
    Dim arch As String
    #If Win64 Then
        arch = "x64"
    #Else
        arch = "x86"
    #End If

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", baseUrl & "agent/download/?arch=" & arch, False
    http.setTimeouts 10000, 10000, 60000, 60000
    http.send
    If http.Status <> 200 Then GoTo Fail

    ' Yanıt: zip veya tek dll — sunucu dll döndürür
    Dim bin() As Byte
    bin = http.responseBody
    Dim dllPath As String
    dllPath = agentDir & "\TeklifAgent.Com.dll"
    SaveBinaryFile dllPath, bin

    ' exe de indir
    http.Open "GET", baseUrl & "agent/download/?arch=" & arch & "&file=exe", False
    http.send
    If http.Status = 200 Then
        bin = http.responseBody
        SaveBinaryFile agentDir & "\TeklifAgent.exe", bin
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
    Dim agent As Object
    Set agent = CreateObject("TeklifAgent.Agent")
    agent.Stop
    ' exe process sonlandir
    Dim sh As Object : Set sh = CreateObject("WScript.Shell")
    sh.Run "taskkill /F /IM TeklifAgent.exe", 0, True
End Sub

' ─── Yardımcılar ────────────────────────────────────────────────────────────
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

Private Sub SaveBinaryFile(path As String, data() As Byte)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1
    stm.Open
    stm.Write data
    stm.SaveToFile path, 2
    stm.Close
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
