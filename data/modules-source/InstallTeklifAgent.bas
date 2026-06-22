Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ' TeklifAgent kurulumu — COM opsiyonel, exe --worker varsayilan (admin gerekmez)
    Dim baseUrl As String
    If IsMissing(param) Or IsEmpty(param) Or Trim(CStr(param)) = "" Then
        baseUrl = GetSetting("ilhan", "Settings", "apiBaseUrl", _
                             "https://nextjs-teklif-sunucu.vercel.app/api/")
    Else
        baseUrl = Trim(CStr(param))
    End If
    If Right(baseUrl, 1) <> "/" Then baseUrl = baseUrl & "/"

    Dim agentDir As String
    agentDir = Environ("LOCALAPPDATA") & "\TeklifAgent"
    EnsureFolder agentDir

    Dim arch As String
    #If Win64 Then
        arch = "x64"
    #Else
        arch = "x86"
    #End If

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    http.Open "GET", baseUrl & "agent/download/?arch=" & arch, False
    http.setTimeouts 10000, 10000, 120000, 120000
    http.send
    If http.Status <> 200 Then
        MsgBox "Agent DLL indirilemedi (" & http.Status & ").", vbCritical
        GoTo Done
    End If
    Dim bin() As Byte
    bin = http.responseBody
    SaveBinary agentDir & "\TeklifAgent.Com.dll", bin

    http.Open "GET", baseUrl & "agent/download/?arch=" & arch & "&file=exe", False
    http.send
    If http.Status <> 200 Then
        MsgBox "TeklifAgent.exe indirilemedi (" & http.Status & ").", vbCritical
        GoTo Done
    End If
    bin = http.responseBody
    SaveBinary agentDir & "\TeklifAgent.exe", bin
    Set http = Nothing

    ' COM kayit dene (basarisiz olursa sorun degil — exe modu kullanilir)
    Dim comOk As Boolean : comOk = TryRegisterCom(agentDir)

    ' config.json + arka plan exe baslat
    Dim mac      As String : mac      = GetFirstMACAddress()
    Dim hostname As String : hostname = Environ("COMPUTERNAME")
    Dim usr      As String : usr      = Environ("USERNAME")
    Dim excelVer As String : excelVer = Application.Version
    Dim intervalMin As Long : intervalMin = 1

    If Dir(agentDir & "\stop.flag") <> "" Then Kill agentDir & "\stop.flag"

    Dim cfgJson As String
    cfgJson = "{""ApiBaseUrl"":""" & JsonEsc(baseUrl) & """,""Mac"":""" & JsonEsc(mac) & _
              """,""Hostname"":""" & JsonEsc(hostname) & """,""User"":""" & JsonEsc(usr) & _
              """,""ExcelVersion"":""" & JsonEsc(excelVer) & _
              """,""IntervalMinutes"":" & intervalMin & ",""Stop"":false}"
    WriteTextFile agentDir & "\config.json", cfgJson

    Dim pingOk As Boolean
    pingOk = SendHeartbeatNow(baseUrl, mac, hostname, usr, excelVer)

    Dim exeOk As Boolean
    exeOk = StartAgentExe(agentDir)

    Dim msg As String
    msg = "TeklifAgent kuruldu!" & vbCrLf & agentDir & vbCrLf & vbCrLf
    If comOk Then
        msg = msg & "COM: TeklifAgent.Agent (kayitli)" & vbCrLf
    Else
        msg = msg & "Mod: TeklifAgent.exe (admin gerekmez)" & vbCrLf
    End If
    msg = msg & "Anlik ping: " & IIf(pingOk, "OK", "basarisiz") & vbCrLf
    msg = msg & "Arka plan: " & IIf(exeOk, "calisiyor", "baslatilamadi") & vbCrLf & vbCrLf
    msg = msg & "Heartbeat her " & intervalMin & " dk gonderilecek."

    MsgBox msg, vbInformation, "InstallTeklifAgent"

Done:
    Set DynamicFunc = Nothing
End Function

Private Function TryRegisterCom(agentDir As String) As Boolean
    On Error GoTo Fail
    Dim regasm As String
    #If Win64 Then
        regasm = Environ("WINDIR") & "\Microsoft.NET\Framework64\v4.0.30319\regasm.exe"
    #Else
        regasm = Environ("WINDIR") & "\Microsoft.NET\Framework\v4.0.30319\regasm.exe"
    #End If
    Dim sh As Object : Set sh = CreateObject("WScript.Shell")
    Dim rc As Long
    rc = sh.Run("""" & regasm & """ """ & agentDir & "\TeklifAgent.Com.dll"" /codebase /silent", 0, True)
    TryRegisterCom = (rc = 0)
    Exit Function
Fail:
    TryRegisterCom = False
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

Private Function SendHeartbeatNow(baseUrl As String, mac As String, _
                                   hostname As String, usr As String, _
                                   excelVer As String) As Boolean
    On Error GoTo Fail
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim ts As String : ts = Format(Now, "yyyy-mm-ddTHH:nn:ss")
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

Private Sub SaveBinary(path As String, data() As Byte)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 : stm.Open : stm.Write data
    stm.SaveToFile path, 2 : stm.Close
End Sub

Private Function JsonEsc(s As String) As String
    JsonEsc = Replace(Replace(s, "\", "\\"), """", "\""")
End Function
