#If Win64 Then
Private Const AGENT_ARCH As String = "x64"
#Else
Private Const AGENT_ARCH As String = "x86"
#End If

#If Win64 Then
Private Const REGASM_PATH As String = "\Microsoft.NET\Framework64\v4.0.30319\regasm.exe"
#Else
Private Const REGASM_PATH As String = "\Microsoft.NET\Framework\v4.0.30319\regasm.exe"
#End If

Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
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
    StopAgentForUpdate agentDir

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    http.Open "GET", baseUrl & "agent/download/?arch=" & AGENT_ARCH, False
    http.setTimeouts 10000, 10000, 120000, 120000
    http.send
    If http.Status <> 200 Then
        Debug.Print "[InstallTeklifAgent] DLL indirilemedi HTTP " & http.Status
        GoTo Done
    End If
    Dim bin() As Byte
    bin = http.responseBody
    SaveBinarySafe agentDir & "\TeklifAgent.Com.dll", bin

    http.Open "GET", baseUrl & "agent/download/?arch=" & AGENT_ARCH & "&file=exe", False
    http.send
    If http.Status <> 200 Then
        Debug.Print "[InstallTeklifAgent] EXE indirilemedi HTTP " & http.Status
        GoTo Done
    End If
    bin = http.responseBody
    SaveBinarySafe agentDir & "\TeklifAgent.exe", bin
    Set http = Nothing

    Dim comOk As Boolean : comOk = TryRegisterCom(agentDir)

    Dim mac As String : mac = GetFirstMACAddress()
    Dim hostname As String : hostname = Environ("COMPUTERNAME")
    Dim usr As String : usr = Environ("USERNAME")
    Dim excelVer As String
    On Error Resume Next
    excelVer = Application.Version
    If Len(excelVer) = 0 Then excelVer = "agent-only"
    On Error GoTo 0
    Dim intervalMin As Long : intervalMin = 1

    If Dir(agentDir & "\stop.flag") <> "" Then Kill agentDir & "\stop.flag"
    If Dir(agentDir & "\excel-session.ready") <> "" Then Kill agentDir & "\excel-session.ready"

    Dim cfgJson As String
    cfgJson = "{""ApiBaseUrl"":""" & JsonEsc(baseUrl) & """,""Mac"":""" & JsonEsc(mac) & _
              """,""Hostname"":""" & JsonEsc(hostname) & """,""User"":""" & JsonEsc(usr) & _
              """,""ExcelVersion"":""" & JsonEsc(excelVer) & _
              """,""IntervalMinutes"":" & intervalMin & ",""Stop"":false}"
    WriteTextFile agentDir & "\config.json", cfgJson

    Call RegisterLogonScheduledTask(agentDir)

    Dim exeOk As Boolean
    exeOk = StartAgentExe(agentDir)

    Debug.Print "[InstallTeklifAgent] Kurulum tamam | COM=" & IIf(comOk, "OK", "yok") & _
                " | exe=" & IIf(exeOk, "calisiyor", "hata") & _
                " | ping Excel acilinca (HeartbeatPing)"

Done:
    Set DynamicFunc = Nothing
End Function

Private Function TryRegisterCom(agentDir As String) As Boolean
    On Error GoTo Fail
    Dim regasm As String
    regasm = Environ("WINDIR") & REGASM_PATH
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
    On Error Resume Next
    Dim sh0 As Object : Set sh0 = CreateObject("WScript.Shell")
    sh0.Run "taskkill /F /IM TeklifAgent.exe", 0, True
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
        On Error GoTo Fail
    End If
    fso.MoveFile tmpPath, destPath
    Exit Sub
Fail:
    Err.Raise vbObjectError + 3004, , "Dosyaya yazilamadi: " & destPath
End Sub

Private Sub RegisterLogonScheduledTask(agentDir As String)
    On Error GoTo Fail
    Dim exePath As String
    exePath = agentDir & "\TeklifAgent.exe"
    If Dir(exePath) = "" Then Exit Sub
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    Dim taskName As String
    taskName = "TeklifAgentBoot"
    sh.Run "schtasks /Delete /TN """ & taskName & """ /F", 0, True
    Dim cmd As String
    cmd = "schtasks /Create /F /SC ONLOGON /RL LIMITED /TN """ & taskName & """ /TR """ & exePath & " --worker"""
    sh.Run cmd, 0, True
    Debug.Print "[InstallTeklifAgent] Gorev zamanlayici: " & taskName
    Exit Sub
Fail:
    Debug.Print "[InstallTeklifAgent] Logon task kurulamadi: " & Err.Description
End Sub

Private Function JsonEsc(s As String) As String
    JsonEsc = Replace(Replace(s, "\", "\\"), """", "\""")
End Function
