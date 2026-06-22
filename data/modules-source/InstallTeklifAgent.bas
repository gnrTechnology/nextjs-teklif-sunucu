Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ' TeklifAgent COM DLL + exe kurulumu
    ' param: opsiyonel apiBaseUrl (string)
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

    ' DLL indir
    http.Open "GET", baseUrl & "agent/download/?arch=" & arch, False
    http.setTimeouts 10000, 10000, 120000, 120000
    http.send
    If http.Status <> 200 Then
        MsgBox "Agent DLL indirilemedi (" & http.Status & ").", vbCritical
        GoTo Done
    End If
    SaveBinary agentDir & "\TeklifAgent.Com.dll", http.responseBody

    ' EXE indir
    http.Open "GET", baseUrl & "agent/download/?arch=" & arch & "&file=exe", False
    http.send
    If http.Status = 200 Then
        SaveBinary agentDir & "\TeklifAgent.exe", http.responseBody
    End If
    Set http = Nothing

    ' COM kayit (regasm — yonetici gerekebilir)
    Dim regasm As String
    #If Win64 Then
        regasm = Environ("WINDIR") & "\Microsoft.NET\Framework64\v4.0.30319\regasm.exe"
    #Else
        regasm = Environ("WINDIR") & "\Microsoft.NET\Framework\v4.0.30319\regasm.exe"
    #End If

    Dim sh As Object : Set sh = CreateObject("WScript.Shell")
    Dim rc As Long
    rc = sh.Run("""" & regasm & """ """ & agentDir & "\TeklifAgent.Com.dll"" /codebase /silent", 0, True)

    If rc = 0 Then
        MsgBox "TeklifAgent kuruldu:" & vbCrLf & agentDir & vbCrLf & _
               "COM: TeklifAgent.Agent", vbInformation
    Else
        MsgBox "DLL kopyalandi ancak COM kayit basarisiz (rc=" & rc & ")." & vbCrLf & _
               "Yonetici olarak calistirin veya TeklifAgent.exe --worker kullanilir.", vbExclamation
    End If

Done:
    Set DynamicFunc = Nothing
End Function

Private Sub EnsureFolder(path As String)
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(path) Then fso.CreateFolder path
End Sub

Private Sub SaveBinary(path As String, data() As Byte)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 : stm.Open : stm.Write data
    stm.SaveToFile path, 2 : stm.Close
End Sub
