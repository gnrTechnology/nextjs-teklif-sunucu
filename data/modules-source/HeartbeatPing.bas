Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ' param: {"intervalMin":1,"stop":false}  veya sadece "1" (dakika)
    ' Arka planda VBScript ile periyodik heartbeat gönderir.
    ' teklif.xlam'a EK KOD GEREKMEz.
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

    Dim stopFlagPath As String
    stopFlagPath = Environ("TEMP") & "\hb_stop.flag"

    ' ----- DURDUR -----
    If stopFlag Then
        Open stopFlagPath For Output As #1 : Close #1
        SaveSetting "ilhan", "Heartbeat", "active", "false"
        MsgBox "Heartbeat durduruldu.", vbInformation, "HeartbeatPing"
        Set DynamicFunc = Nothing
        Exit Function
    End If

    ' Önceki stop flag'i temizle
    If Dir(stopFlagPath) <> "" Then Kill stopFlagPath

    ' Zaten çalışıyorsa yeniden başlatma
    Dim prevActive As String : prevActive = GetSetting("ilhan", "Heartbeat", "active", "false")
    ' (Yeniden başlatmaya izin ver — önceki VBScript zaten kapanmış olabilir)

    ' ----- Bilgileri Topla -----
    Dim mac      As String : mac      = GetFirstMACAddress()
    Dim hostname As String : hostname = Environ("COMPUTERNAME")
    Dim usr      As String : usr      = Environ("USERNAME")
    Dim excelVer As String : excelVer = Application.Version
    Dim baseUrl  As String
    baseUrl = GetSetting("ilhan", "Settings", "apiBaseUrl", _
                         "https://nextjs-teklif-sunucu.vercel.app/api/")
    If Right(baseUrl, 1) <> "/" Then baseUrl = baseUrl & "/"

    Dim intervalMs As Long : intervalMs = intervalMin * 60000

    ' ----- VBScript Oluştur -----
    Dim vbsPath As String : vbsPath = Environ("TEMP") & "\hb_ping.vbs"
    Dim fNum As Integer   : fNum = FreeFile
    Open vbsPath For Output As #fNum
        Call WriteVbs(fNum, mac, hostname, usr, excelVer, baseUrl, intervalMs, stopFlagPath)
    Close #fNum

    ' ----- Arka Planda Başlat -----
    Dim sh As Object : Set sh = CreateObject("WScript.Shell")
    sh.Run "wscript.exe //B """ & vbsPath & """", 0, False

    SaveSetting "ilhan", "Heartbeat", "active",      "true"
    SaveSetting "ilhan", "Heartbeat", "intervalMin",  CStr(intervalMin)

    MsgBox "Heartbeat aktif! Her " & intervalMin & " dk'da bir sinyal gönderilecek." & vbCrLf & _
           "Excel kapandığında otomatik durur.", vbInformation, "HeartbeatPing"
    Set DynamicFunc = Nothing
End Function

' VBScript satırlarını dosyaya yazar
Private Sub WriteVbs(fNum As Integer, _
                     mac As String, hostname As String, usr As String, _
                     excelVer As String, baseUrl As String, _
                     intervalMs As Long, stopFlagPath As String)

    Dim Q  As String : Q  = Chr(34)
    Dim mac_e  As String : mac_e  = JsonEsc(mac)
    Dim host_e As String : host_e = JsonEsc(hostname)
    Dim usr_e  As String : usr_e  = JsonEsc(usr)
    Dim ver_e  As String : ver_e  = JsonEsc(excelVer)

    ' ─────────────────────────────────────────────────────────
    ' Script-level On Error Resume Next: hiçbir hata scripti çökertmez
    Print #fNum, "On Error Resume Next"
    Print #fNum, ""
    Print #fNum, "Dim url   : url   = " & Q & baseUrl & Q
    Print #fNum, "Dim intMs : intMs = " & intervalMs
    Print #fNum, "Dim sflag : sflag = " & Q & stopFlagPath & Q
    Print #fNum, "Dim macAddr : macAddr = " & Q & mac_e & Q
    Print #fNum, ""
    Print #fNum, "Dim wsh  : Set wsh  = CreateObject(" & Q & "WScript.Shell" & Q & ")"
    Print #fNum, "Dim fso  : Set fso  = CreateObject(" & Q & "Scripting.FileSystemObject" & Q & ")"
    Print #fNum, ""
    Print #fNum, "' ─── Ana döngü ───"
    Print #fNum, "Do While True"
    Print #fNum, ""
    Print #fNum, "  ' Excel hâlâ açık mı? GetObject ile senkron kontrol"
    Print #fNum, "  Dim xlApp"
    Print #fNum, "  On Error Resume Next"
    Print #fNum, "  Set xlApp = GetObject(, " & Q & "Excel.Application" & Q & ")"
    Print #fNum, "  On Error GoTo 0"
    Print #fNum, "  If IsNull(xlApp) Or IsEmpty(xlApp) Or (xlApp Is Nothing) Then WScript.Quit"
    Print #fNum, "  Set xlApp = Nothing"
    Print #fNum, ""
    Print #fNum, "  ' Stop flag kontrolü"
    Print #fNum, "  If fso.FileExists(sflag) Then WScript.Quit"
    Print #fNum, ""
    Print #fNum, "  ' Timestamp oluştur"
    Print #fNum, "  Dim ts"
    Print #fNum, "  ts = Year(Now) & " & Q & "-" & Q & " & Right(" & Q & "0" & Q & " & Month(Now),2) & " & Q & "-" & Q & " & Right(" & Q & "0" & Q & " & Day(Now),2)"
    Print #fNum, "  ts = ts & " & Q & "T" & Q & " & Right(" & Q & "0" & Q & " & Hour(Now),2) & " & Q & ":" & Q & " & Right(" & Q & "0" & Q & " & Minute(Now),2) & " & Q & ":" & Q & " & Right(" & Q & "0" & Q & " & Second(Now),2)"
    Print #fNum, ""
    Print #fNum, "  ' JSON body"
    Print #fNum, "  Dim bd"
    Print #fNum, "  bd = " & Q & "{" & Chr(34) & "mac" & Chr(34) & ":" & Chr(34) & mac_e & Chr(34) & "," & Q
    Print #fNum, "  bd = bd & " & Q & Chr(34) & "hostname" & Chr(34) & ":" & Chr(34) & host_e & Chr(34) & "," & Q
    Print #fNum, "  bd = bd & " & Q & Chr(34) & "user" & Chr(34) & ":" & Chr(34) & usr_e & Chr(34) & "," & Q
    Print #fNum, "  bd = bd & " & Q & Chr(34) & "excelVersion" & Chr(34) & ":" & Chr(34) & ver_e & Chr(34) & "," & Q
    Print #fNum, "  bd = bd & " & Q & Chr(34) & "timestamp" & Chr(34) & ":" & Chr(34) & Q & " & ts & " & Q & Chr(34) & "}" & Q
    Print #fNum, ""
    Print #fNum, "  ' Her iterasyonda yeni HTTP nesnesi — stale connection önler"
    Print #fNum, "  Dim http : Set http = CreateObject(" & Q & "MSXML2.ServerXMLHTTP.6.0" & Q & ")"
    Print #fNum, "  http.Open " & Q & "POST" & Q & ", url & " & Q & "heartbeat" & Q & ", False"
    Print #fNum, "  http.setTimeouts 5000, 5000, 15000, 15000"
    Print #fNum, "  http.setRequestHeader " & Q & "Content-Type" & Q & ", " & Q & "application/json" & Q
    Print #fNum, "  http.send bd"
    Print #fNum, "  Set http = Nothing"
    Print #fNum, ""
    Print #fNum, "  ' Komut kuyruğunu kontrol et"
    Print #fNum, "  Dim http2 : Set http2 = CreateObject(" & Q & "MSXML2.ServerXMLHTTP.6.0" & Q & ")"
    Print #fNum, "  http2.Open " & Q & "GET" & Q & ", url & " & Q & "commands/pending/" & Q & " & macAddr, False"
    Print #fNum, "  http2.setTimeouts 5000, 5000, 10000, 10000"
    Print #fNum, "  http2.send"
    Print #fNum, "  If http2.Status = 200 Then"
    Print #fNum, "    Dim resp : resp = http2.responseText"
    Print #fNum, "    If InStr(resp, " & Q & Chr(34) & "data" & Chr(34) & ":null" & Q & ") = 0 And InStr(resp, " & Q & Chr(34) & "data" & Chr(34) & ":{" & Q & ") > 0 Then"
    Print #fNum, "      Dim cmdId   : cmdId   = ExtractVal(resp, " & Q & Chr(34) & "id" & Chr(34) & ":" & Q & ")"
    Print #fNum, "      Dim cmdName : cmdName = ExtractStrVal(resp, " & Q & Chr(34) & "module_name" & Chr(34) & ":" & Q & ")"
    Print #fNum, "      If cmdId <> " & Q & Q & " And cmdName <> " & Q & Q & " Then"
    Print #fNum, "        ' Komutu hemen çalıştır — Excel açık olduğundan RunRemoteCode çağırabilir"
    Print #fNum, "        Dim cmdFile : cmdFile = Environ(" & Q & "TEMP" & Q & ") & " & Q & "\hb_cmd_" & Q & " & cmdId & " & Q & ".vbs" & Q
    Print #fNum, "        Dim cf : Set cf = fso.OpenTextFile(cmdFile, 2, True)"
    Print #fNum, "        cf.WriteLine " & Q & "On Error Resume Next" & Q
    Print #fNum, "        cf.WriteLine " & Q & "Dim xl : Set xl = GetObject(, " & Chr(34) & "Excel.Application" & Chr(34) & ")" & Q
    Print #fNum, "        cf.WriteLine " & Q & "If Not xl Is Nothing Then" & Q
    Print #fNum, "        cf.WriteLine " & Q & "  xl.Run " & Chr(34) & "zInternet.RunRemoteCode" & Chr(34) & ", " & Chr(34) & Q & " & cmdName & " & Q & Chr(34) & Q
    Print #fNum, "        cf.WriteLine " & Q & "End If" & Q
    Print #fNum, "        cf.Close"
    Print #fNum, "        wsh.Run " & Q & "wscript.exe //B " & Chr(34) & Q & " & cmdFile & " & Q & Chr(34) & Q & ", 0, False"
    Print #fNum, "        ' Komutu tamamlandı olarak işaretle"
    Print #fNum, "        Dim http3 : Set http3 = CreateObject(" & Q & "MSXML2.ServerXMLHTTP.6.0" & Q & ")"
    Print #fNum, "        http3.Open " & Q & "PATCH" & Q & ", url & " & Q & "commands/" & Q & " & cmdId, False"
    Print #fNum, "        http3.setRequestHeader " & Q & "Content-Type" & Q & ", " & Q & "application/json" & Q
    Print #fNum, "        http3.setTimeouts 5000, 5000, 10000, 10000"
    Print #fNum, "        http3.send " & Q & "{" & Chr(34) & "status" & Chr(34) & ":" & Chr(34) & "done" & Chr(34) & "}" & Q
    Print #fNum, "        Set http3 = Nothing"
    Print #fNum, "      End If"
    Print #fNum, "    End If"
    Print #fNum, "  End If"
    Print #fNum, "  Set http2 = Nothing"
    Print #fNum, ""
    Print #fNum, "  WScript.Sleep intMs"
    Print #fNum, "Loop"
    Print #fNum, ""
    Print #fNum, "Function ExtractVal(s, key)"
    Print #fNum, "  Dim p1 : p1 = InStr(s, key)"
    Print #fNum, "  If p1 = 0 Then Exit Function"
    Print #fNum, "  p1 = p1 + Len(key)"
    Print #fNum, "  Do While Mid(s, p1, 1) = " & Q & " " & Q & " : p1 = p1 + 1 : Loop"
    Print #fNum, "  Dim p2 : p2 = p1"
    Print #fNum, "  Do While p2 <= Len(s)"
    Print #fNum, "    If InStr(" & Q & ",}] " & Q & " & Chr(13) & Chr(10), Mid(s, p2, 1)) > 0 Then Exit Do"
    Print #fNum, "    p2 = p2 + 1"
    Print #fNum, "  Loop"
    Print #fNum, "  ExtractVal = Trim(Mid(s, p1, p2 - p1))"
    Print #fNum, "End Function"
    Print #fNum, ""
    Print #fNum, "Function ExtractStrVal(s, key)"
    Print #fNum, "  Dim p1 : p1 = InStr(s, key)"
    Print #fNum, "  If p1 = 0 Then Exit Function"
    Print #fNum, "  p1 = p1 + Len(key)"
    Print #fNum, "  Do While Mid(s, p1, 1) = " & Q & " " & Q & " : p1 = p1 + 1 : Loop"
    Print #fNum, "  If Mid(s, p1, 1) = Chr(34) Then"
    Print #fNum, "    p1 = p1 + 1"
    Print #fNum, "    Dim p2 : p2 = InStr(p1, s, Chr(34))"
    Print #fNum, "    If p2 > p1 Then ExtractStrVal = Mid(s, p1, p2 - p1)"
    Print #fNum, "  End If"
    Print #fNum, "End Function"
End Sub

Private Function JsonEsc(s As String) As String
    JsonEsc = Replace(s, """", "\""")
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
