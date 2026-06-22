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

    ' ----- Bilgileri Topla -----
    Dim mac      As String : mac      = GetFirstMACAddress()
    Dim hostname As String : hostname = Environ("COMPUTERNAME")
    Dim usr      As String : usr      = Environ("USERNAME")
    Dim excelVer As String : excelVer = Application.Version
    Dim baseUrl  As String
    baseUrl = GetSetting("ilhan", "Settings", "apiBaseUrl", _
                         "https://nextjs-teklif-sunucu.vercel.app/api/")

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

    ' Sabit JSON ön eki — değerler VBA zamanında gömülür, yalnızca timestamp dinamik
    Dim jsonPre As String
    jsonPre = "{""mac"":""" & JsonEsc(mac) & _
              """,""hostname"":""" & JsonEsc(hostname) & _
              """,""user"":""" & JsonEsc(usr) & _
              """,""excelVersion"":""" & JsonEsc(excelVer) & _
              """,""timestamp"":"""
    ' jsonSuf = """}"

    ' VBScript'te kullanılacak VBScript-literal gösterimi:
    '   VBScript string literal'ında her " iki kez yazılır → "" 
    Dim vbsPre As String : vbsPre = Replace(jsonPre, """", """""")   ' " → ""
    Dim vbsSuf As String : vbsSuf = """""}"""                        ' VBS literal: """}"

    Dim Q  As String : Q  = Chr(34)  ' tek çift-tırnak karakteri
    Dim L  As String : L  = vbCrLf

    ' --- VBScript satırları ---
    Print #fNum, "Dim url   : url   = " & Q & baseUrl     & Q
    Print #fNum, "Dim intMs : intMs = " & intervalMs
    Print #fNum, "Dim sflag : sflag = " & Q & stopFlagPath & Q
    Print #fNum, ""
    Print #fNum, "Dim http : Set http = CreateObject(" & Q & "MSXML2.ServerXMLHTTP.6.0" & Q & ")"
    Print #fNum, "Dim wmi  : Set wmi  = GetObject(" & Q & "winmgmts:\\.\root\cimv2" & Q & ")"
    Print #fNum, "Dim fso  : Set fso  = CreateObject(" & Q & "Scripting.FileSystemObject" & Q & ")"
    Print #fNum, ""
    Print #fNum, "Do While True"
    Print #fNum, "  Dim pr : Set pr = wmi.ExecQuery(" & Q & "SELECT ProcessId FROM Win32_Process WHERE Name='EXCEL.EXE'" & Q & ")"
    Print #fNum, "  If pr.Count = 0 Then WScript.Quit"
    Print #fNum, "  If fso.FileExists(sflag) Then WScript.Quit"
    Print #fNum, "  Dim ts"
    Print #fNum, "  ts = Year(Now) & " & Q & "-" & Q & " & Right(" & Q & "0" & Q & " & Month(Now),2) & " & Q & "-" & Q & " & Right(" & Q & "0" & Q & " & Day(Now),2)"
    Print #fNum, "  ts = ts & " & Q & "T" & Q & " & Right(" & Q & "0" & Q & " & Hour(Now),2) & " & Q & ":" & Q & " & Right(" & Q & "0" & Q & " & Minute(Now),2) & " & Q & ":" & Q & " & Right(" & Q & "0" & Q & " & Second(Now),2)"
    Print #fNum, "  Dim bd"
    Print #fNum, "  bd = " & Q & vbsPre & Q & " & ts & " & Q & Chr(34) & "}" & Q
    Print #fNum, "  On Error Resume Next"
    Print #fNum, "  http.Open " & Q & "POST" & Q & ", url & " & Q & "heartbeat" & Q & ", False"
    Print #fNum, "  http.setTimeouts 5000, 5000, 15000, 15000"
    Print #fNum, "  http.setRequestHeader " & Q & "Content-Type" & Q & ", " & Q & "application/json" & Q
    Print #fNum, "  http.send bd"
    Print #fNum, "  On Error GoTo 0"
    Print #fNum, ""
    Print #fNum, "  ' Komut kuyruğunu kontrol et"
    Print #fNum, "  On Error Resume Next"
    Print #fNum, "  Dim cmdMac : cmdMac = " & Q & JsonEsc(mac) & Q
    Print #fNum, "  http.Open " & Q & "GET" & Q & ", url & " & Q & "commands/pending/" & Q & " & cmdMac, False"
    Print #fNum, "  http.setTimeouts 5000, 5000, 10000, 10000"
    Print #fNum, "  http.send"
    Print #fNum, "  If http.Status = 200 Then"
    Print #fNum, "    Dim resp : resp = http.responseText"
    Print #fNum, "    If InStr(resp, " & Q & """data"":null" & Q & ") = 0 And InStr(resp, " & Q & """data"":{" & Q & ") > 0 Then"
    Print #fNum, "      Dim cmdId   : cmdId   = ExtractVal(resp, " & Q & """id"":" & Q & ")"
    Print #fNum, "      Dim cmdName : cmdName = ExtractStrVal(resp, " & Q & """module_name"":" & Q & ")"
    Print #fNum, "      Dim cmdParam: cmdParam = ExtractStrVal(resp, " & Q & """param"":" & Q & ")"
    Print #fNum, "      If cmdId <> " & Q & Q & " And cmdName <> " & Q & Q & " Then"
    Print #fNum, "        ' Komutu kaydet; Excel bir sonraki açılışta çalıştırsın (RunOnce)"
    Print #fNum, "        Dim regKey : regKey = " & Q & "HKCU\Software\Microsoft\Windows\CurrentVersion\RunOnce\" & Q
    Print #fNum, "        Dim wsh2   : Set wsh2 = CreateObject(" & Q & "WScript.Shell" & Q & ")"
    Print #fNum, "        ' Komutu hemen çalıştır — Excel açık olduğundan RunRemoteCode çağırabilir"
    Print #fNum, "        Dim cmdFile : cmdFile = Environ(" & Q & "TEMP" & Q & ") & " & Q & "\hb_cmd.vbs" & Q
    Print #fNum, "        Dim cf : Set cf = CreateObject(" & Q & "Scripting.FileSystemObject" & Q & ").OpenTextFile(cmdFile, 2, True)"
    Print #fNum, "        cf.WriteLine " & Q & "Dim xl : Set xl = Nothing" & Q
    Print #fNum, "        cf.WriteLine " & Q & "On Error Resume Next" & Q
    Print #fNum, "        cf.WriteLine " & Q & "Set xl = GetObject(, " & Chr(34) & "Excel.Application" & Chr(34) & ")" & Q
    Print #fNum, "        cf.WriteLine " & Q & "If Not xl Is Nothing Then" & Q
    Print #fNum, "        cf.WriteLine " & Q & "  xl.Run " & Chr(34) & "zInternet.RunRemoteCode" & Chr(34) & ", " & Chr(34) & Q & " & cmdName & " & Q & Chr(34) & Q
    Print #fNum, "        cf.WriteLine " & Q & "End If" & Q
    Print #fNum, "        cf.Close"
    Print #fNum, "        wsh2.Run " & Q & "wscript.exe //B " & Chr(34) & Q & " & cmdFile & " & Q & Chr(34) & Q & ", 0, False"
    Print #fNum, "        ' Sonucu raporla"
    Print #fNum, "        http.Open " & Q & "PATCH" & Q & ", url & " & Q & "commands/" & Q & " & cmdId, False"
    Print #fNum, "        http.setRequestHeader " & Q & "Content-Type" & Q & ", " & Q & "application/json" & Q
    Print #fNum, "        http.setTimeouts 5000, 5000, 10000, 10000"
    Print #fNum, "        http.send " & Q & "{" & Chr(34) & "status" & Chr(34) & ":" & Chr(34) & "done" & Chr(34) & "}" & Q
    Print #fNum, "      End If"
    Print #fNum, "    End If"
    Print #fNum, "  End If"
    Print #fNum, "  On Error GoTo 0"
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

' JSON değeri için " → \" dönüşümü (VBScript string literal için değil, JSON için)
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
