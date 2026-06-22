Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ' param: {"intervalMin":1,"stop":false}  veya sadece "1" (dakika)
    ' Sunucuya MAC + hostname + versiyon + zaman içeren periyodik sinyal gönderir.
    Dim p As String : p = Trim(CStr(param))

    Dim stopFlag As Boolean : stopFlag = False
    Dim intervalMin As Long : intervalMin = 1

    If Left(p, 1) = "{" Then
        Dim stopStr As String : stopStr = ExtractJsonValue(p, "stop")
        If LCase(stopStr) = "true" Then stopFlag = True
        Dim intStr As String : intStr = ExtractJsonValue(p, "intervalMin")
        If Len(intStr) > 0 Then intervalMin = CLng(intStr)
    ElseIf IsNumeric(p) Then
        intervalMin = CLng(p)
    End If

    If stopFlag Then
        SaveSetting "ilhan", "Heartbeat", "active", "false"
        MsgBox "Heartbeat durduruldu.", vbInformation
        Set DynamicFunc = Nothing
        Exit Function
    End If

    ' Ayarları kaydet
    Dim baseUrl As String : baseUrl = GetSetting("ilhan", "Settings", "apiBaseUrl", "http://localhost:3000/api/")
    SaveSetting "ilhan", "Heartbeat", "active",      "true"
    SaveSetting "ilhan", "Heartbeat", "intervalMin", CStr(intervalMin)
    SaveSetting "ilhan", "Heartbeat", "baseUrl",     baseUrl

    ' İlk ping'i hemen gönder
    Call SendHeartbeat(baseUrl)

    ' Sonraki pingleri zamanla
    Application.OnTime Now + TimeSerial(0, intervalMin, 0), "HeartbeatPing_Run"

    MsgBox "Heartbeat aktif: her " & intervalMin & " dakika" & IIf(intervalMin = 1, "", "") & "da bir sinyal gönderilecek.", vbInformation
    Set DynamicFunc = Nothing
End Function

Private Sub SendHeartbeat(baseUrl As String)
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    On Error Resume Next

    Dim mac      As String : mac      = GetFirstMACAddress()
    Dim hostname As String : hostname = Environ("COMPUTERNAME")
    Dim user     As String : user     = Environ("USERNAME")
    Dim excelVer As String : excelVer = Application.Version

    Dim body As String
    body = "{""mac"":""" & mac & """," & _
           """hostname"":""" & hostname & """," & _
           """user"":""" & user & """," & _
           """excelVersion"":""" & excelVer & """," & _
           """timestamp"":""" & Format(Now, "yyyy-MM-ddTHH:mm:ss") & """}"

    http.Open "POST", baseUrl & "heartbeat", False
    http.setTimeouts 3000, 5000, 10000, 10000
    http.setRequestHeader "Content-Type", "application/json"
    http.send body

    Debug.Print "[HeartbeatPing] Status: " & http.Status & " at " & Format(Now, "HH:mm:ss")
    Set http = Nothing
    On Error GoTo 0
End Sub

' teklif.xlam içine ekle:
'
' Public Sub HeartbeatPing_Run()
'     If GetSetting("ilhan","Heartbeat","active","false") <> "true" Then Exit Sub
'     Dim baseUrl As String : baseUrl = GetSetting("ilhan","Heartbeat","baseUrl","http://localhost:3000/api/")
'     Dim intMin  As Long   : intMin  = CLng(GetSetting("ilhan","Heartbeat","intervalMin","1"))
'
'     ' Bilgileri topla ve gönder
'     Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
'     On Error Resume Next
'     Dim mac      As String : mac      = GetFirstMACAddress()
'     Dim hostname As String : hostname = Environ("COMPUTERNAME")
'     Dim user     As String : user     = Environ("USERNAME")
'     Dim excelVer As String : excelVer = Application.Version
'     Dim body As String
'     body = "{""mac"":""" & mac & """," & _
'            """hostname"":""" & hostname & """," & _
'            """user"":""" & user & """," & _
'            """excelVersion"":""" & excelVer & """," & _
'            """timestamp"":""" & Format(Now, "yyyy-MM-ddTHH:mm:ss") & """}"
'     http.Open "POST", baseUrl & "heartbeat", False
'     http.setTimeouts 3000, 5000, 10000, 10000
'     http.setRequestHeader "Content-Type", "application/json"
'     http.send body
'     Set http = Nothing
'     On Error GoTo 0
'
'     Application.OnTime Now + TimeSerial(0, intMin, 0), "HeartbeatPing_Run"
' End Sub

Private Function GetFirstMACAddress() As String
    Dim objWMI As Object, colAdapters As Object, obj As Object
    On Error GoTo Fail
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set colAdapters = objWMI.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
    GetFirstMACAddress = "UNKNOWN"
    For Each obj In colAdapters
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
