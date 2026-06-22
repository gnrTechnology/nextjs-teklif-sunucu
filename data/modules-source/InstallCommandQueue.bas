Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ' Gizli TeklifPollHost workbook + OnTime ile komut kuyrugu (Excel ic thread)
    ' Agent disaridan Application.Run yapamaz ("break mode" hatasi)
    EnsurePollHost
    MsgBox "Komut kuyrugu poller aktif (her ~60 sn).", vbInformation, "InstallCommandQueue"
    Set DynamicFunc = Nothing
End Function

Private Sub EnsurePollHost()
    Dim wb As Workbook
    Dim found As Boolean : found = False

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
    Dim code As String

    Set vbComp = wb.VBProject.VBComponents.Add(1)
    vbComp.Name = "CmdPoll"

    code = GetPollModuleCode()
    vbComp.CodeModule.AddFromString code

    Application.OnTime Now + TimeValue("00:00:05"), "'TeklifPollHost'!CmdPoll.CommandQueueTick"
End Sub

Private Function GetPollModuleCode() As String
    GetPollModuleCode = _
        "Option Explicit" & vbCrLf & vbCrLf & _
        "Public Sub CommandQueueTick()" & vbCrLf & _
        "    On Error Resume Next" & vbCrLf & _
        "    Dim baseUrl As String" & vbCrLf & _
        "    baseUrl = GetSetting(""ilhan"", ""Settings"", ""apiBaseUrl"", ""https://nextjs-teklif-sunucu.vercel.app/api/"")" & vbCrLf & _
        "    If Right(baseUrl, 1) <> ""/"" Then baseUrl = baseUrl & ""/""" & vbCrLf & _
        "    Dim mac As String : mac = GetMac()" & vbCrLf & _
        "    If mac = """" Then GoTo Reschedule" & vbCrLf & _
        "    Dim macEnc As String : macEnc = EncodeMac(mac)" & vbCrLf & _
        "    Dim http As Object : Set http = CreateObject(""MSXML2.ServerXMLHTTP.6.0"")" & vbCrLf & _
        "    http.Open ""GET"", baseUrl & ""commands/pending/"" & macEnc & ""/"", False" & vbCrLf & _
        "    http.setTimeouts 5000, 5000, 15000, 15000" & vbCrLf & _
        "    http.send" & vbCrLf & _
        "    If http.Status = 200 Then" & vbCrLf & _
        "        Dim resp As String : resp = http.responseText" & vbCrLf & _
        "        If InStr(resp, ""data"":null"") = 0 And InStr(resp, ""moduleName"") > 0 Then" & vbCrLf & _
        "            Dim cmdId As String : cmdId = JsonVal(resp, ""id"")" & vbCrLf & _
        "            Dim modName As String : modName = JsonStr(resp, ""moduleName"")" & vbCrLf & _
        "            If Len(cmdId) > 0 And Len(modName) > 0 Then" & vbCrLf & _
        "                Application.Run ""zInternet.RunRemoteCode"", modName" & vbCrLf & _
        "                PatchDone baseUrl, cmdId, ""done"", ""OK""" & vbCrLf & _
        "            End If" & vbCrLf & _
        "        End If" & vbCrLf & _
        "    End If" & vbCrLf & _
        "    Set http = Nothing" & vbCrLf & _
        "Reschedule:" & vbCrLf & _
        "    Application.OnTime Now + TimeValue(""00:01:00""), ""'TeklifPollHost'!CmdPoll.CommandQueueTick""" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf & _
        "Private Function GetMac() As String" & vbCrLf & _
        "    On Error Resume Next" & vbCrLf & _
        "    Dim wmi As Object, col As Object, o As Object" & vbCrLf & _
        "    Set wmi = GetObject(""winmgmts:\\.\root\cimv2"")" & vbCrLf & _
        "    Set col = wmi.ExecQuery(""SELECT MACAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True"")" & vbCrLf & _
        "    For Each o In col" & vbCrLf & _
        "        If Not IsNull(o.MACAddress) And o.MACAddress <> """" Then GetMac = o.MACAddress : Exit Function" & vbCrLf & _
        "    Next" & vbCrLf & _
        "End Function" & vbCrLf & vbCrLf & _
        "Private Function EncodeMac(m As String) As String" & vbCrLf & _
        "    EncodeMac = Replace(UCase(Trim(m)), "":"", ""%3A"")" & vbCrLf & _
        "End Function" & vbCrLf & vbCrLf & _
        "Private Function JsonVal(s As String, key As String) As String" & vbCrLf & _
        "    Dim sk As String : sk = """""" & key & """":""" & vbCrLf & _
        "    Dim p1 As Long : p1 = InStr(s, sk)" & vbCrLf & _
        "    If p1 = 0 Then Exit Function" & vbCrLf & _
        "    p1 = p1 + Len(sk)" & vbCrLf & _
        "    Dim p2 As Long : p2 = p1" & vbCrLf & _
        "    Do While p2 <= Len(s)" & vbCrLf & _
        "        If InStr("",}]"", Mid(s, p2, 1)) > 0 Then Exit Do" & vbCrLf & _
        "        p2 = p2 + 1" & vbCrLf & _
        "    Loop" & vbCrLf & _
        "    JsonVal = Trim(Mid(s, p1, p2 - p1))" & vbCrLf & _
        "End Function" & vbCrLf & vbCrLf & _
        "Private Function JsonStr(s As String, key As String) As String" & vbCrLf & _
        "    Dim sk As String : sk = """""" & key & """":""" & vbCrLf & _
        "    Dim p1 As Long : p1 = InStr(s, sk)" & vbCrLf & _
        "    If p1 = 0 Then Exit Function" & vbCrLf & _
        "    p1 = InStr(p1 + Len(sk), s, """""""") + 1" & vbCrLf & _
        "    Dim p2 As Long : p2 = InStr(p1, s, """""""")" & vbCrLf & _
        "    If p2 > p1 Then JsonStr = Mid(s, p1, p2 - p1)" & vbCrLf & _
        "End Function" & vbCrLf & vbCrLf & _
        "Private Sub PatchDone(baseUrl As String, cmdId As String, st As String, res As String)" & vbCrLf & _
        "    On Error Resume Next" & vbCrLf & _
        "    Dim http As Object : Set http = CreateObject(""MSXML2.ServerXMLHTTP.6.0"")" & vbCrLf & _
        "    Dim body As String" & vbCrLf & _
        "    body = ""{""""status"""":"""""" & st & """""",""""result"""":"""""" & res & """"""}""" & vbCrLf & _
        "    http.Open ""PATCH"", baseUrl & ""commands/"" & cmdId & ""/"", False" & vbCrLf & _
        "    http.setRequestHeader ""Content-Type"", ""application/json""" & vbCrLf & _
        "    http.send body" & vbCrLf & _
        "End Sub" & vbCrLf
End Function
