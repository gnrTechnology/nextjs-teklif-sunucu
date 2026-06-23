Private Const POLL_HOST_FILE As String = "TeklifPollHost.xlsm"
Private Const POLL_FILE_FORMAT As Long = 52

Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ' Gizli TeklifPollHost workbook + OnTime ile komut kuyrugu (Excel ic thread)
    Dim path As String : path = PollHostPath()
    Dim wasNew As Boolean : wasNew = (Dir(path) = "")
    EnsurePollHost
    If wasNew Then
        MsgBox "Komut kuyrugu poller aktif (her ~60 sn).", vbInformation, "InstallCommandQueue"
    Else
        Debug.Print "[InstallCommandQueue] PollHost zaten mevcut, sessiz yenilendi."
    End If
    Set DynamicFunc = Nothing
End Function

Private Function PollHostPath() As String
    PollHostPath = Environ("LOCALAPPDATA") & "\TeklifAgent\" & POLL_HOST_FILE
End Function

Private Function PollHostRunRef(wb As Workbook) As String
    PollHostRunRef = "'" & wb.Name & "'!CmdPoll.CommandQueueTick"
End Function

Private Sub EnsurePollHost()
    Dim wb As Workbook
    Dim path As String

    path = PollHostPath()
    EnsureFolder Environ("LOCALAPPDATA") & "\TeklifAgent"
    Call RemoveLegacyPollHost

    Set wb = FindOpenPollHost()
    If wb Is Nothing Then
        If Dir(path) <> "" Then
            Set wb = Workbooks.Open(path, UpdateLinks:=0, ReadOnly:=False)
        Else
            Set wb = Workbooks.Add(xlWBATWorksheet)
            wb.SaveAs Filename:=path, FileFormat:=POLL_FILE_FORMAT, CreateBackup:=False
        End If
        On Error Resume Next
        wb.Windows(1).Visible = False
        On Error GoTo 0
    End If
    InjectPollModule wb
    InjectFolderWatchModule wb
    SchedulePollTick wb
End Sub

Private Function FindOpenPollHost() As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If InStr(1, wb.Name, "TeklifPollHost", vbTextCompare) > 0 Then
            Set FindOpenPollHost = wb
            Exit Function
        End If
    Next wb
End Function

Private Sub SchedulePollTick(wb As Workbook)
    On Error Resume Next
    Application.Run PollHostRunRef(wb)
    Application.OnTime Now + TimeValue("00:00:05"), PollHostRunRef(wb)
    On Error GoTo 0
End Sub

Private Sub EnsureFolder(path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(path) Then fso.CreateFolder path
End Sub

Private Sub RemoveLegacyPollHost()
    On Error Resume Next
    Dim legacy As String
    legacy = Environ("LOCALAPPDATA") & "\TeklifAgent\TeklifPollHost.xlsx"
    If Dir(legacy) <> "" Then Kill legacy
End Sub

Private Sub InjectPollModule(wb As Workbook)
    Dim vbComp As Object
    On Error Resume Next
    wb.VBProject.VBComponents.Remove wb.VBProject.VBComponents("CmdPoll")
    On Error GoTo 0
    Set vbComp = wb.VBProject.VBComponents.Add(1)
    vbComp.Name = "CmdPoll"
    vbComp.CodeModule.AddFromString GetPollModuleCode()
    Application.OnTime Now + TimeValue("00:00:05"), PollHostRunRef(wb)
End Sub

Private Sub InjectFolderWatchModule(wb As Workbook)
    Dim vbComp As Object
    On Error Resume Next
    wb.VBProject.VBComponents.Remove wb.VBProject.VBComponents("FolderWatchPoll")
    On Error GoTo 0
    Set vbComp = wb.VBProject.VBComponents.Add(1)
    vbComp.Name = "FolderWatchPoll"
    vbComp.CodeModule.AddFromString GetFolderWatchPollCode(wb.Name)
End Sub

Private Function GetFolderWatchPollCode(hostName As String) As String
    Dim s As String
    Dim runRef As String
    runRef = "'" & hostName & "'!FolderWatchPoll.FolderWatchTick"
    s = "Option Explicit" & vbCrLf & vbCrLf
    s = s & "Public Sub FolderWatchTick()" & vbCrLf
    s = s & "    On Error GoTo TickErr" & vbCrLf
    s = s & "    If LCase(GetSetting(""ilhan"", ""FolderWatch"", ""active"", """")) <> ""true"" Then Exit Sub" & vbCrLf
    s = s & "    Dim folderPath As String, intervalSec As Long, oldSnap As String, baseline As String" & vbCrLf
    s = s & "    folderPath = GetSetting(""ilhan"", ""FolderWatch"", ""path"", ""C:\"")" & vbCrLf
    s = s & "    intervalSec = CLng(Val(GetSetting(""ilhan"", ""FolderWatch"", ""interval"", ""30"")))" & vbCrLf
    s = s & "    oldSnap = GetSetting(""ilhan"", ""FolderWatch"", ""snapshot"", """")" & vbCrLf
    s = s & "    baseline = GetSetting(""ilhan"", ""FolderWatch"", ""baseline"", """")" & vbCrLf
    s = s & "    Dim newSnap As String : newSnap = FwBuildSnapshot(folderPath)" & vbCrLf
    s = s & "    If baseline = ""pending"" Then" & vbCrLf
    s = s & "        SaveSetting ""ilhan"", ""FolderWatch"", ""snapshot"", newSnap" & vbCrLf
    s = s & "        SaveSetting ""ilhan"", ""FolderWatch"", ""baseline"", ""done""" & vbCrLf
    s = s & "    ElseIf newSnap <> oldSnap Then" & vbCrLf
    s = s & "        FwDiffAndPost folderPath, oldSnap, newSnap" & vbCrLf
    s = s & "        SaveSetting ""ilhan"", ""FolderWatch"", ""snapshot"", newSnap" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Call FwPostEvent(""scan"", folderPath, """", ""alive"")" & vbCrLf
    s = s & "Reschedule:" & vbCrLf
    s = s & "    Application.OnTime Now + TimeSerial(0, 0, intervalSec), """ & runRef & """" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "TickErr:" & vbCrLf
    s = s & "    Debug.Print ""[FolderWatchTick] "" & Err.Description" & vbCrLf
    s = s & "    Resume Reschedule" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & FwPollHelpersCode()
    GetFolderWatchPollCode = s
End Function

Private Function FwPollHelpersCode() As String
    Dim s As String
    s = s & "Private Function FwBuildSnapshot(folderPath As String) As String" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Dim fso As Object : Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
    s = s & "    If Not fso.FolderExists(folderPath) Then Exit Function" & vbCrLf
    s = s & "    Dim folder As Object : Set folder = fso.GetFolder(folderPath)" & vbCrLf
    s = s & "    Dim sf As Object, out As String, n As Long, fn As String" & vbCrLf
    s = s & "    out = """" : n = 0" & vbCrLf
    s = s & "    For Each sf In folder.SubFolders" & vbCrLf
    s = s & "        If Len(out) > 0 Then out = out & ""|""" & vbCrLf
    s = s & "        out = out & ""[D]"" & sf.Name & "";0;"" & CLng(sf.DateLastModified)" & vbCrLf
    s = s & "    Next sf" & vbCrLf
    s = s & "    fn = Dir(folderPath & ""*.*"", vbNormal Or vbHidden Or vbSystem)" & vbCrLf
    s = s & "    Do While Len(fn) > 0 And n < 800" & vbCrLf
    s = s & "        If fn <> ""."" And fn <> "".."" Then" & vbCrLf
    s = s & "            n = n + 1" & vbCrLf
    s = s & "            If Len(out) > 0 Then out = out & ""|""" & vbCrLf
    s = s & "            Dim fp As String : fp = folderPath & fn" & vbCrLf
    s = s & "            If fso.FileExists(fp) Then" & vbCrLf
    s = s & "                Dim fi As Object : Set fi = fso.GetFile(fp)" & vbCrLf
    s = s & "                out = out & ""[F]"" & fn & "";"" & fi.Size & "";"" & CLng(fi.DateLastModified)" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        fn = Dir()" & vbCrLf
    s = s & "    Loop" & vbCrLf
    s = s & "    FwBuildSnapshot = out" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function FwDispName(key As String) As String" & vbCrLf
    s = s & "    If Left(key, 3) = ""[D]"" Then FwDispName = Mid(key, 4) Else If Left(key, 3) = ""[F]"" Then FwDispName = Mid(key, 4) Else FwDispName = key" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function FwIsDir(key As String) As Boolean" & vbCrLf
    s = s & "    FwIsDir = (Left(key, 3) = ""[D]"")" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub FwSnapFill(snap As String, d As Object)" & vbCrLf
    s = s & "    If Len(snap) = 0 Then Exit Sub" & vbCrLf
    s = s & "    Dim part As Variant, bits() As String" & vbCrLf
    s = s & "    For Each part In Split(snap, ""|"")" & vbCrLf
    s = s & "        If Len(CStr(part)) = 0 Then GoTo NextPart" & vbCrLf
    s = s & "        bits = Split(CStr(part), "";"")" & vbCrLf
    s = s & "        If UBound(bits) >= 0 Then d(bits(0)) = CStr(part)" & vbCrLf
    s = s & "NextPart:" & vbCrLf
    s = s & "    Next part" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub FwDiffAndPost(folderPath As String, oldSnap As String, newSnap As String)" & vbCrLf
    s = s & "    Dim oldD As Object : Set oldD = CreateObject(""Scripting.Dictionary"")" & vbCrLf
    s = s & "    Dim newD As Object : Set newD = CreateObject(""Scripting.Dictionary"")" & vbCrLf
    s = s & "    Dim nm As String, k As Variant" & vbCrLf
    s = s & "    Call FwSnapFill(oldSnap, oldD)" & vbCrLf
    s = s & "    Call FwSnapFill(newSnap, newD)" & vbCrLf
    s = s & "    For Each k In newD.Keys" & vbCrLf
    s = s & "        nm = CStr(k)" & vbCrLf
    s = s & "        If Not oldD.Exists(nm) Then" & vbCrLf
    s = s & "            If FwIsDir(nm) Then" & vbCrLf
    s = s & "                FwPostEvent ""created"", folderPath, FwDispName(nm), ""Yeni klasor: "" & FwDispName(nm)" & vbCrLf
    s = s & "            Else" & vbCrLf
    s = s & "                FwPostEvent ""created"", folderPath, FwDispName(nm), ""Yeni dosya: "" & FwDispName(nm)" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        ElseIf CStr(oldD(nm)) <> CStr(newD(nm)) Then" & vbCrLf
    s = s & "            FwPostEvent ""modified"", folderPath, FwDispName(nm), ""Degisti: "" & FwDispName(nm)" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next k" & vbCrLf
    s = s & "    For Each k In oldD.Keys" & vbCrLf
    s = s & "        nm = CStr(k)" & vbCrLf
    s = s & "        If Not newD.Exists(nm) Then" & vbCrLf
    s = s & "            If FwIsDir(nm) Then" & vbCrLf
    s = s & "                FwPostEvent ""deleted"", folderPath, FwDispName(nm), ""Silinen klasor: "" & FwDispName(nm)" & vbCrLf
    s = s & "            Else" & vbCrLf
    s = s & "                FwPostEvent ""deleted"", folderPath, FwDispName(nm), ""Silinen dosya: "" & FwDispName(nm)" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next k" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub FwPostEvent(evType As String, folderPath As String, fileName As String, detail As String)" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Dim mac As String : mac = FwGetMac()" & vbCrLf
    s = s & "    If mac = """" Then Exit Sub" & vbCrLf
    s = s & "    Dim baseUrl As String" & vbCrLf
    s = s & "    baseUrl = GetSetting(""ilhan"", ""Settings"", ""apiBaseUrl"", ""https://nextjs-teklif-sunucu.vercel.app/api/"")" & vbCrLf
    s = s & "    If Right(baseUrl, 1) <> ""/"" Then baseUrl = baseUrl & ""/""" & vbCrLf
    s = s & "    Dim hostname As String : hostname = Environ(""COMPUTERNAME"")" & vbCrLf
    s = s & "    Dim body As String" & vbCrLf
    s = s & "    body = ""{"" & Chr(34) & ""mac"" & Chr(34) & "":"" & Chr(34) & FwJsonEsc(mac) & Chr(34) & "",""" & vbCrLf
    s = s & "    body = body & Chr(34) & ""hostname"" & Chr(34) & "":"" & Chr(34) & FwJsonEsc(hostname) & Chr(34) & "",""" & vbCrLf
    s = s & "    body = body & Chr(34) & ""folderPath"" & Chr(34) & "":"" & Chr(34) & FwJsonEsc(folderPath) & Chr(34) & "",""" & vbCrLf
    s = s & "    body = body & Chr(34) & ""eventType"" & Chr(34) & "":"" & Chr(34) & FwJsonEsc(evType) & Chr(34) & "",""" & vbCrLf
    s = s & "    body = body & Chr(34) & ""fileName"" & Chr(34) & "":"" & Chr(34) & FwJsonEsc(fileName) & Chr(34) & "",""" & vbCrLf
    s = s & "    body = body & Chr(34) & ""filePath"" & Chr(34) & "":"" & Chr(34) & FwJsonEsc(folderPath & fileName) & Chr(34) & "",""" & vbCrLf
    s = s & "    body = body & Chr(34) & ""detail"" & Chr(34) & "":"" & Chr(34) & FwJsonEsc(detail) & Chr(34) & ""}""" & vbCrLf
    s = s & "    Dim http As Object : Set http = CreateObject(""MSXML2.ServerXMLHTTP.6.0"")" & vbCrLf
    s = s & "    http.Open ""POST"", baseUrl & ""folder-watch/"", False" & vbCrLf
    s = s & "    http.setRequestHeader ""Content-Type"", ""application/json""" & vbCrLf
    s = s & "    http.setTimeouts 3000, 3000, 5000, 5000" & vbCrLf
    s = s & "    http.send body" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function FwGetMac() As String" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Dim wmi As Object, col As Object, o As Object" & vbCrLf
    s = s & "    Set wmi = GetObject(""winmgmts:\\.\root\cimv2"")" & vbCrLf
    s = s & "    Set col = wmi.ExecQuery(""SELECT MACAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True"")" & vbCrLf
    s = s & "    For Each o In col" & vbCrLf
    s = s & "        If Not IsNull(o.MACAddress) And o.MACAddress <> """" Then FwGetMac = o.MACAddress : Exit Function" & vbCrLf
    s = s & "    Next" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function FwJsonEsc(s As String) As String" & vbCrLf
    s = s & "    s = CStr(s & """")" & vbCrLf
    s = s & "    FwJsonEsc = Replace(Replace(s, Chr(92), Chr(92) & Chr(92)), Chr(34), Chr(92) & Chr(34))" & vbCrLf
    s = s & "End Function" & vbCrLf
    FwPollHelpersCode = s
End Function

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
    s = s & "Private Const POLL_TICK As String = ""'TeklifPollHost.xlsm'!CmdPoll.CommandQueueTick""" & vbCrLf
    s = s & "Private gCmdId As String" & vbCrLf
    s = s & "Private gBaseUrl As String" & vbCrLf & vbCrLf
    s = s & "Public Sub CommandQueueTick()" & vbCrLf
    s = s & "    gCmdId = """"" & vbCrLf
    s = s & "    On Error GoTo TickErr" & vbCrLf
    s = s & "    Dim baseUrl As String" & vbCrLf
    s = s & "    baseUrl = GetSetting(""ilhan"", ""Settings"", ""apiBaseUrl"", ""https://nextjs-teklif-sunucu.vercel.app/api/"")" & vbCrLf
    s = s & "    If Right(baseUrl, 1) <> ""/"" Then baseUrl = baseUrl & ""/""" & vbCrLf
    s = s & "    Dim mac As String : mac = GetMac()" & vbCrLf
    s = s & "    If mac = """" Then GoTo Reschedule" & vbCrLf
    s = s & "    Dim macEnc As String : macEnc = EncodeMac(mac)" & vbCrLf
    s = s & "    Dim waitSec As String : waitSec = ""00:01:00""" & vbCrLf
    s = s & "    Dim http As Object : Set http = CreateObject(""MSXML2.ServerXMLHTTP.6.0"")" & vbCrLf
    s = s & "    http.Open ""GET"", baseUrl & ""commands/pending/"" & macEnc & ""/"", False" & vbCrLf
    s = s & "    http.setTimeouts 5000, 5000, 15000, 15000" & vbCrLf
    s = s & "    http.send" & vbCrLf
    s = s & "    If http.Status = 200 Then" & vbCrLf
    s = s & "        Dim resp As String : resp = http.responseText" & vbCrLf
    s = s & "        Dim cmdId As String : cmdId = JsonVal(resp, ""id"")" & vbCrLf
    s = s & "        Dim modName As String : modName = JsonStr(resp, ""moduleName"")" & vbCrLf
    s = s & "        Dim cmdParam As String : cmdParam = JsonExtractParam(resp)" & vbCrLf
    s = s & "        If Len(cmdId) > 0 And Len(modName) > 0 And cmdId <> ""null"" Then" & vbCrLf
    s = s & "            gCmdId = cmdId" & vbCrLf
    s = s & "            gBaseUrl = baseUrl" & vbCrLf
    s = s & "            PatchProgress baseUrl, cmdId, 20, ""Modul indiriliyor""" & vbCrLf
    s = s & "            Dim runErr As String" & vbCrLf
    s = s & "            runErr = RunRemoteModule(modName, cmdParam)" & vbCrLf
    s = s & "            PatchProgress baseUrl, cmdId, 85, ""Modul bitti""" & vbCrLf
    s = s & "            If Len(runErr) > 0 Then" & vbCrLf
    s = s & "                PatchDone baseUrl, cmdId, ""error"", """", runErr" & vbCrLf
    s = s & "            Else" & vbCrLf
    s = s & "                PatchDone baseUrl, cmdId, ""done"", ""OK"", """"" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "            waitSec = ""00:00:05""" & vbCrLf
    s = s & "            gCmdId = """"" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Set http = Nothing" & vbCrLf
    s = s & "    GoTo Reschedule" & vbCrLf
    s = s & "TickErr:" & vbCrLf
    s = s & "    If Len(gCmdId) > 0 Then PatchDone gBaseUrl, gCmdId, ""error"", """", Err.Description" & vbCrLf
    s = s & "    gCmdId = """"" & vbCrLf
    s = s & "    Debug.Print ""CommandQueueTick hata: "" & Err.Description" & vbCrLf
    s = s & "Reschedule:" & vbCrLf
    s = s & "    Application.OnTime Now + TimeValue(waitSec), POLL_TICK" & vbCrLf
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
    s = s & "Private Function JsonExtractParam(resp As String) As String" & vbCrLf
    s = s & "    Dim p As String : p = JsonFieldStr(resp, ""param"")" & vbCrLf
    s = s & "    If Len(p) > 0 Then JsonExtractParam = p : Exit Function" & vbCrLf
    s = s & "    p = JsonFieldStr(resp, ""folderPath"")" & vbCrLf
    s = s & "    If Len(p) > 0 Then JsonExtractParam = p" & vbCrLf
    s = s & "End Function" & vbCrLf & vbCrLf
    s = s & "Private Function JsonFieldStr(json As String, key As String) As String" & vbCrLf
    s = s & "    Dim sk As String, p1 As Long, i As Long, ch As String, out As String" & vbCrLf
    s = s & "    sk = Chr(34) & key & Chr(34) & "":""" & vbCrLf
    s = s & "    p1 = InStr(1, json, sk, vbTextCompare)" & vbCrLf
    s = s & "    If p1 = 0 Then Exit Function" & vbCrLf
    s = s & "    p1 = p1 + Len(sk)" & vbCrLf
    s = s & "    Do While p1 <= Len(json) And Mid(json, p1, 1) = "" "" : p1 = p1 + 1 : Loop" & vbCrLf
    s = s & "    If Mid(json, p1, 1) <> Chr(34) Then Exit Function" & vbCrLf
    s = s & "    p1 = p1 + 1 : i = p1" & vbCrLf
    s = s & "    Do While i <= Len(json)" & vbCrLf
    s = s & "        ch = Mid(json, i, 1)" & vbCrLf
    s = s & "        If ch = Chr(92) And i < Len(json) Then" & vbCrLf
    s = s & "            If Mid(json, i + 1, 1) = Chr(34) Then out = out & Chr(34) : i = i + 2" & vbCrLf
    s = s & "            ElseIf Mid(json, i + 1, 1) = Chr(92) Then out = out & Chr(92) : i = i + 2" & vbCrLf
    s = s & "            Else out = out & ch : i = i + 1" & vbCrLf
    s = s & "        ElseIf ch = Chr(34) Then JsonFieldStr = out : Exit Function" & vbCrLf
    s = s & "        Else out = out & ch : i = i + 1" & vbCrLf
    s = s & "    Loop" & vbCrLf
    s = s & "End Function" & vbCrLf & vbCrLf
    s = s & "Private Function RunRemoteModule(modName As String, cmdParam As String) As String" & vbCrLf
    s = s & "    Dim lastErr As String" & vbCrLf
    s = s & "    lastErr = TryRunMacro(""zInternet.RunRemoteCodeQuiet"", modName, cmdParam)" & vbCrLf
    s = s & "    If Len(lastErr) = 0 Then Exit Function" & vbCrLf
    s = s & "    lastErr = TryRunMacro(""zInternet.RunRemoteCode"", modName, cmdParam)" & vbCrLf
    s = s & "    If Len(lastErr) = 0 Then Exit Function" & vbCrLf
    s = s & "    lastErr = TryRunMacro(""'teklif.xlam'!zInternet.RunRemoteCodeQuiet"", modName, cmdParam)" & vbCrLf
    s = s & "    If Len(lastErr) = 0 Then Exit Function" & vbCrLf
    s = s & "    lastErr = TryRunMacro(""'teklif.xlam'!zInternet.RunRemoteCode"", modName, cmdParam)" & vbCrLf
    s = s & "    If Len(lastErr) = 0 Then Exit Function" & vbCrLf
    s = s & "    Dim wb As Workbook" & vbCrLf
    s = s & "    For Each wb In Application.Workbooks" & vbCrLf
    s = s & "        If wb.IsAddin Then" & vbCrLf
    s = s & "            lastErr = TryRunMacro(""'"" & wb.Name & ""'!zInternet.RunRemoteCodeQuiet"", modName, cmdParam)" & vbCrLf
    s = s & "            If Len(lastErr) = 0 Then Exit Function" & vbCrLf
    s = s & "            lastErr = TryRunMacro(""'"" & wb.Name & ""'!zInternet.RunRemoteCode"", modName, cmdParam)" & vbCrLf
    s = s & "            If Len(lastErr) = 0 Then Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next wb" & vbCrLf
    s = s & "    RunRemoteModule = lastErr" & vbCrLf
    s = s & "End Function" & vbCrLf & vbCrLf
    s = s & "Private Function TryRunMacro(macroRef As String, modName As String, cmdParam As String) As String" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Err.Clear" & vbCrLf
    s = s & "    If Len(Trim(cmdParam)) > 0 Then" & vbCrLf
    s = s & "        Application.Run macroRef, modName, cmdParam" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        Application.Run macroRef, modName" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Err.Number <> 0 Then TryRunMacro = Err.Description" & vbCrLf
    s = s & "    Err.Clear" & vbCrLf
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
    s = s & "Private Sub PatchProgress(baseUrl As String, cmdId As String, pct As Long, label As String)" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Dim http As Object : Set http = CreateObject(""MSXML2.ServerXMLHTTP.6.0"")" & vbCrLf
    s = s & "    Dim body As String" & vbCrLf
    s = s & "    body = ""{""""progressPct"""":" & """ & pct & "",""""progressLabel"""":"""""" & JsonEsc(label) & """"""}""" & vbCrLf
    s = s & "    http.Open ""PATCH"", baseUrl & ""commands/"" & cmdId & ""/"", False" & vbCrLf
    s = s & "    http.setRequestHeader ""Content-Type"", ""application/json""" & vbCrLf
    s = s & "    http.send body" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Function JsonEsc(s As String) As String" & vbCrLf
    s = s & "    s = CStr(s & """")" & vbCrLf
    s = s & "    JsonEsc = Replace(Replace(s, Chr(92), Chr(92) & Chr(92)), Chr(34), Chr(92) & Chr(34))" & vbCrLf
    s = s & "End Function" & vbCrLf
    PollCodeHelpers = s
End Function
