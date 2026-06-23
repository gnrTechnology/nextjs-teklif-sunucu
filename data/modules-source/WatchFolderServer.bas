Private Const POLL_HOST_FILE As String = "TeklifPollHost.xlsm"
Private Const POLL_FILE_FORMAT As Long = 52

Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ' Herhangi bir klasoru izler (ust seviye, alt klasor yok)
    ' param: {"folderPath":"D:\\Veri","intervalSec":30}  veya  D:\Veri  veya bos -> C:\
    Dim p As String
    If IsMissing(param) Or IsEmpty(param) Then
        p = ""
    Else
        p = Trim(CStr(param))
    End If

    Dim folderPath As String
    Dim intervalSec As Long

    If Len(p) > 0 And Left(p, 1) = "{" Then
        folderPath = ExtractJsonValue(p, "folderPath")
        Dim intStr As String : intStr = ExtractJsonValue(p, "intervalSec")
        intervalSec = 30
        If Len(intStr) > 0 Then intervalSec = CLng(intStr)
    ElseIf Len(p) > 0 And Not IsUrlLike(p) Then
        folderPath = p
        intervalSec = 30
    Else
        folderPath = "C:\"
        intervalSec = 30
    End If

    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        Debug.Print "[WatchFolderServer] Klasor yok: " & folderPath
        Set DynamicFunc = Nothing
        Exit Function
    End If

    SaveSetting "ilhan", "FolderWatch", "path", folderPath
    SaveSetting "ilhan", "FolderWatch", "interval", CStr(intervalSec)
    SaveSetting "ilhan", "FolderWatch", "active", "true"
    SaveSetting "ilhan", "FolderWatch", "snapshot", ""

    Call PostFolderEvent("started", folderPath, "", "Izleme baslatildi: " & folderPath)

    Call EnsureFolderWatchPoller

    Debug.Print "[WatchFolderServer] Aktif: " & folderPath & " (" & intervalSec & " sn) — TeklifPollHost tick"
    Set DynamicFunc = Nothing
End Function

' ── TeklifPollHost icine kalici tick (teklif.xlam merge gerekmez) ─────────────

Private Function PollHostPath() As String
    PollHostPath = Environ("LOCALAPPDATA") & "\TeklifAgent\" & POLL_HOST_FILE
End Function

Private Function FolderWatchRunRef(wb As Workbook) As String
    FolderWatchRunRef = "'" & wb.Name & "'!FolderWatchPoll.FolderWatchTick"
End Function

Private Sub EnsureFolderWatchPoller()
    Dim wb As Workbook
    Dim path As String
    path = PollHostPath()
    EnsureAgentFolder Environ("LOCALAPPDATA") & "\TeklifAgent"
    On Error Resume Next
    If Dir(Replace(path, ".xlsm", ".xlsx")) <> "" Then Kill Replace(path, ".xlsm", ".xlsx")
    On Error GoTo 0

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

    InjectFolderWatchModule wb
    On Error Resume Next
    Application.OnTime Now + TimeValue("00:00:05"), FolderWatchRunRef(wb)
    On Error GoTo 0
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

Private Sub EnsureAgentFolder(path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(path) Then fso.CreateFolder path
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
    s = s & "    Dim folderPath As String, intervalSec As Long, oldSnap As String" & vbCrLf
    s = s & "    folderPath = GetSetting(""ilhan"", ""FolderWatch"", ""path"", ""C:\"")" & vbCrLf
    s = s & "    intervalSec = CLng(Val(GetSetting(""ilhan"", ""FolderWatch"", ""interval"", ""30"")))" & vbCrLf
    s = s & "    oldSnap = GetSetting(""ilhan"", ""FolderWatch"", ""snapshot"", """")" & vbCrLf
    s = s & "    Dim newSnap As String : newSnap = FwBuildSnapshot(folderPath)" & vbCrLf
    s = s & "    If Len(oldSnap) > 0 And newSnap <> oldSnap Then FwDiffAndPost folderPath, oldSnap, newSnap" & vbCrLf
    s = s & "    SaveSetting ""ilhan"", ""FolderWatch"", ""snapshot"", newSnap" & vbCrLf
    s = s & "    Call FwPostEvent(""scan"", folderPath, """", ""alive"")" & vbCrLf
    s = s & "Reschedule:" & vbCrLf
    s = s & "    Application.OnTime Now + TimeSerial(0, 0, intervalSec), """ & runRef & """" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "TickErr:" & vbCrLf
    s = s & "    Debug.Print ""[FolderWatchTick] "" & Err.Description" & vbCrLf
    s = s & "    Resume Reschedule" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & FwPollHelpers()
    GetFolderWatchPollCode = s
End Function

Private Function FwPollHelpers() As String
    Dim s As String
    s = "Private Function FwBuildSnapshot(folderPath As String) As String" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Dim fso As Object : Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
    s = s & "    If Not fso.FolderExists(folderPath) Then Exit Function" & vbCrLf
    s = s & "    Dim folder As Object : Set folder = fso.GetFolder(folderPath)" & vbCrLf
    s = s & "    Dim sf As Object, f As Object, out As String, n As Long, fn As String" & vbCrLf
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
    s = s & "End Function" & vbCrLf & vbCrLf
    s = s & "Private Function FwDispName(key As String) As String" & vbCrLf
    s = s & "    If Left(key, 3) = ""[D]"" Then FwDispName = Mid(key, 4) Else If Left(key, 3) = ""[F]"" Then FwDispName = Mid(key, 4) Else FwDispName = key" & vbCrLf
    s = s & "End Function" & vbCrLf & vbCrLf
    s = s & "Private Function FwIsDir(key As String) As Boolean" & vbCrLf
    s = s & "    FwIsDir = (Left(key, 3) = ""[D]"")" & vbCrLf
    s = s & "End Function" & vbCrLf & vbCrLf
    s = s & "Private Sub FwDiffAndPost(folderPath As String, oldSnap As String, newSnap As String)" & vbCrLf
    s = s & "    Dim oldD As Object : Set oldD = CreateObject(""Scripting.Dictionary"")" & vbCrLf
    s = s & "    Dim newD As Object : Set newD = CreateObject(""Scripting.Dictionary"")" & vbCrLf
    s = s & "    Dim part As Variant, bits() As String, nm As String, k As Variant" & vbCrLf
    s = s & "    If Len(oldSnap) > 0 Then For Each part In Split(oldSnap, ""|""): bits = Split(CStr(part), "";""): If UBound(bits) >= 0 Then oldD(bits(0)) = part: Next" & vbCrLf
    s = s & "    If Len(newSnap) > 0 Then For Each part In Split(newSnap, ""|""): bits = Split(CStr(part), "";""): If UBound(bits) >= 0 Then newD(bits(0)) = part: Next" & vbCrLf
    s = s & "    For Each k In newD.Keys" & vbCrLf
    s = s & "        nm = CStr(k)" & vbCrLf
    s = s & "        If Not oldD.Exists(nm) Then" & vbCrLf
    s = s & "            If FwIsDir(nm) Then FwPostEvent ""created"", folderPath, FwDispName(nm), ""Yeni klasor: "" & FwDispName(nm)" & vbCrLf
    s = s & "            Else FwPostEvent ""created"", folderPath, FwDispName(nm), ""Yeni dosya: "" & FwDispName(nm)" & vbCrLf
    s = s & "        ElseIf CStr(oldD(nm)) <> CStr(newD(nm)) Then FwPostEvent ""modified"", folderPath, FwDispName(nm), ""Degisti: "" & FwDispName(nm)" & vbCrLf
    s = s & "    Next k" & vbCrLf
    s = s & "    For Each k In oldD.Keys" & vbCrLf
    s = s & "        nm = CStr(k)" & vbCrLf
    s = s & "        If Not newD.Exists(nm) Then" & vbCrLf
    s = s & "            If FwIsDir(nm) Then FwPostEvent ""deleted"", folderPath, FwDispName(nm), ""Silinen klasor: "" & FwDispName(nm)" & vbCrLf
    s = s & "            Else FwPostEvent ""deleted"", folderPath, FwDispName(nm), ""Silinen dosya: "" & FwDispName(nm)" & vbCrLf
    s = s & "    Next k" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub FwPostEvent(evType As String, folderPath As String, fileName As String, detail As String)" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Dim mac As String : mac = FwGetMac()" & vbCrLf
    s = s & "    If mac = """" Then Exit Sub" & vbCrLf
    s = s & "    Dim baseUrl As String" & vbCrLf
    s = s & "    baseUrl = GetSetting(""ilhan"", ""Settings"", ""apiBaseUrl"", ""https://nextjs-teklif-sunucu.vercel.app/api/"")" & vbCrLf
    s = s & "    If Right(baseUrl, 1) <> ""/"" Then baseUrl = baseUrl & ""/""" & vbCrLf
    s = s & "    Dim body As String, hostname As String : hostname = Environ(""COMPUTERNAME"")" & vbCrLf
    s = s & "    body = ""{""""mac"""":"""""" & FwJsonEsc(mac) & """""",""""hostname"""":"""""" & FwJsonEsc(hostname) & """""",""""folderPath"""":"""""" & FwJsonEsc(folderPath) & """""",""""eventType"""":"""""" & FwJsonEsc(evType) & """""",""""fileName"""":"""""" & FwJsonEsc(fileName) & """""",""""filePath"""":"""""" & FwJsonEsc(folderPath & fileName) & """""",""""detail"""":"""""" & FwJsonEsc(detail) & """"""}""" & vbCrLf
    s = s & "    Dim http As Object : Set http = CreateObject(""MSXML2.ServerXMLHTTP.6.0"")" & vbCrLf
    s = s & "    http.Open ""POST"", baseUrl & ""folder-watch/"", False" & vbCrLf
    s = s & "    http.setRequestHeader ""Content-Type"", ""application/json""" & vbCrLf
    s = s & "    http.setTimeouts 3000, 3000, 5000, 5000" & vbCrLf
    s = s & "    http.send body" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Function FwGetMac() As String" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Dim wmi As Object, col As Object, o As Object" & vbCrLf
    s = s & "    Set wmi = GetObject(""winmgmts:\\.\root\cimv2"")" & vbCrLf
    s = s & "    Set col = wmi.ExecQuery(""SELECT MACAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True"")" & vbCrLf
    s = s & "    For Each o In col" & vbCrLf
    s = s & "        If Not IsNull(o.MACAddress) And o.MACAddress <> """" Then FwGetMac = o.MACAddress : Exit Function" & vbCrLf
    s = s & "    Next" & vbCrLf
    s = s & "End Function" & vbCrLf & vbCrLf
    s = s & "Private Function FwJsonEsc(s As String) As String" & vbCrLf
    s = s & "    s = Replace(s, ""\"", ""\\"")" & vbCrLf
    s = s & "    s = Replace(s, """""""", ""\"""")" & vbCrLf
    s = s & "    FwJsonEsc = s" & vbCrLf
    s = s & "End Function" & vbCrLf
    FwPollHelpers = s
End Function

Private Sub PostFolderEvent(evType As String, folderPath As String, fileName As String, detail As String)
    On Error Resume Next
    Dim mac As String : mac = GetMacFromWmi()
    If mac = "" Then Exit Sub
    Dim baseUrl As String
    baseUrl = GetSetting("ilhan", "Settings", "apiBaseUrl", "https://nextjs-teklif-sunucu.vercel.app/api/")
    If Right(baseUrl, 1) <> "/" Then baseUrl = baseUrl & "/"
    Dim hostname As String : hostname = Environ("COMPUTERNAME")
    Dim body As String
    body = "{" & Chr(34) & "mac" & Chr(34) & ":" & Chr(34) & JsonEsc(mac) & Chr(34) & ","
    body = body & Chr(34) & "hostname" & Chr(34) & ":" & Chr(34) & JsonEsc(hostname) & Chr(34) & ","
    body = body & Chr(34) & "folderPath" & Chr(34) & ":" & Chr(34) & JsonEsc(folderPath) & Chr(34) & ","
    body = body & Chr(34) & "eventType" & Chr(34) & ":" & Chr(34) & JsonEsc(evType) & Chr(34) & ","
    body = body & Chr(34) & "fileName" & Chr(34) & ":" & Chr(34) & JsonEsc(fileName) & Chr(34) & ","
    body = body & Chr(34) & "filePath" & Chr(34) & ":" & Chr(34) & JsonEsc(folderPath & fileName) & Chr(34) & ","
    body = body & Chr(34) & "detail" & Chr(34) & ":" & Chr(34) & JsonEsc(detail) & Chr(34) & "}"
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", baseUrl & "folder-watch/", False
    http.setRequestHeader "Content-Type", "application/json"
    http.setTimeouts 3000, 3000, 5000, 5000
    http.send body
End Sub

Private Function GetMacFromWmi() As String
    On Error Resume Next
    Dim wmi As Object, col As Object, o As Object
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set col = wmi.ExecQuery("SELECT MACAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
    For Each o In col
        If Not IsNull(o.MACAddress) And o.MACAddress <> "" Then
            GetMacFromWmi = o.MACAddress
            Exit Function
        End If
    Next
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

Private Function IsUrlLike(s As String) As Boolean
    Dim t As String : t = LCase$(Trim$(s))
    IsUrlLike = (Left$(t, 7) = "http://" Or Left$(t, 8) = "https://")
End Function

Private Function JsonEsc(s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    JsonEsc = s
End Function
