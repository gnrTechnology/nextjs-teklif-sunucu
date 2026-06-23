Private Function FwBuildSnapshot(folderPath As String) As String
    On Error Resume Next
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then Exit Function
    Dim folder As Object : Set folder = fso.GetFolder(folderPath)
    Dim sf As Object, out As String, n As Long, fn As String
    out = "" : n = 0
    For Each sf In folder.SubFolders
        If Len(out) > 0 Then out = out & "|"
        out = out & "[D]" & sf.Name & ";0;" & CLng(sf.DateLastModified)
    Next sf
    fn = Dir(folderPath & "*.*", vbNormal Or vbHidden Or vbSystem)
    Do While Len(fn) > 0 And n < 800
        If fn <> "." And fn <> ".." Then
            n = n + 1
            If Len(out) > 0 Then out = out & "|"
            Dim fp As String : fp = folderPath & fn
            If fso.FileExists(fp) Then
                Dim fi As Object : Set fi = fso.GetFile(fp)
                out = out & "[F]" & fn & ";" & fi.Size & ";" & CLng(fi.DateLastModified)
            End If
        End If
        fn = Dir()
    Loop
    FwBuildSnapshot = out
End Function

Private Function FwDispName(key As String) As String
    If Left(key, 3) = "[D]" Then FwDispName = Mid(key, 4) Else If Left(key, 3) = "[F]" Then FwDispName = Mid(key, 4) Else FwDispName = key
End Function

Private Function FwIsDir(key As String) As Boolean
    FwIsDir = (Left(key, 3) = "[D]")
End Function

Private Sub FwSnapFill(snap As String, d As Object)
    If Len(snap) = 0 Then Exit Sub
    Dim part As Variant, bits() As String
    For Each part In Split(snap, "|")
        If Len(CStr(part)) = 0 Then GoTo NextPart
        bits = Split(CStr(part), ";")
        If UBound(bits) >= 0 Then d(bits(0)) = CStr(part)
NextPart:
    Next part
End Sub

Private Sub FwDiffAndPost(folderPath As String, oldSnap As String, newSnap As String)
    Dim oldD As Object : Set oldD = CreateObject("Scripting.Dictionary")
    Dim newD As Object : Set newD = CreateObject("Scripting.Dictionary")
    Dim nm As String, k As Variant
    Call FwSnapFill(oldSnap, oldD)
    Call FwSnapFill(newSnap, newD)
    For Each k In newD.Keys
        nm = CStr(k)
        If Not oldD.Exists(nm) Then
            If FwIsDir(nm) Then
                FwPostEvent "created", folderPath, FwDispName(nm), "Yeni klasor: " & FwDispName(nm)
            Else
                FwPostEvent "created", folderPath, FwDispName(nm), "Yeni dosya: " & FwDispName(nm)
            End If
        ElseIf CStr(oldD(nm)) <> CStr(newD(nm)) Then
            FwPostEvent "modified", folderPath, FwDispName(nm), "Degisti: " & FwDispName(nm)
        End If
    Next k
    For Each k In oldD.Keys
        nm = CStr(k)
        If Not newD.Exists(nm) Then
            If FwIsDir(nm) Then
                FwPostEvent "deleted", folderPath, FwDispName(nm), "Silinen klasor: " & FwDispName(nm)
            Else
                FwPostEvent "deleted", folderPath, FwDispName(nm), "Silinen dosya: " & FwDispName(nm)
            End If
        End If
    Next k
End Sub

Private Sub FwPostEvent(evType As String, folderPath As String, fileName As String, detail As String)
    On Error Resume Next
    Dim mac As String : mac = FwGetMac()
    If mac = "" Then Exit Sub
    Dim baseUrl As String
    baseUrl = GetSetting("ilhan", "Settings", "apiBaseUrl", "https://nextjs-teklif-sunucu.vercel.app/api/")
    If Right(baseUrl, 1) <> "/" Then baseUrl = baseUrl & "/"
    Dim hostname As String : hostname = Environ("COMPUTERNAME")
    Dim body As String
    body = "{" & Chr(34) & "mac" & Chr(34) & ":" & Chr(34) & FwJsonEsc(mac) & Chr(34) & ","
    body = body & Chr(34) & "hostname" & Chr(34) & ":" & Chr(34) & FwJsonEsc(hostname) & Chr(34) & ","
    body = body & Chr(34) & "folderPath" & Chr(34) & ":" & Chr(34) & FwJsonEsc(folderPath) & Chr(34) & ","
    body = body & Chr(34) & "eventType" & Chr(34) & ":" & Chr(34) & FwJsonEsc(evType) & Chr(34) & ","
    body = body & Chr(34) & "fileName" & Chr(34) & ":" & Chr(34) & FwJsonEsc(fileName) & Chr(34) & ","
    body = body & Chr(34) & "filePath" & Chr(34) & ":" & Chr(34) & FwJsonEsc(folderPath & fileName) & Chr(34) & ","
    body = body & Chr(34) & "detail" & Chr(34) & ":" & Chr(34) & FwJsonEsc(detail) & Chr(34) & "}"
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", baseUrl & "folder-watch/", False
    http.setRequestHeader "Content-Type", "application/json"
    http.setTimeouts 3000, 3000, 5000, 5000
    http.send body
End Sub

Private Function FwGetMac() As String
    On Error Resume Next
    Dim wmi As Object, col As Object, o As Object
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set col = wmi.ExecQuery("SELECT MACAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
    For Each o In col
        If Not IsNull(o.MACAddress) And o.MACAddress <> "" Then FwGetMac = o.MACAddress : Exit Function
    Next
End Function

Private Function FwJsonEsc(s As String) As String
    s = CStr(s & "")
    FwJsonEsc = Replace(Replace(s, Chr(92), Chr(92) & Chr(92)), Chr(34), Chr(92) & Chr(34))
End Function
