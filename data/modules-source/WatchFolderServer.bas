Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ' C:\ kok klasorunu izler (ust seviye), degisiklikleri sunucuya POST eder
    ' param: {"folderPath":"C:\\","intervalSec":30}  veya bos -> varsayilan C:\
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
    ElseIf Len(p) > 0 Then
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
    SaveSetting "ilhan", "FolderWatch", "snapshot", BuildFolderSnapshot(folderPath)

    Call PostFolderEvent("started", folderPath, "", "Izleme baslatildi: " & folderPath)

    On Error Resume Next
    Application.OnTime Now + TimeValue("00:00:05"), "zInternet.FolderWatchServer_Tick"
    On Error GoTo 0

    Debug.Print "[WatchFolderServer] Aktif: " & folderPath & " (" & intervalSec & " sn)"
    Set DynamicFunc = Nothing
End Function

Private Function BuildFolderSnapshot(folderPath As String) As String
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder As Object : Set folder = fso.GetFolder(folderPath)
    Dim f As Object
    Dim s As String : s = ""
    For Each f In folder.Files
        If Len(s) > 0 Then s = s & "|"
        s = s & f.Name & ";" & f.Size & ";" & CLng(f.DateLastModified)
    Next f
    BuildFolderSnapshot = s
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
    http.setTimeouts 5000, 5000, 15000, 15000
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

Private Function JsonEsc(s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    JsonEsc = s
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
