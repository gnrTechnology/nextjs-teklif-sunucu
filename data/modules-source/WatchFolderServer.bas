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

    ' PollHost icinde modul enjekte etme — deadlock. Sadece tick zamanla.
    Call ScheduleFolderWatchTickOnly

    Debug.Print "[WatchFolderServer] Aktif: " & folderPath & " (" & intervalSec & " sn)"
    Set DynamicFunc = Nothing
End Function

Private Sub ScheduleFolderWatchTickOnly()
    On Error Resume Next
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If InStr(1, wb.Name, "TeklifPollHost", vbTextCompare) > 0 Then
            Application.OnTime Now + TimeValue("00:00:05"), "'" & wb.Name & "'!FolderWatchPoll.FolderWatchTick"
            Exit Sub
        End If
    Next wb
    Debug.Print "[WatchFolderServer] FolderWatchPoll yok — once InstallCommandQueue calistirin"
End Sub

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
    s = CStr(s & "")
    JsonEsc = Replace(Replace(s, Chr(92), Chr(92) & Chr(92)), Chr(34), Chr(92) & Chr(34))
End Function
