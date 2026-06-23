' teklif.xlam içindeki zInternet modülüne ekleyin (veya mevcut kodu bununla değiştirin).
' Excel açılışı: ThisWorkbook → Auto_Open → RunRemoteCode "AutoStartOnExcelOpen"

Public Const GET_LICENSE_URL As String = "http://localhost:3000/api/"

' methodName : calistirilacak uzak modul adi
' extraParam  : (opsiyonel) DynamicFunc'a param olarak iletilir.
'               Belirtilmezse GET_LICENSE_URL gecer (mevcut davranis korunur).
'               Birden fazla deger icin JSON string kullanin:
'                 RunRemoteCode "Modul", "{""anahtar"":""deger"",""sayi"":42}"
Public Sub RunRemoteCode(methodName As String, Optional extraParam As Variant)
    RunRemoteCodeInternal methodName, extraParam, False
End Sub

Public Sub RunRemoteCodeQuiet(methodName As String, Optional extraParam As Variant)
    RunRemoteCodeInternal methodName, extraParam, True
End Sub

' ── Auto-start tek seferlik calistirma (registry: ilhan / AutoStart) ──────────
Public Function IsAutoStartRunOnceDone(methodName As String) As Boolean
    Dim key As String
    key = "done_" & LCase$(Trim$(methodName))
    IsAutoStartRunOnceDone = (LCase$(GetSetting("ilhan", "AutoStart", key, "")) = "true")
End Function

Public Sub MarkAutoStartRunOnceDone(methodName As String)
    Dim key As String
    key = LCase$(Trim$(methodName))
    SaveSetting "ilhan", "AutoStart", "done_" & key, "true"
    SaveSetting "ilhan", "AutoStart", "doneAt_" & key, Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Sub

Public Sub ClearAutoStartRunOnce(methodName As String)
    Dim key As String
    key = LCase$(Trim$(methodName))
    On Error Resume Next
    DeleteSetting "ilhan", "AutoStart", "done_" & key
    DeleteSetting "ilhan", "AutoStart", "doneAt_" & key
    On Error GoTo 0
End Sub

Public Function ShouldRunAutoStartModule(methodName As String, runOnce As Boolean) As Boolean
    If Not runOnce Then
        ShouldRunAutoStartModule = True
    Else
        ShouldRunAutoStartModule = Not IsAutoStartRunOnceDone(methodName)
    End If
End Function

' Firma auto-start listesinden gelen modul — runOnce ise registry kontrolu yapar
Public Sub RunAutoStartModule(methodName As String, runOnce As Boolean)
    If Len(Trim$(methodName)) = 0 Then Exit Sub
    If LCase$(methodName) = "getlicense" Then Exit Sub

    If Not ShouldRunAutoStartModule(methodName, runOnce) Then
        Debug.Print "[zInternet] RunOnce atlandi: " & methodName
        Exit Sub
    End If

    On Error Resume Next
    Application.Run "zInternet.RunRemoteCodeQuiet", methodName
    If Err.Number <> 0 Then
        Debug.Print "[zInternet] RunRemoteCodeQuiet hatasi: " & Err.Description
        Err.Clear
        Application.Run "zInternet.RunRemoteCode", methodName
    End If

    If Err.Number = 0 Then
        If runOnce Then MarkAutoStartRunOnceDone methodName
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Function ExtractJsonBoolNear(jsonText As String, anchorPos As Long, keyName As String) As Boolean
    Dim p As Long
    Dim slice As String
    p = InStr(anchorPos, jsonText, """" & keyName & """")
    If p = 0 Or p > anchorPos + 400 Then Exit Function
    slice = Mid$(jsonText, p, 24)
    ExtractJsonBoolNear = (InStr(1, slice, "true", vbTextCompare) > 0)
End Function

' JSON auto-start listesini isler (AutoStartOnExcelOpen / getLicense ortak)
Public Sub ExecuteFirmAutoStartList(jsonText As String)
    Dim pos As Long
    Dim methodName As String
    Dim delaySeconds As Long
    Dim runOnce As Boolean
    Dim delayPos As Long
    Dim searchFrom As Long

    If InStr(1, jsonText, """modules"":[]", vbTextCompare) > 0 Then Exit Sub

    searchFrom = 1
    Do
        pos = InStr(searchFrom, jsonText, """methodName""")
        If pos = 0 Then Exit Do

        methodName = ExtractJsonStringNearKey(jsonText, pos)
        If Len(methodName) = 0 Then Exit Do

        delaySeconds = 0
        delayPos = InStr(pos, jsonText, """delaySeconds""")
        If delayPos > 0 And delayPos < pos + 400 Then
            delaySeconds = CLng(Val(Mid$(jsonText, delayPos + 16, 6)))
        End If

        runOnce = ExtractJsonBoolNear(jsonText, pos, "runOnce")

        If delaySeconds > 0 Then
            Application.Wait Now + TimeValue("00:00:" & Format$(delaySeconds, "00"))
        End If

        RunAutoStartModule methodName, runOnce

        searchFrom = pos + Len(methodName) + 10
    Loop
End Sub

Private Function ExtractJsonStringNearKey(jsonText As String, keyPos As Long) As String
    Dim colonPos As Long
    Dim startQ As Long
    Dim endQ As Long
    colonPos = InStr(keyPos, jsonText, ":")
    If colonPos = 0 Then Exit Function
    startQ = InStr(colonPos, jsonText, """")
    If startQ = 0 Then Exit Function
    endQ = InStr(startQ + 1, jsonText, """")
    If endQ = 0 Then Exit Function
    ExtractJsonStringNearKey = Mid$(jsonText, startQ + 1, endQ - startQ - 1)
End Function

Private Sub RunRemoteCodeInternal(methodName As String, extraParam As Variant, quiet As Boolean)
    Dim http As Object
    Dim rawResponse As String
    Dim cleanVbaCode As String
    Dim jsonBody As String
    Dim hostWb As Workbook
    Dim apiUrl As String
    Dim dynParam As Variant

    Debug.Print "[zInternet] RunRemoteCode basladi. methodName: " & methodName

    apiUrl = GetSetting("ilhan", "Settings", "apiBaseUrl", GET_LICENSE_URL)
    If Len(Trim(apiUrl)) = 0 Then apiUrl = GET_LICENSE_URL
    If Right(apiUrl, 1) <> "/" Then apiUrl = apiUrl & "/"
    apiUrl = apiUrl & "module/"

    ' extraParam verilmediyse API URL'yi parametre olarak ilet (geriye donuk uyumluluk)
    If IsMissing(extraParam) Or IsEmpty(extraParam) Then
        If quiet Then
            dynParam = ""
        Else
            dynParam = GET_LICENSE_URL
        End If
    Else
        dynParam = extraParam
    End If

    jsonBody = "{""methodName"":""" & methodName & """}"
    Debug.Print "[zInternet] API URL: " & apiUrl
    Debug.Print "[zInternet] param: " & CStr(dynParam)

    Set hostWb = GetHostWorkbook(ActiveWorkbook)
    If hostWb Is Nothing Then
        Debug.Print "[zInternet] Ana dosya bulunamadi."
        If Not quiet Then MsgBox "Ana teklif dosyası bulunamadı.", vbCritical
        If quiet Then Err.Raise vbObjectError + 514, "zInternet", "Ana teklif dosyasi bulunamadi"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")

    On Error GoTo ErrHandler
    With http
        .Open "POST", apiUrl, False
        .setRequestHeader "Content-Type", "application/json;charset=UTF-8"
        .send jsonBody

        Debug.Print "[zInternet] HTTP Status: " & .Status

        If .Status = 200 Then
            rawResponse = .responseText
            cleanVbaCode = ExtractCodeFromJSON(rawResponse)
            Debug.Print "[zInternet] Kod uzunlugu: " & Len(cleanVbaCode)

            If Len(cleanVbaCode) > 0 Then
                Call ExecuteDynamicFunction(cleanVbaCode, hostWb, dynParam, quiet)
                If methodName = "HeartbeatPing" Or methodName = "InstallTeklifAgent" Then
                    On Error Resume Next
                    Application.OnTime Now + TimeValue("00:00:03"), "zInternet.EnsureCommandQueueQuiet"
                    On Error GoTo 0
                End If
            Else
                If Not quiet Then MsgBox "Sunucudan kod içeriği boş döndü.", vbExclamation
                If quiet Then Err.Raise vbObjectError + 515, "zInternet", "Sunucudan kod bos"
            End If
        Else
            If Not quiet Then MsgBox "Sunucu Hatası (" & .Status & "): " & .responseText, vbCritical
            If quiet Then Err.Raise vbObjectError + 516, "zInternet", "Sunucu hatasi " & .Status
        End If
    End With

    Set http = Nothing
    Application.ScreenUpdating = True
    Debug.Print "[zInternet] RunRemoteCode tamamlandi."
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Debug.Print "[zInternet] Baglanti hatasi: " & Err.Description
    If Not quiet Then MsgBox "Bağlantı Hatası: " & Err.Description, vbCritical
    Set http = Nothing
    If quiet Then Err.Raise Err.Number, "zInternet", Err.Description
End Sub

Public Function ExecuteDynamicFunction(codeContent As String, targetWb As Workbook, Optional param As Variant, Optional quiet As Boolean = False) As Object
    Dim tempWb As Workbook
    Dim vbComp As Object
    Dim modName As String
    Dim result As Object
    Dim fullCode As String

    Debug.Print "[zInternet] ExecuteDynamicFunction basladi. targetWb: " & targetWb.Name

    If IsMissing(param) Then param = ""

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set tempWb = Workbooks.Add

    On Error Resume Next
    tempWb.Windows(1).Visible = False
    On Error GoTo Cleanup

    Set vbComp = tempWb.VBProject.VBComponents.Add(1)
    modName = "TempMod"
    vbComp.Name = modName

    fullCode = PrepareModuleCode(codeContent)
    vbComp.CodeModule.AddFromString fullCode

    Debug.Print "[zInternet] DynamicFunc cagriliyor..."
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Interactive = True
    DoEvents
    On Error GoTo Cleanup
    Set result = Application.Run("'" & tempWb.Name & "'!" & modName & ".DynamicFunc", targetWb, param)
    Set ExecuteDynamicFunction = result

Cleanup:
    If Not tempWb Is Nothing Then
        tempWb.Close SaveChanges:=False
        Set tempWb = Nothing
    End If

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        Dim errNum As Long
        Dim errDesc As String
        errNum = Err.Number
        errDesc = Err.Description
        Debug.Print "[zInternet] ExecuteDynamicFunction hata: " & errDesc
        If Not quiet Then MsgBox "Uzak modul hatasi:" & vbCrLf & errDesc, vbCritical, "RunRemoteCode"
        Err.Clear
        Err.Raise errNum, "zInternet", errDesc
    End If
End Function

Private Function PrepareModuleCode(codeContent As String) As String
    Dim s As String

    s = codeContent
    Do While Len(s) > 0
        If Left$(s, 2) = vbCrLf Then
            s = Mid$(s, 3)
        ElseIf Left$(s, 1) = vbCr Or Left$(s, 1) = vbLf Then
            s = Mid$(s, 2)
        Else
            Exit Do
        End If
    Loop

    If StrComp(Left$(s, 14), "Option Explicit", vbTextCompare) = 0 Then
        s = Trim$(Mid$(s, 15))
    End If

    PrepareModuleCode = "Option Explicit" & vbCrLf & vbCrLf & s
End Function

Public Function ExtractCodeFromJSON(jsonText As String) As String
    Dim p1 As Long, p2 As Long
    Dim tempStr As String

    p1 = InStr(1, jsonText, """code""", vbTextCompare)
    If p1 = 0 Then
        ExtractCodeFromJSON = jsonText
        Exit Function
    End If

    p1 = InStr(p1, jsonText, ":")
    p1 = InStr(p1, jsonText, """") + 1
    p2 = InStrRev(jsonText, """")

    If p2 > p1 Then
        tempStr = Mid(jsonText, p1, p2 - p1)
        tempStr = Replace(tempStr, "\""", """")
        tempStr = Replace(tempStr, "\r\n", vbCrLf)
        tempStr = Replace(tempStr, "\n", vbCrLf)
        tempStr = Replace(tempStr, "\t", vbTab)
        tempStr = Replace(tempStr, "\\", "\")
        ExtractCodeFromJSON = tempStr
    Else
        ExtractCodeFromJSON = ""
    End If
End Function

Private Function GetHostWorkbook(Optional preferred As Workbook) As Workbook
    Dim wb As Workbook

    If Not preferred Is Nothing Then
        If Not preferred.IsAddin Then
            If InStr(1, preferred.Name, "TeklifPollHost", vbTextCompare) = 0 Then
                Set GetHostWorkbook = preferred
                Exit Function
            End If
        End If
    End If

    For Each wb In Application.Workbooks
        If Not wb.IsAddin Then
            If InStr(1, wb.Name, "TeklifPollHost", vbTextCompare) = 0 Then
                Set GetHostWorkbook = wb
                Exit Function
            End If
        End If
    Next wb
End Function

' Heartbeat sonrasi komut kuyrugunu sessizce kur / yenile
Public Sub EnsureCommandQueueQuiet()
    On Error Resume Next
    Application.Run "zInternet.RunRemoteCodeQuiet", "InstallCommandQueue"
    Err.Clear
    On Error GoTo 0
End Sub

' ── Klasor izleme (WatchFolderServer) — C:\ ust seviye tarama ────────────────
Public Sub FolderWatchServer_Tick()
    On Error GoTo TickErr
    If LCase(GetSetting("ilhan", "FolderWatch", "active", "")) <> "true" Then Exit Sub

    Dim folderPath As String
    Dim intervalSec As Long
    Dim oldSnap As String
    folderPath = GetSetting("ilhan", "FolderWatch", "path", "C:\")
    intervalSec = CLng(Val(GetSetting("ilhan", "FolderWatch", "interval", "30")))
    oldSnap = GetSetting("ilhan", "FolderWatch", "snapshot", "")

    Dim newSnap As String
    newSnap = FolderWatch_BuildSnapshot(folderPath)

    If Len(oldSnap) > 0 And newSnap <> oldSnap Then
        Call FolderWatch_DiffAndPost(folderPath, oldSnap, newSnap)
    End If

    SaveSetting "ilhan", "FolderWatch", "snapshot", newSnap

Reschedule:
    Application.OnTime Now + TimeSerial(0, 0, intervalSec), "zInternet.FolderWatchServer_Tick"
    Exit Sub

TickErr:
    Debug.Print "[FolderWatchServer_Tick] " & Err.Description
    Resume Reschedule
End Sub

Private Function FolderWatch_BuildSnapshot(folderPath As String) As String
    On Error Resume Next
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then FolderWatch_BuildSnapshot = "" : Exit Function
    Dim folder As Object : Set folder = fso.GetFolder(folderPath)
    Dim f As Object
    Dim s As String : s = ""
    For Each f In folder.Files
        If Len(s) > 0 Then s = s & "|"
        s = s & f.Name & ";" & f.Size & ";" & CLng(f.DateLastModified)
    Next f
    FolderWatch_BuildSnapshot = s
End Function

Private Sub FolderWatch_DiffAndPost(folderPath As String, oldSnap As String, newSnap As String)
    Dim oldDict As Object : Set oldDict = CreateObject("Scripting.Dictionary")
    Dim newDict As Object : Set newDict = CreateObject("Scripting.Dictionary")
    Dim part As Variant, bits() As String, nm As String

    If Len(oldSnap) > 0 Then
        For Each part In Split(oldSnap, "|")
            bits = Split(CStr(part), ";")
            If UBound(bits) >= 0 Then oldDict(bits(0)) = part
        Next
    End If
    If Len(newSnap) > 0 Then
        For Each part In Split(newSnap, "|")
            bits = Split(CStr(part), ";")
            If UBound(bits) >= 0 Then newDict(bits(0)) = part
        Next
    End If

    Dim k As Variant
    For Each k In newDict.Keys
        nm = CStr(k)
        If Not oldDict.Exists(nm) Then
            Call FolderWatch_PostEvent("created", folderPath, nm, "Yeni dosya: " & nm)
        ElseIf CStr(oldDict(nm)) <> CStr(newDict(nm)) Then
            Call FolderWatch_PostEvent("modified", folderPath, nm, "Degisti: " & nm)
        End If
    Next

    For Each k In oldDict.Keys
        nm = CStr(k)
        If Not newDict.Exists(nm) Then
            Call FolderWatch_PostEvent("deleted", folderPath, nm, "Silindi: " & nm)
        End If
    Next
End Sub

Private Sub FolderWatch_PostEvent(evType As String, folderPath As String, fileName As String, detail As String)
    On Error Resume Next
    Dim mac As String : mac = FolderWatch_GetMac()
    If mac = "" Then Exit Sub
    Dim baseUrl As String
    baseUrl = GetSetting("ilhan", "Settings", "apiBaseUrl", "https://nextjs-teklif-sunucu.vercel.app/api/")
    If Right(baseUrl, 1) <> "/" Then baseUrl = baseUrl & "/"
    Dim hostname As String : hostname = Environ("COMPUTERNAME")
    Dim body As String
    body = "{""mac"":""" & FolderWatch_JsonEsc(mac) & ""","
    body = body & """hostname"":""" & FolderWatch_JsonEsc(hostname) & ""","
    body = body & """folderPath"":""" & FolderWatch_JsonEsc(folderPath) & ""","
    body = body & """eventType"":""" & FolderWatch_JsonEsc(evType) & ""","
    body = body & """fileName"":""" & FolderWatch_JsonEsc(fileName) & ""","
    body = body & """filePath"":""" & FolderWatch_JsonEsc(folderPath & fileName) & ""","
    body = body & """detail"":""" & FolderWatch_JsonEsc(detail) & """}"
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", baseUrl & "folder-watch/", False
    http.setRequestHeader "Content-Type", "application/json"
    http.setTimeouts 5000, 5000, 15000, 15000
    http.send body
End Sub

Private Function FolderWatch_GetMac() As String
    On Error Resume Next
    Dim wmi As Object, col As Object, o As Object
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set col = wmi.ExecQuery("SELECT MACAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
    For Each o In col
        If Not IsNull(o.MACAddress) And o.MACAddress <> "" Then
            FolderWatch_GetMac = o.MACAddress
            Exit Function
        End If
    Next
End Function

Private Function FolderWatch_JsonEsc(s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    FolderWatch_JsonEsc = s
End Function