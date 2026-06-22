Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ' param: {"folderPath":"C:\\Izle","triggerModule":"ImportCsvToSheet","intervalSec":10}
    Dim p As String : p = CStr(param)
    Dim folderPath    As String : folderPath    = ExtractJsonValue(p, "folderPath")
    Dim triggerModule As String : triggerModule = ExtractJsonValue(p, "triggerModule")
    Dim intervalStr   As String : intervalStr   = ExtractJsonValue(p, "intervalSec")

    Dim intervalSec As Long : intervalSec = 10
    If Len(intervalStr) > 0 Then intervalSec = CLng(intervalStr)

    If Len(folderPath) = 0 Then
        MsgBox "folderPath parametresi gerekli.", vbExclamation
        Set DynamicFunc = Nothing
        Exit Function
    End If

    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        MsgBox "Klasör bulunamadı: " & folderPath, vbExclamation
        Set DynamicFunc = Nothing
        Exit Function
    End If

    ' Mevcut dosya sayısını baseline olarak kaydet
    Dim folder As Object : Set folder = fso.GetFolder(folderPath)
    Dim baseCount As Long : baseCount = folder.Files.Count

    SaveSetting "ilhan", "FolderMonitor", "path",     folderPath
    SaveSetting "ilhan", "FolderMonitor", "module",   triggerModule
    SaveSetting "ilhan", "FolderMonitor", "baseCount", CStr(baseCount)
    SaveSetting "ilhan", "FolderMonitor", "interval",  CStr(intervalSec)

    Application.OnTime Now + TimeSerial(0, 0, intervalSec), "FolderMonitor_Check"

    MsgBox "Klasör izleniyor: " & folderPath & vbCrLf & _
           "Yeni dosya gelince " & triggerModule & " çalışacak." & vbCrLf & _
           "Kontrol sıklığı: " & intervalSec & " saniye", vbInformation
    Set DynamicFunc = Nothing
End Function

' teklif.xlam içine ekle:
'
' Public Sub FolderMonitor_Check()
'     Dim folderPath As String : folderPath = GetSetting("ilhan","FolderMonitor","path","")
'     Dim modName    As String : modName    = GetSetting("ilhan","FolderMonitor","module","")
'     Dim baseCount  As Long   : baseCount  = CLng(GetSetting("ilhan","FolderMonitor","baseCount","0"))
'     Dim intervalSec As Long  : intervalSec = CLng(GetSetting("ilhan","FolderMonitor","interval","10"))
'     If Len(folderPath) = 0 Then Exit Sub
'     Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
'     If Not fso.FolderExists(folderPath) Then Exit Sub
'     Dim folder As Object : Set folder = fso.GetFolder(folderPath)
'     If folder.Files.Count > baseCount Then
'         SaveSetting "ilhan","FolderMonitor","baseCount", CStr(folder.Files.Count)
'         If Len(modName) > 0 Then
'             On Error Resume Next
'             Application.Run "zInternet.RunRemoteCode", modName
'             On Error GoTo 0
'         End If
'     End If
'     Application.OnTime Now + TimeSerial(0,0,intervalSec), "FolderMonitor_Check"
' End Sub

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
