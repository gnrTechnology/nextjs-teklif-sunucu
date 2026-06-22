Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ' param: {"folderPath":"C:\\Yedek","daysOld":30,"pattern":"*.bak"}
    Dim p As String : p = CStr(param)
    Dim folderPath As String : folderPath = ExtractJsonValue(p, "folderPath")
    Dim daysStr    As String : daysStr    = ExtractJsonValue(p, "daysOld")
    Dim pattern    As String : pattern    = ExtractJsonValue(p, "pattern")

    Dim daysOld As Long : daysOld = 30
    If Len(daysStr) > 0 Then daysOld = CLng(daysStr)
    If Len(pattern) = 0 Then pattern = "*.*"
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

    Dim cutOff As Date : cutOff = Now - daysOld
    Dim folder As Object : Set folder = fso.GetFolder(folderPath)
    Dim deleted As Long, skipped As Long, totalSize As Double

    Dim f As Object
    For Each f In folder.Files
        If f.DateLastModified < cutOff Then
            On Error Resume Next
            totalSize = totalSize + f.Size
            f.Delete True
            If Err.Number = 0 Then
                deleted = deleted + 1
            Else
                skipped = skipped + 1
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next f

    MsgBox "Temizlik tamamlandı." & vbCrLf & vbCrLf & _
           "Silinen: " & deleted & " dosya (" & Format(totalSize / 1048576, "0.0") & " MB)" & vbCrLf & _
           "Atlandı: " & skipped & " dosya (kilitli/izin yok)" & vbCrLf & _
           "Kriter: " & daysOld & " günden eski, " & folderPath, vbInformation

    Set DynamicFunc = Nothing
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
