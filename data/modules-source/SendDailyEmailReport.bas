Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ' param: {"to":"a@b.com","subject":"Günlük Rapor","sheetName":"Rapor","scheduledHour":8}
    Dim p As String : p = CStr(param)
    Dim toAddr      As String : toAddr      = ExtractJsonValue(p, "to")
    Dim subject     As String : subject     = ExtractJsonValue(p, "subject")
    Dim sheetName   As String : sheetName   = ExtractJsonValue(p, "sheetName")
    Dim hourStr     As String : hourStr     = ExtractJsonValue(p, "scheduledHour")

    If Len(subject)   = 0 Then subject   = "Günlük Rapor - " & Format(Now, "dd.MM.yyyy")
    If Len(sheetName) = 0 Then sheetName = targetWb.Sheets(1).Name

    ' PDF oluştur
    Dim pdfPath As String
    pdfPath = Environ("TEMP") & "\rapor_" & Format(Now, "yyyyMMdd") & ".pdf"

    On Error GoTo HataIsle
    Dim ws As Worksheet : Set ws = targetWb.Sheets(sheetName)
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard

    If Len(toAddr) > 0 Then
        ' Outlook COM ile gönder
        Dim olApp As Object : Set olApp = CreateObject("Outlook.Application")
        Dim mail  As Object : Set mail  = olApp.CreateItem(0)
        With mail
            .To      = toAddr
            .Subject = subject
            .Body    = "Merhaba," & vbCrLf & vbCrLf & _
                       Format(Now, "dd.MM.yyyy HH:mm") & " tarihli rapor ektedir." & vbCrLf & vbCrLf & _
                       "Saygılarımızla."
            .Attachments.Add pdfPath
            .Send
        End With
        MsgBox "Rapor gönderildi: " & toAddr, vbInformation
    Else
        ' E-posta adresi yoksa PDF'i aç
        Shell "explorer.exe """ & pdfPath & """", vbNormalFocus
        MsgBox "PDF oluşturuldu: " & pdfPath & vbCrLf & "(to adresi belirtilmedi)", vbInformation
    End If

    Set DynamicFunc = Nothing
    Exit Function

HataIsle:
    MsgBox "Rapor hatası: " & Err.Description, vbCritical
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
