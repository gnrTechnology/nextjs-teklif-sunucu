Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ' Sunucudan modül listesini çeker; yeni / güncellenen modülleri rapor eder.
    ' param: API base URL
    Dim baseUrl As String
    baseUrl = Trim(CStr(param))
    If Len(baseUrl) = 0 Then baseUrl = "http://localhost:3000/api/"
    If Right(baseUrl, 1) <> "/" Then baseUrl = baseUrl & "/"

    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    On Error GoTo HataIsle

    ' GET /api/modules
    http.Open "GET", baseUrl & "modules", False
    http.setTimeouts 5000, 10000, 30000, 30000
    http.send

    If http.Status <> 200 Then
        MsgBox "Sunucuya bağlanılamadı. HTTP " & http.Status, vbExclamation
        Set DynamicFunc = Nothing
        Exit Function
    End If

    Dim response As String : response = http.responseText
    Set http = Nothing

    ' Modül sayısını raporla
    Dim count As Long : count = 0
    Dim pos As Long : pos = 1
    Do
        pos = InStr(pos, response, """methodName""", vbTextCompare)
        If pos = 0 Then Exit Do
        count = count + 1
        pos = pos + 12
    Loop

    ' Sonucu sayfaya yaz
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = targetWb.Sheets("Modüller")
    If ws Is Nothing Then
        Set ws = targetWb.Sheets.Add
        ws.Name = "Modüller"
    End If
    On Error GoTo 0

    ws.Range("A1").Value = "Sunucudaki Modüller (" & Format(Now, "dd.MM.yyyy HH:mm") & ")"
    ws.Range("A2").Value = "Toplam modül: " & count

    ' Basit JSON parsing - methodName değerlerini listele
    Dim row As Long : row = 4
    ws.Range("A3").Value = "Modül Adı"
    ws.Range("B3").Value = "Kategori"
    ws.Range("C3").Value = "Durum"

    pos = 1
    Dim methodName As String, category As String, active As String
    Do
        Dim mPos As Long : mPos = InStr(pos, response, """methodName""", vbTextCompare)
        If mPos = 0 Then Exit Do
        methodName = ExtractJsonStringAt(response, mPos + 14)
        Dim cPos As Long : cPos = InStr(mPos, response, """category""", vbTextCompare)
        Dim aPos As Long : aPos = InStr(mPos, response, """active""", vbTextCompare)
        If cPos > 0 Then category = ExtractJsonStringAt(response, cPos + 11) Else category = "genel"
        If aPos > 0 Then active = IIf(InStr(response, "true", vbTextCompare) > aPos - 5, "Aktif", "Pasif") Else active = "?"

        ws.Cells(row, 1).Value = methodName
        ws.Cells(row, 2).Value = category
        ws.Cells(row, 3).Value = active
        row = row + 1
        pos = mPos + 14
    Loop

    ws.Columns("A:C").AutoFit
    ws.Activate

    MsgBox count & " modül sunucudan alındı. 'Modüller' sayfasına yazıldı.", vbInformation
    Set DynamicFunc = Nothing
    Exit Function

HataIsle:
    MsgBox "Bağlantı hatası: " & Err.Description, vbCritical
    Set DynamicFunc = Nothing
End Function

Private Function ExtractJsonStringAt(json As String, startPos As Long) As String
    Dim p1 As Long, p2 As Long
    p1 = InStr(startPos, json, """")
    If p1 = 0 Then Exit Function
    p1 = p1 + 1
    p2 = InStr(p1, json, """")
    If p2 > p1 Then ExtractJsonStringAt = Mid(json, p1, p2 - p1)
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
    End If
End Function
