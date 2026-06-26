Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim baseUrl As String
    baseUrl = Trim(CStr(param))
    If Len(baseUrl) = 0 Then baseUrl = "https://nextjs-teklif-sunucu.vercel.app/api/"
    If Right(baseUrl, 1) <> "/" Then baseUrl = baseUrl & "/"

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    On Error GoTo HataIsle

    http.Open "GET", baseUrl & "modules", False
    http.setTimeouts 5000, 10000, 30000, 30000
    http.send

    If http.Status <> 200 Then
        Debug.Print "[AutoUpdateModules] Sunucuya baglanilamadi. HTTP " & http.Status
        Set DynamicFunc = Nothing
        Exit Function
    End If

    Dim response As String
    response = http.responseText
    Set http = Nothing

    Dim count As Long
    count = 0
    Dim pos As Long
    pos = 1
    Do
        pos = InStr(pos, response, """methodName""", vbTextCompare)
        If pos = 0 Then Exit Do
        count = count + 1
        pos = pos + 12
    Loop

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = targetWb.Sheets("Moduller")
    If ws Is Nothing Then
        Set ws = targetWb.Sheets.Add
        ws.Name = "Moduller"
    End If
    On Error GoTo HataIsle

    ws.Cells.ClearContents
    ws.Range("A1").Value = "Sunucudaki Moduller (" & Format(Now, "dd.MM.yyyy HH:mm") & ")"
    ws.Range("A2").Value = "Toplam modul: " & count

    Dim row As Long
    row = 4
    ws.Range("A3").Value = "Modul Adi"
    ws.Range("B3").Value = "Kategori"
    ws.Range("C3").Value = "Durum"

    pos = 1
    Dim methodName As String
    Dim category As String
    Dim active As String
    Do
        Dim mPos As Long
        mPos = InStr(pos, response, """methodName""", vbTextCompare)
        If mPos = 0 Then Exit Do
        methodName = ExtractJsonStringAt(response, mPos + 14)
        Dim cPos As Long
        cPos = InStr(mPos, response, """category""", vbTextCompare)
        Dim aPos As Long
        aPos = InStr(mPos, response, """active""", vbTextCompare)
        If cPos > 0 And cPos < mPos + 500 Then
            category = ExtractJsonStringAt(response, cPos + 11)
        Else
            category = "genel"
        End If
        If aPos > 0 And aPos < mPos + 500 Then
            active = IIf(InStr(Mid(response, aPos, 20), "true", vbTextCompare) > 0, "Aktif", "Pasif")
        Else
            active = "?"
        End If

        ws.Cells(row, 1).Value = methodName
        ws.Cells(row, 2).Value = category
        ws.Cells(row, 3).Value = active
        row = row + 1
        pos = mPos + 14
        If row > 83 Then Exit Do
    Loop

    ws.Range("A1:C" & row - 1).Columns.AutoFit

    Debug.Print "[AutoUpdateModules] " & count & " modul sunucudan alindi. Ilk " & (row - 4) & " satir Moduller sayfasina yazildi."
    Set DynamicFunc = Nothing
    Exit Function

HataIsle:
    Debug.Print "[AutoUpdateModules] Baglanti hatasi: " & Err.Description
    Set DynamicFunc = Nothing
End Function

Private Function ExtractJsonStringAt(json As String, startPos As Long) As String
    Dim p1 As Long
    Dim p2 As Long
    p1 = InStr(startPos, json, """")
    If p1 = 0 Then Exit Function
    p1 = p1 + 1
    p2 = InStr(p1, json, """")
    If p2 > p1 Then ExtractJsonStringAt = Mid(json, p1, p2 - p1)
End Function
