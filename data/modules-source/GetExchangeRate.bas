Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Döviz" : ws.Range("B1").Value = "Alış" : ws.Range("C1").Value = "Satış"
    ws.Range("A1:C1").Font.Bold = True
    ' exchangerate-api (ücretsiz, kayıt gerektirmeyen)
    On Error Resume Next
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", "https://api.exchangerate-api.com/v4/latest/TRY", False
    http.setTimeouts 5000, 10000, 20000, 20000
    http.send
    If http.Status = 200 Then
        Dim resp As String : resp = http.responseText
        ' Basit JSON parse - rates içinden USD, EUR, GBP
        Dim currencies() As String : currencies = Split("USD,EUR,GBP,CHF,JPY", ",")
        Dim r As Long : r = 2
        Dim c As Variant
        For Each c In currencies
            Dim key As String : key = Chr(34) & c & Chr(34) & ":"
            Dim p1 As Long : p1 = InStr(resp, key)
            If p1 > 0 Then
                p1 = p1 + Len(key)
                Dim p2 As Long : p2 = p1
                Do While InStr(",}", Mid(resp,p2,1)) = 0 : p2=p2+1 : Loop
                Dim rate As Double : rate = CDbl(Trim(Mid(resp,p1,p2-p1)))
                ws.Cells(r,1).Value = c
                ws.Cells(r,2).Value = Format(1/rate,"0.0000") & " TL"
                r = r + 1
            End If
        Next c
        ws.Cells(r,1).Value = "Güncelleme" : ws.Cells(r,2).Value = Now()
    Else
        ws.Range("A2").Value = "Hata: HTTP " & http.Status
    End If
    On Error GoTo 0
    ws.Columns.AutoFit

    ' Sonuclari sunucuya gonder
    Dim outJson As String : outJson = "{"
    Dim rr As Long : Dim firstField As Boolean : firstField = True
    For rr = 1 To ws.UsedRange.Rows.Count
        Dim k As String : k = Trim(CStr(ws.Cells(rr, 1).Value))
        Dim v As String : v = Trim(CStr(ws.Cells(rr, 2).Value))
        If k <> "" Then
            If Not firstField Then outJson = outJson & ","
            k = Replace(Replace(k, Chr(34), "'"), "|", "-")
            v = Replace(Replace(Replace(v, Chr(34), "'"), Chr(10), " "), Chr(13), "")
            outJson = outJson & Chr(34) & k & Chr(34) & ":" & Chr(34) & v & Chr(34)
            firstField = False
        End If
    Next rr
    outJson = outJson & "}"
    Call PostOutputToServer("GetExchangeRate", outJson)

    Set DynamicFunc = Nothing
End Function
Private Sub PostOutputToServer(moduleName As String, outputJson As String)
    On Error Resume Next
    Dim mac      As String : mac      = GetSetting("ilhan", "Settings", "mac", "")
    If mac = "" Then Exit Sub
    Dim baseUrl  As String : baseUrl  = GetSetting("ilhan", "Settings", "apiBaseUrl", "https://nextjs-teklif-sunucu.vercel.app/api/")
    If Right(baseUrl, 1) <> "/" Then baseUrl = baseUrl & "/"
    Dim hostname As String : hostname = Environ("COMPUTERNAME")
    Dim firmaAdi As String : firmaAdi = GetSetting("ilhan", "Settings", "TBveren", "")
    Dim body As String
    body = "{" & Chr(34) & "mac" & Chr(34) & ":" & Chr(34) & Replace(mac, Chr(34), "'") & Chr(34) & ","
    body = body & Chr(34) & "moduleName" & Chr(34) & ":" & Chr(34) & moduleName & Chr(34) & ","
    body = body & Chr(34) & "hostname" & Chr(34) & ":" & Chr(34) & Replace(hostname, Chr(34), "'") & Chr(34) & ","
    body = body & Chr(34) & "firmaAdi" & Chr(34) & ":" & Chr(34) & Replace(firmaAdi, Chr(34), "'") & Chr(34) & ","
    body = body & Chr(34) & "output" & Chr(34) & ":" & outputJson & "}"
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", baseUrl & "module-output", False
    http.setRequestHeader "Content-Type", "application/json"
    http.setTimeouts 3000, 3000, 10000, 10000
    http.send body
    On Error GoTo 0
End Sub
