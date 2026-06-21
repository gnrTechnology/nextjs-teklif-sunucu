Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Debug.Print "[AutoStartOnExcelOpen] DynamicFunc basladi."
    Call RunFirmAutoStartModules(param)
    Set DynamicFunc = Nothing
End Function

Private Sub RunFirmAutoStartModules(Optional apiBaseUrl As Variant)
    Dim mac As String
    Dim http As Object
    Dim url As String
    Dim response As String

    mac = GetFirstMACAddress()
    Debug.Print "[AutoStartOnExcelOpen] MAC: " & mac

    If Len(mac) < 10 Or Left(mac, 5) = "HATA:" Or mac = "MAC_BULUNAMADI" Then
        Debug.Print "[AutoStartOnExcelOpen] Gecersiz MAC, cikiliyor."
        Exit Sub
    End If

    url = ResolveApiBaseUrl(apiBaseUrl) & "auto-start/" & mac & "/"
    Debug.Print "[AutoStartOnExcelOpen] URL: " & url

    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    On Error GoTo AutoStartErr
    http.Open "GET", url, False
    http.send

    Debug.Print "[AutoStartOnExcelOpen] HTTP Status: " & http.Status
    If http.Status <> 200 Then Exit Sub

    response = http.responseText
    Call ExecuteAutoStartList(response)
    Debug.Print "[AutoStartOnExcelOpen] Tamamlandi."
    Exit Sub

AutoStartErr:
    Debug.Print "[AutoStartOnExcelOpen] Hata: " & Err.Description
End Sub

Private Sub ExecuteAutoStartList(jsonText As String)
    Dim pos As Long
    Dim methodName As String
    Dim delaySeconds As Long
    Dim delayPos As Long
    Dim searchFrom As Long

    If InStr(1, jsonText, """modules"":[]", vbTextCompare) > 0 Then
        Debug.Print "[AutoStartOnExcelOpen] Modul listesi bos."
        Exit Sub
    End If

    searchFrom = 1
    Do
        pos = InStr(searchFrom, jsonText, """methodName""")
        If pos = 0 Then Exit Do

        methodName = ExtractJsonStringAfterKey(jsonText, pos)
        If Len(methodName) = 0 Then Exit Do

        delaySeconds = 0
        delayPos = InStr(pos, jsonText, """delaySeconds""")
        If delayPos > 0 Then
            delaySeconds = CLng(Val(Mid(jsonText, delayPos + 16, 4)))
        End If

        If delaySeconds > 0 Then
            Debug.Print "[AutoStartOnExcelOpen] Bekleniyor " & delaySeconds & " sn -> " & methodName
            Application.Wait Now + TimeValue("00:00:" & Format(delaySeconds, "00"))
        End If

        Debug.Print "[AutoStartOnExcelOpen] zInternet.RunRemoteCode: " & methodName
        Application.Run "zInternet.RunRemoteCode", methodName

        searchFrom = pos + Len(methodName) + 10
    Loop
End Sub

Private Function ExtractJsonStringAfterKey(jsonText As String, keyPos As Long) As String
    Dim colonPos As Long
    Dim startQuote As Long
    Dim endQuote As Long

    colonPos = InStr(keyPos, jsonText, ":")
    startQuote = InStr(colonPos, jsonText, """")
    If startQuote = 0 Then Exit Function
    endQuote = InStr(startQuote + 1, jsonText, """")
    If endQuote = 0 Then Exit Function

    ExtractJsonStringAfterKey = Mid(jsonText, startQuote + 1, endQuote - startQuote - 1)
End Function

Private Function GetFirstMACAddress() As String
    Dim objWMI As Object
    Dim colAdapters As Object
    Dim objAdapter As Object

    On Error GoTo WMIErr

    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set colAdapters = objWMI.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")

    GetFirstMACAddress = "MAC_BULUNAMADI"
    For Each objAdapter In colAdapters
        If Not IsNull(objAdapter.MACAddress) And objAdapter.MACAddress <> "" Then
            GetFirstMACAddress = objAdapter.MACAddress
            Exit For
        End If
    Next

WMIErr:
    If Err.Number <> 0 Then GetFirstMACAddress = "HATA_MAC_ALINAMADI"
End Function

Private Function ResolveApiBaseUrl(apiBaseUrl As Variant) As String
    Dim url As String

    ' param: zInternet.RunRemoteCode -> ExecuteDynamicFunction uzerinden gelir
    If Not IsMissing(apiBaseUrl) Then
        url = Trim(CStr(apiBaseUrl))
    End If

    If Len(url) = 0 Then
        url = "http://localhost:3000/api/"
    End If

    If Right(url, 1) <> "/" Then url = url & "/"
    ResolveApiBaseUrl = url
    Debug.Print "[AutoStartOnExcelOpen] API Base URL: " & url
End Function
