Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Debug.Print "[getLicense] DynamicFunc basladi. Kitap: " & targetWb.Name

    Dim mac As String
    Dim firmaAdi As String
    Dim userAdi As String
    Dim baseUrl As String

    mac = GetFirstMACAddress()
    Debug.Print "[getLicense] MAC: " & mac

    If Len(mac) < 10 Or Left(mac, 5) = "HATA:" Or mac = "MAC_BULUNAMADI" Then
        Debug.Print "[getLicense] Gecersiz MAC, cikiliyor."
        Set DynamicFunc = Nothing
        Exit Function
    End If

    ' Registry'den firma ve kullanıcı bilgilerini oku
    firmaAdi = Trim$(GetSetting("ilhan", "Settings", "mdip", "EPRON"))
    If Len(firmaAdi) = 0 Then firmaAdi = "EPRON"

    userAdi = Trim$(GetSetting("ilhan", "Settings", "TBveren", ""))

    Debug.Print "[getLicense] firmaAdi: " & firmaAdi
    Debug.Print "[getLicense] userAdi: " & userAdi

    baseUrl = ResolveBaseUrl(param)
    Call CheckAndRegisterLicense(mac, firmaAdi, userAdi, baseUrl)

    Set DynamicFunc = Nothing
End Function

Private Sub CheckAndRegisterLicense(mac As String, firmaAdi As String, userAdi As String, baseUrl As String)
    Dim http As Object
    Dim getUrl As String
    Dim postUrl As String
    Dim jsonBody As String

    getUrl = baseUrl & "license/" & mac & "/"
    Debug.Print "[getLicense] GET: " & getUrl

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    On Error GoTo ErrHandler

    http.Open "GET", getUrl, False
    http.setTimeouts 5000, 10000, 30000, 30000
    http.send

    Debug.Print "[getLicense] GET Status: " & http.Status

    If http.Status = 200 Then
        Debug.Print "[getLicense] Lisans mevcut, guncelleniyor."
        ' Mevcut kaydi firmaAdi/userAdi ile guncelle
        postUrl = baseUrl & "license/"
        jsonBody = BuildLicenseJson(mac, firmaAdi, userAdi)
        http.Open "POST", postUrl, False
        http.setTimeouts 5000, 10000, 30000, 30000
        http.setRequestHeader "Content-Type", "application/json"
        http.send jsonBody
        Debug.Print "[getLicense] POST (guncelle) Status: " & http.Status
        If http.Status = 200 Or http.Status = 201 Then
            SaveLicenseFromResponse http.responseText
        End If
        Set http = Nothing
        Exit Sub
    End If

    ' Lisans yok, yeni kayit olustur
    postUrl = baseUrl & "license/"
    jsonBody = BuildLicenseJson(mac, firmaAdi, userAdi)

    Debug.Print "[getLicense] POST (yeni): " & postUrl
    Debug.Print "[getLicense] Body: " & jsonBody

    http.Open "POST", postUrl, False
    http.setTimeouts 5000, 10000, 30000, 30000
    http.setRequestHeader "Content-Type", "application/json"
    http.send jsonBody

    Debug.Print "[getLicense] POST Status: " & http.Status

    If http.Status = 200 Or http.Status = 201 Then
        Debug.Print "[getLicense] Lisans kaydedildi."
        SaveLicenseFromResponse http.responseText
    Else
        Debug.Print "[getLicense] Lisans kayit hatasi: " & http.responseText
    End If

    Set http = Nothing
    Exit Sub

ErrHandler:
    Debug.Print "[getLicense] Hata: " & Err.Description
    Set http = Nothing
End Sub

Private Function BuildLicenseJson(mac As String, firmaAdi As String, userAdi As String) As String
    BuildLicenseJson = "{" & _
        """macAdresi"":""" & EscapeJson(mac) & """," & _
        """firmaAdi"":""" & EscapeJson(firmaAdi) & """," & _
        """userAdi"":""" & EscapeJson(userAdi) & """" & _
        "}"
End Function

Private Sub SaveLicenseFromResponse(responseText As String)
    Dim licenseValue As String

    licenseValue = ExtractJsonValue(responseText, "license")
    If Len(licenseValue) = 0 Then Exit Sub

    SaveSetting "ilhan", "Settings", "license", licenseValue
    SaveSetting "scngnr", "Settings", "license", licenseValue
    Debug.Print "[getLicense] License registry'e kaydedildi: " & licenseValue
End Sub

Private Function ExtractJsonValue(jsonText As String, key As String) As String
    Dim searchKey As String
    Dim p1 As Long
    Dim p2 As Long

    searchKey = """" & key & """:"
    p1 = InStr(1, jsonText, searchKey, vbTextCompare)
    If p1 = 0 Then Exit Function

    p1 = p1 + Len(searchKey)
    ' Bosluk atla
    Do While p1 <= Len(jsonText) And Mid$(jsonText, p1, 1) = " "
        p1 = p1 + 1
    Loop

    If Mid$(jsonText, p1, 1) = """" Then
        ' String deger
        p1 = p1 + 1
        p2 = InStr(p1, jsonText, """")
        If p2 > p1 Then
            ExtractJsonValue = Mid$(jsonText, p1, p2 - p1)
        End If
    Else
        ' Sayisal veya boolean deger
        p2 = p1
        Do While p2 <= Len(jsonText)
            Dim ch As String
            ch = Mid$(jsonText, p2, 1)
            If ch = "," Or ch = "}" Or ch = "]" Or ch = " " Then Exit Do
            p2 = p2 + 1
        Loop
        ExtractJsonValue = Mid$(jsonText, p1, p2 - p1)
    End If
End Function

Private Function EscapeJson(s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    EscapeJson = s
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

Private Function ResolveBaseUrl(apiBaseUrl As Variant) As String
    Dim url As String

    If Not IsMissing(apiBaseUrl) Then
        url = Trim(CStr(apiBaseUrl))
    End If

    If Len(url) = 0 Then url = "http://localhost:3000/api/"
    If Right(url, 1) <> "/" Then url = url & "/"
    ResolveBaseUrl = url
End Function
