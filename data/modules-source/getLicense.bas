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
        MsgBox "MAC adresi alınamadı. Lisans kaydı yapılamıyor.", vbExclamation, "getLicense"
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

    ' Tek POST ile kayit olustur veya guncelle; sunucu mevcut license degerini korur
    Call RegisterOrUpdate(mac, firmaAdi, userAdi, baseUrl)

    Set DynamicFunc = Nothing
End Function

Private Sub RegisterOrUpdate(mac As String, firmaAdi As String, userAdi As String, baseUrl As String)
    Dim http As Object
    Dim postUrl As String
    Dim jsonBody As String

    postUrl = baseUrl & "license/"
    jsonBody = BuildLicenseJson(mac, firmaAdi, userAdi)

    Debug.Print "[getLicense] POST: " & postUrl
    Debug.Print "[getLicense] Body: " & jsonBody

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    On Error GoTo ErrHandler

    http.Open "POST", postUrl, False
    http.setTimeouts 5000, 10000, 30000, 30000
    http.setRequestHeader "Content-Type", "application/json"
    http.send jsonBody

    Debug.Print "[getLicense] HTTP Status: " & http.Status

    Select Case http.Status
        Case 200
            Debug.Print "[getLicense] Mevcut lisans guncellendi."
            SaveLicenseFromResponse http.responseText
        Case 201
            Debug.Print "[getLicense] Yeni lisans kaydi olusturuldu."
            SaveLicenseFromResponse http.responseText
        Case Else
            Debug.Print "[getLicense] Sunucu hatasi (" & http.Status & "): " & http.responseText
            MsgBox "Lisans sunucusuna ulasılamadı." & vbCrLf & _
                   "HTTP " & http.Status & ": " & http.responseText, vbExclamation, "getLicense"
    End Select

    Set http = Nothing
    Exit Sub

ErrHandler:
    Debug.Print "[getLicense] Baglanti hatasi: " & Err.Description
    MsgBox "Bağlantı hatası: " & Err.Description, vbCritical, "getLicense"
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
    ' Sunucudan gelen license degerini (true/false) registry'e yaz
    Dim licenseValue As String

    licenseValue = ExtractJsonValue(responseText, "license")
    If Len(licenseValue) = 0 Then
        Debug.Print "[getLicense] Response'ta license degeri bulunamadi."
        Exit Sub
    End If

    SaveSetting "ilhan", "Settings", "license", licenseValue
    SaveSetting "scngnr", "Settings", "license", licenseValue
    Debug.Print "[getLicense] Registry guncellendi -> license=" & licenseValue
End Sub

' "key":"value" veya "key":booleanvalue formatlari icin
Private Function ExtractJsonValue(jsonText As String, key As String) As String
    Dim searchKey As String
    Dim p1 As Long
    Dim p2 As Long
    Dim ch As String

    searchKey = """" & key & """:"
    p1 = InStr(1, jsonText, searchKey, vbTextCompare)
    If p1 = 0 Then Exit Function

    p1 = p1 + Len(searchKey)

    ' Bosluk atla
    Do While p1 <= Len(jsonText) And Mid$(jsonText, p1, 1) = " "
        p1 = p1 + 1
    Loop

    If Mid$(jsonText, p1, 1) = """" Then
        ' Tirnak icindeki string
        p1 = p1 + 1
        p2 = InStr(p1, jsonText, """")
        If p2 > p1 Then
            ExtractJsonValue = Mid$(jsonText, p1, p2 - p1)
        End If
    Else
        ' Tirnak olmayan deger (true, false, sayi)
        p2 = p1
        Do While p2 <= Len(jsonText)
            ch = Mid$(jsonText, p2, 1)
            If ch = "," Or ch = "}" Or ch = "]" Or ch = " " Or ch = vbCr Or ch = vbLf Then
                Exit Do
            End If
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
