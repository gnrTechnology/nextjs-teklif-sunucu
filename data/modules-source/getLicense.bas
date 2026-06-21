Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Debug.Print "[getLicense] DynamicFunc basladi. Kitap: " & targetWb.Name

    Dim mac As String
    Dim firmaAdi As String
    Dim userAdi As String
    Dim dosyaAdi As String
    Dim dosyaYolu As String
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

    ' IHLAl TESPİTI - dosyaAdi belirleme stratejisi:
    '
    ' 1. ONCELIK: Workbook_Open'in Registry'e yazdigi "startingAddin" anahtari.
    '    Workbook_Open'da "SaveSetting "ilhan","Settings","startingAddin",ThisWorkbook.FullName"
    '    satiri olmalidir. Kopya dosya da ayni kodu tasidigi icin otomatik calisir.
    '    Bu sayede Application.Run hangi zInternet'i bulursa bulsun doğru xlam bilinir.
    '
    ' 2. FALLBACK: targetWb.Name (ExecuteDynamicFunction'in pasladigi workbook).
    '    teklif.xlam aktifken bu her zaman "teklif.xlam" donebilir; bu nedenle
    '    registry onceliklıdir.
    Dim startingAddinPath As String
    startingAddinPath = Trim(GetSetting("ilhan", "Settings", "startingAddin", ""))

    If Len(startingAddinPath) > 0 Then
        ' Registry'den tam yol alindi → dosya adini cikart
        Dim sepIdx As Long
        Dim ci As Long
        sepIdx = 0
        For ci = Len(startingAddinPath) To 1 Step -1
            If Mid(startingAddinPath, ci, 1) = "\" Or Mid(startingAddinPath, ci, 1) = "/" Then
                sepIdx = ci
                Exit For
            End If
        Next ci
        dosyaYolu = startingAddinPath
        If sepIdx > 0 Then
            dosyaAdi = Mid(startingAddinPath, sepIdx + 1)
        Else
            dosyaAdi = startingAddinPath
        End If
        Debug.Print "[getLicense] startingAddin registry: " & dosyaAdi
    Else
        ' Fallback: targetWb
        On Error Resume Next
        dosyaAdi = targetWb.Name
        dosyaYolu = targetWb.FullName
        On Error GoTo 0
        If Len(dosyaAdi) = 0 Then dosyaAdi = "bilinmiyor.xlam"
        Debug.Print "[getLicense] Fallback targetWb: " & dosyaAdi
    End If

    Debug.Print "[getLicense] firmaAdi: " & firmaAdi
    Debug.Print "[getLicense] userAdi:  " & userAdi
    Debug.Print "[getLicense] dosyaAdi: " & dosyaAdi

    baseUrl = ResolveBaseUrl(param)

    ' Tek POST ile kayit olustur veya guncelle; sunucu mevcut license degerini korur
    Call RegisterOrUpdate(mac, firmaAdi, userAdi, dosyaAdi, dosyaYolu, baseUrl)

    Set DynamicFunc = Nothing
End Function

Private Sub RegisterOrUpdate(mac As String, firmaAdi As String, userAdi As String, _
                              dosyaAdi As String, dosyaYolu As String, baseUrl As String)
    Dim http As Object
    Dim postUrl As String
    Dim jsonBody As String

    postUrl = baseUrl & "license/"
    jsonBody = BuildLicenseJson(mac, firmaAdi, userAdi, dosyaAdi)

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
        Case 200, 201
            Debug.Print "[getLicense] Sunucu yaniti alindi."
            Call HandleLicenseResponse(http.responseText, dosyaYolu, baseUrl)
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

' Sunucu yanitini degerlendirip ihlal veya normal akim yonlendirir
Private Sub HandleLicenseResponse(responseText As String, dosyaYolu As String, baseUrl As String)
    Dim isIhlal As Boolean
    Dim kopyaAdi As String

    isIhlal = (InStr(1, responseText, """ihlal"":true", vbTextCompare) > 0)

    If isIhlal Then
        Debug.Print "[getLicense] IHLAL TESPIT EDILDI!"

        ' Kopya dosyanin tam yolunu registry'ye kaydet - ihlal.xlsm bunu okuyacak
        SaveSetting "ilhan", "Settings", "ihlalDosyaYolu", dosyaYolu
        kopyaAdi = ExtractJsonValue(responseText, "kopyaDosyaAdi")
        If Len(kopyaAdi) > 0 Then
            SaveSetting "ilhan", "Settings", "ihlalDosyaAdi", kopyaAdi
        End If

        Debug.Print "[getLicense] Registry guncellendi -> ihlalDosyaYolu=" & dosyaYolu

        ' Arkaplanda VBScript ile dosyalari sil (Excel kapaninca kilitsiz kalir)
        Call RunIhlalCleanup(dosyaYolu)
    Else
        Call SaveLicenseFromResponse(responseText)
    End If
End Sub

' Kopya dosyayi ve teklif.xlam'i silmek icin arka planda VBScript olusturur ve calistirir.
' VBScript, Excel COM otomasyonu ile xlam workbook'lari kilitli kapatir sonra dosyalari siler.
' kopyaYolu: silinecek kopya dosyanin tam yolu
Private Sub RunIhlalCleanup(kopyaYolu As String)
    Dim teklifYolu As String
    Dim vbsPath As String
    Dim fileNum As Integer

    teklifYolu = Environ("APPDATA") & "\Microsoft\AddIns\teklif.xlam"
    vbsPath = Environ("TEMP") & "\ihlal_cleanup.vbs"

    Debug.Print "[getLicense] Silme VBScript olusturuluyor: " & vbsPath

    fileNum = FreeFile
    Open vbsPath For Output As #fileNum
    Print #fileNum, "Dim fso : Set fso = CreateObject(""Scripting.FileSystemObject"")"
    Print #fileNum, ""
    Print #fileNum, "' --- 1. ADIM: Excel COM ile xlam workbook'lari kapat ---"
    Print #fileNum, "WScript.Sleep 2000"
    Print #fileNum, "On Error Resume Next"
    Print #fileNum, "Dim xlApp"
    Print #fileNum, "Set xlApp = GetObject(, ""Excel.Application"")"
    Print #fileNum, "If Not IsEmpty(xlApp) And Not IsNull(xlApp) Then"
    Print #fileNum, "    Dim wb"
    Print #fileNum, "    For Each wb In xlApp.Workbooks"
    Print #fileNum, "        If LCase(Right(wb.Name, 5)) = "".xlam"" Then"
    Print #fileNum, "            wb.Saved = True"
    Print #fileNum, "            wb.Close False"
    Print #fileNum, "        End If"
    Print #fileNum, "    Next"
    Print #fileNum, "End If"
    Print #fileNum, "Set xlApp = Nothing"
    Print #fileNum, "On Error GoTo 0"
    Print #fileNum, ""
    Print #fileNum, "' --- 2. ADIM: Dosyalari sil (max 10 deneme) ---"
    Print #fileNum, "Dim attempts : attempts = 0"
    Print #fileNum, "Do While attempts < 10"
    Print #fileNum, "    WScript.Sleep 3000"
    Print #fileNum, "    On Error Resume Next"
    Print #fileNum, "    fso.DeleteFile """ & kopyaYolu & """, True"
    Print #fileNum, "    fso.DeleteFile """ & teklifYolu & """, True"
    Print #fileNum, "    On Error GoTo 0"
    Print #fileNum, "    If Not fso.FileExists(""" & kopyaYolu & """) And Not fso.FileExists(""" & teklifYolu & """) Then"
    Print #fileNum, "        Exit Do"
    Print #fileNum, "    End If"
    Print #fileNum, "    attempts = attempts + 1"
    Print #fileNum, "Loop"
    Print #fileNum, ""
    Print #fileNum, "' --- 3. ADIM: Kendini sil ---"
    Print #fileNum, "On Error Resume Next"
    Print #fileNum, "fso.DeleteFile WScript.ScriptFullName, True"
    Close #fileNum

    Debug.Print "[getLicense] VBScript kaydedildi. Arkaplanda baslatiliyor..."
    Shell "wscript.exe """ & vbsPath & """ //nologo", vbHide
    Debug.Print "[getLicense] Silme scripti baslatildi."
End Sub

Private Function BuildLicenseJson(mac As String, firmaAdi As String, userAdi As String, dosyaAdi As String) As String
    BuildLicenseJson = "{" & _
        """macAdresi"":""" & EscapeJson(mac) & """," & _
        """firmaAdi"":""" & EscapeJson(firmaAdi) & """," & _
        """userAdi"":""" & EscapeJson(userAdi) & """," & _
        """dosyaAdi"":""" & EscapeJson(dosyaAdi) & """" & _
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

    ' Registry kaydından sonra Custom Tab durumunu guncelle
    Call UpdateCustomTab(licenseValue)
End Sub

' Lisans durumuna gore Custom Tab'i goster veya gizle.
' Application.Run ana workbook'taki Module1'i arar; bulamazsa sessizce gecilir.
Private Sub UpdateCustomTab(licenseValue As String)
    Dim lv As String
    lv = LCase(Trim(licenseValue))
    Dim isLicensed As Boolean
    isLicensed = (lv = "true" Or lv = "1" Or lv = "active" Or lv = "evet")

    On Error Resume Next
    If isLicensed Then
        Debug.Print "[getLicense] Lisans AKTIF -> Module1.ShowCustomTab cagriliyor"
        Application.Run "Module1.ShowCustomTab"
    Else
        Debug.Print "[getLicense] Lisans PASIF -> Module1.HideCustomTab cagriliyor"
        Application.Run "Module1.HideCustomTab"
    End If
    If Err.Number <> 0 Then
        Debug.Print "[getLicense] Tab guncelleme hatasi (sorun degil): " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
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
