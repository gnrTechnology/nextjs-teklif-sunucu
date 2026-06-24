Option Explicit

' Gecikmeli bağlama (late binding) için HTTP nesnesini tanımlıyoruz.
Private http As Object

' --- URL Sabitleri ---
' GET isteği (lisans sorgulama) için URL
Public Const GET_LICENSE_URL As String = "https://nextjs-teklif-sunucu.vercel.app/api/"
' POST isteği (yeni kayıt) için URL
Public Const POST_LICENSE_URL As String = "https://nextjs-teklif-sunucu.vercel.app/api/license/"


' --- 1. İNTERNET BAĞLANTI KONTROLÜ ---

' Windows API fonksiyonunu VBA'da kullanmak için bildiriyoruz.
#If VBA7 Then
    Public Declare PtrSafe Function InternetGetConnectedState Lib "wininet.dll" ( _
        ByRef lpdwFlags As Long, _
        ByVal dwReserved As Long) As Boolean
#Else
    Public Declare Function InternetGetConnectedState Lib "wininet.dll" ( _
        ByRef lpdwFlags As Long, _
        ByVal dwReserved As Long) As Boolean
#End If

'''
' Aktif bir internet bağlantısı olup olmadığını kontrol eden ana fonksiyon
'
Public Function IsInternetConnected() As Boolean
    Dim lngFlags As Long
    IsInternetConnected = InternetGetConnectedState(lngFlags, 0&)
End Function


' --- 2. DONANIM BİLGİLERİNİ ALMA (WMI) ---

'''
' Aktif, IP etkin ağ bağdaştırıcısının ilk MAC adresini alır.
'
Private Function GetFirstMACAddress() As String
    Dim objWMI As Object
    Dim colAdapters As Object
    Dim objAdapter As Object
    Dim strQuery As String
    
    On Error GoTo WMIErr
    
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    strQuery = "SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True"
    Set colAdapters = objWMI.ExecQuery(strQuery)
    
    GetFirstMACAddress = "MAC_BULUNAMADI"
    For Each objAdapter In colAdapters
        ' *** KRİTİK DÜZELTME: .macAddress -> .MACAddress (Büyük harf) ***
        If Not IsNull(objAdapter.macAddress) And objAdapter.macAddress <> "" Then
            GetFirstMACAddress = objAdapter.macAddress
            Exit For ' Sadece ilkini alıp çıkıyoruz
        End If
    Next
    
WMIErr:
    If Err.Number <> 0 Then GetFirstMACAddress = "HATA_MAC_ALINAMADI"
    Set objAdapter = Nothing
    Set colAdapters = Nothing
    Set objWMI = Nothing
End Function

'''
' Aktif, IP etkin ağ bağdaştırıcısının ilk IPv4 adresini alır.
'
Private Function GetFirstIPAddress() As String
    Dim objWMI As Object
    Dim colAdapters As Object
    Dim objAdapter As Object
    Dim strQuery As String
    
    On Error GoTo WMIErr
    
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    strQuery = "SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True"
    Set colAdapters = objWMI.ExecQuery(strQuery)
    
    GetFirstIPAddress = "IP_BULUNAMADI"
    For Each objAdapter In colAdapters
         ' *** KRİTİK DÜZELTME: .ipAddress -> .IPAddress (Büyük harf) ***
        If Not IsNull(objAdapter.ipAddress) Then
            ' IPAddress bir dizidir (array), ilkini (genellikle IPv4) alıyoruz
            If UBound(objAdapter.ipAddress) >= 0 Then
                GetFirstIPAddress = objAdapter.ipAddress(0)
                Exit For
            End If
        End If
    Next

WMIErr:
    If Err.Number <> 0 Then GetFirstIPAddress = "HATA_IP_ALINAMADI"
    Set objAdapter = Nothing
    Set colAdapters = Nothing
    Set objWMI = Nothing
End Function


' --- 3. SUNUCU İLE İLETİŞİM (GET VE POST) ---

'''
' Belirtilen sunucuya MAC adresi ile lisans sorgusu yapar (GET isteği)
'
Public Function GetLicenseStatus() As String
    Dim macAddress As String
    Dim requestUrl As String
    
    macAddress = GetFirstMACAddress()
    If macAddress = "MAC_BULUNAMADI" Or macAddress = "HATA_MAC_ALINAMADI" Then
        GetLicenseStatus = "HATA: Gecerli bir MAC adresi bulunamadi."
        Exit Function
    End If
    
    requestUrl = GET_LICENSE_URL & "license/" & macAddress & "/"
    
    On Error GoTo HttpErr
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", requestUrl, False
    http.setTimeouts 5000, 10000, 30000, 30000
    http.send
    
    If http.Status = 200 Then
        GetLicenseStatus = http.responseText
    ElseIf http.Status = 404 Then
        ' 404 = MAC bulunamadi, hata degil — success:false dondurelim
        GetLicenseStatus = "{""success"":false,""data"":null}"
    Else
        GetLicenseStatus = "HATA: Status=" & http.Status & " - " & http.statusText
    End If
    Set http = Nothing
    Exit Function
    
HttpErr:
    GetLicenseStatus = "HATA: Baglanilmadi. " & Err.Description
    Set http = Nothing
End Function
'''
' Sunucuya yeni lisans bilgisi kaydı yapar (POST isteği).
'
Private Function RegisterLicense(macAddress As String, ipAddress As String) As String
    Dim httpPOST As Object
    Dim jsonPayload As String
    Dim firmaAdi As String
    Dim userAdi As String
    
    firmaAdi = Trim(GetSetting("ilhan", "Settings", "mdip", "EPRON"))
    If Len(firmaAdi) = 0 Then firmaAdi = "EPRON"
    userAdi = Trim(GetSetting("ilhan", "Settings", "TBveren", ""))
    
    jsonPayload = "{" & _
        """firmaAdi"":""" & firmaAdi & """," & _
        """userAdi"":""" & userAdi & """," & _
        """ipAdresi"":""" & ipAddress & """," & _
        """macAdresi"":""" & macAddress & """," & _
        """dosyaAdi"":""" & ThisWorkbook.Name & """" & _
    "}"
    
    On Error GoTo PostErr
    Set httpPOST = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    httpPOST.Open "POST", POST_LICENSE_URL, False
    httpPOST.setTimeouts 5000, 10000, 30000, 30000
    httpPOST.setRequestHeader "Content-Type", "application/json;charset=UTF-8"
    httpPOST.send jsonPayload
    
    If httpPOST.Status = 201 Or httpPOST.Status = 200 Then
        RegisterLicense = httpPOST.responseText
    Else
        RegisterLicense = "HATA: Status=" & httpPOST.Status & " - " & httpPOST.responseText
    End If
    Set httpPOST = Nothing
    Exit Function

PostErr:
    RegisterLicense = "HATA: POST AG HATASI: " & Err.Description
    Set httpPOST = Nothing
End Function


Sub TestLicenseCheck()

    Debug.Print "----------------------------------------------------------------"
    Debug.Print "Sunucuya lisans sorgusu (GET) gönderiliyor..."
    
    Dim sunucuYaniti As String
    sunucuYaniti = GetLicenseStatus() ' Ana GET fonksiyonunu çağır
    
    ' Önce hatayı kontrol et
    If Left(sunucuYaniti, 5) = "HATA:" Then
        Debug.Print "Sunucudan Gelen Yanıt:" & vbCrLf & sunucuYaniti
        Exit Sub
    End If
    
    ' Hata yoksa, JSON yanıtını işle
    On Error GoTo JsonErr
    
    ' --- GÜNCELLEME: JsonConverter DÜZELTMESİ ---
    ' "JsonConverter" bir Standart Modül olduğu için "New" kullanılmaz.
    ' "parser" nesnesi kaldırıldı.
    Dim jsonResponse As Dictionary
    ' Fonksiyonu doğrudan Modül Adı ile çağırıyoruz:
    Set jsonResponse = JsonConverter.ParseJson(sunucuYaniti)
    ' ------------------------------------------
    
    Dim success As Boolean
    success = CBool(jsonResponse("success")) ' CBool ile Boolean'a dönüştür
    
    ' Şimdi "success" değerini kontrol et
    If success = False Then
        ' --- LİSANS YOK (success:false), YENİ KAYIT YAP (POST) ---
        Debug.Print "Sunucudan Gelen Yanıt (GET): " & sunucuYaniti
        Debug.Print "Lisans bulunamadı. Yeni kayıt (POST) işlemi başlatılıyor..."
        
        Dim mac As String
        Dim ip As String
        mac = GetFirstMACAddress()
        ip = GetFirstIPAddress()
        
        If Left(mac, 5) = "HATA:" Or Left(ip, 5) = "HATA:" Then
            Debug.Print "Kayıt BAŞARISIZ: POST için MAC (" & mac & ") veya IP (" & ip & ") alınamadı."
        Else
            ' --- YENİ KAYIT İŞLEMİ ---
            Dim postYaniti As String
            postYaniti = RegisterLicense(mac, ip) ' Yeni POST fonksiyonunu çağır
            
            ' Dönen yanıt HATA ile başlıyor mu?
            If Left(postYaniti, 5) = "HATA:" Then
                Debug.Print "Sunucudan Gelen Yanıt (POST): " & vbCrLf & postYaniti
            Else
                ' --- YENİ KAYIT BAŞARILI, KAYIT DEFTERİNE YAZ ---
                Debug.Print "Sunucudan Gelen Yanıt (POST): " & vbCrLf & postYaniti
                
                ' --- GÜNCELLEME: POST YANITINI DA PARSE ET ---
                Dim postJsonResponse As Dictionary
                ' Fonksiyonu doğrudan Modül Adı ile çağırıyoruz:
                Set postJsonResponse = JsonConverter.ParseJson(postYaniti)
                
                Dim licenseValue As String
                ' Yanıt {"data": {...}} şeklinde geldiği için
                licenseValue = CStr(postJsonResponse("data")("license"))
                ' --------------------------------------------------
                
                Debug.Print "Yeni kayıt başarılı. Lisans durumu (" & licenseValue & ") kayıt defterine kaydediliyor..."
                
                ' zLicense modülündeki fonksiyonu çağır
                Call SaveLicenseToRegistry(licenseValue)
                
                Debug.Print "Lisans anahtarı kayıt defterine başarıyla kaydedildi."
            End If
        End If
        
    Else
        ' --- LİSANS VAR (success:true), KAYIT DEFTERİNE YAZ ---
        Debug.Print "Sunucudan Gelen Yanıt (GET): " & vbCrLf & sunucuYaniti
        
        ' --- GÜNCELLEME: Sadece data.license değerini al ---
        'Dim licenseValue As String
        ' Yanıt {"success":true, "data":{...}} şeklinde geldiği için
        licenseValue = CStr(jsonResponse("data")("license"))
        ' --------------------------------------------------
        
        Debug.Print "Lisans durumu: " & licenseValue & ". Anahtar kayıt defterine kaydediliyor..."
                
        ' zLicense modülündeki fonksiyonu çağır
        Call SaveLicenseToRegistry(licenseValue)
        
        Debug.Print "Lisans anahtarı kayıt defterine başarıyla kaydedildi."
    End If
    
    Set jsonResponse = Nothing
    Exit Sub
    
JsonErr:
    Debug.Print "JSON PARSE HATASI: Sunucudan geçersiz yanıt alındı. Yanıt: " & sunucuYaniti & " | Hata: " & Err.Description
    Set jsonResponse = Nothing
End Sub

' --- 4. ANA KONTROL AKIŞI (zLicense modülü ile entegre) ---

'---------------------------------------------------------------
' Lisans kontrolünü başlatan ana Sub
'---------------------------------------------------------------
Sub TestInternetConnection()
    If IsInternetConnected Then
        'Debug.Print "İnternet bağlantısı var."
        Call TestLicenseCheck
        
        If zLicense.GetLicenseFromRegistry Then
            Module1.ShowCustomTab
        Else
            Module1.HideCustomTab
        End If
    Else
        Debug.Print "İnternet bağlantısı yok."
    End If
End Sub


'########################################################################################
'Sunucuya Veri gönder
'########################################################################################

Sub PostDataToServer()
    Dim wb As Workbook
    Dim wsMain As Worksheet, wsDB As Worksheet
    Dim http As Object
    Dim jsonBody As String
    Dim registryFilePath As String
    Dim macAddress As String
    Dim veritabaniListesi As String
    Dim responseText As String
    
    ' Registry'den dosya yolunu kontrol et
    registryFilePath = GetSetting("sercan", "fileOpenWorkBooks", "nowOpenPropsFile", "")
    
    ' Aktif dosya üzerinde işlem yapıyoruz
    Set wb = ActiveWorkbook
    
    ' Sayfa tanımları
    On Error Resume Next
    Set wsMain = wb.Sheets("Sayfa3")
    Set wsDB = wb.Sheets("Sayfa1")
    On Error GoTo 0
    
    If wsMain Is Nothing Or wsDB Is Nothing Then
        'MsgBox "Gerekli sayfalar (Sayfa1 veya Sayfa3) bulunamadı!", vbCritical
        Exit Sub
    End If

    ' --- 2. VERİLERİN HAZIRLANMASI ---
    
    ' !!! DEĞİŞİKLİK YAPILAN KISIM BURASI !!!
    ' Sizin verdiğiniz yeni fonksiyonu çağırıyoruz:
    macAddress = GetFirstMACAddress()
    
    ' Veritabanı Listesini (Array) Oluştur (Sayfa1'den)
    ' Not: CreateVeritabaniJson fonksiyonunun modülde olduğundan emin olun
    veritabaniListesi = CreateVeritabaniJson(wsDB)
    
    ' --- 3. JSON GÖVDESİNİN OLUŞTURULMASI ---
    
    jsonBody = "{"
    jsonBody = jsonBody & """macAdresi"": """ & macAddress & ""","
    jsonBody = jsonBody & """dosyaAdi"": """ & CleanJson(wb.Name) & ""","
    
    ' Opsiyonel Veriler (Sayfa3'ten alınanlar)
    jsonBody = jsonBody & """firmaAdi"": """ & CleanJson(wsMain.Range("C5").Value) & ""","
    jsonBody = jsonBody & """projeAdi"": """ & CleanJson(wsMain.Range("C3").Value) & ""","
    jsonBody = jsonBody & """projeKisaAdresi"": """ & CleanJson(wsMain.Range("C4").Value) & ""","
    
    ' Para Birimleri
    jsonBody = jsonBody & """teklifParaBirimiUSD"": """ & CleanJson(wsMain.Range("D16").Value) & ""","
    jsonBody = jsonBody & """teklifParaBirimiEuro"": """ & CleanJson(wsMain.Range("D17").Value) & ""","
    jsonBody = jsonBody & """teklifParaBirimiGenel"": """ & "" & ""","
    
    ' Gider ve Kar
    jsonBody = jsonBody & """genelGider"": """ & CleanJson(wsMain.Range("I22").Value) & ""","
    jsonBody = jsonBody & """kar"": """ & CleanJson(wsMain.Range("I25").Value) & ""","
    
    ' M31 Değeri
    jsonBody = jsonBody & """m31Degeri"": """ & CleanJson(wsMain.Range("M31").Value) & ""","
    
    ' Veritabanı Teklif Array'i
    jsonBody = jsonBody & """veritabaniTeklif"": " & veritabaniListesi
    
    jsonBody = jsonBody & "}"

    ' --- 4. HTTP POST İSTEĞİ ---
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    With http
        .Open "POST", POST_LICENSE_URL, False
        .setRequestHeader "Content-Type", "application/json"
        .send jsonBody
        responseText = .responseText
    End With
    
    ' --- 5. SONUÇ İŞLEME ---
    
    If http.Status = 200 Or http.Status = 201 Then
        'MsgBox "Veriler başarıyla gönderildi." & vbNewLine & "Sunucu Yanıtı: " & responseText, vbInformation
    Else
        'MsgBox "Hata oluştu!" & vbNewLine & "Status: " & http.Status & vbNewLine & responseText, vbCritical
    End If

    Set http = Nothing
End Sub

' --- YARDIMCI FONKSİYONLAR (Hata almamak için gerekli) ---

Function CreateVeritabaniJson(ws As Worksheet) As String
    Dim lastRow As Long, i As Long
    Dim jsonArr As String
    
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    If lastRow < 2 Then
        CreateVeritabaniJson = "[]"
        Exit Function
    End If
    
    jsonArr = "["
    For i = 2 To lastRow
        If ws.Cells(i, "B").Value <> "" Then
            jsonArr = jsonArr & "{"
            jsonArr = jsonArr & """colB"": """ & CleanJson(ws.Cells(i, "B").Value) & ""","
            jsonArr = jsonArr & """colC"": """ & CleanJson(ws.Cells(i, "C").Value) & ""","
            jsonArr = jsonArr & """colD"": """ & CleanJson(ws.Cells(i, "D").Value) & ""","
            jsonArr = jsonArr & """colF"": """ & CleanJson(ws.Cells(i, "F").Value) & ""","
            jsonArr = jsonArr & """colG"": """ & CleanJson(ws.Cells(i, "G").Value) & ""","
            jsonArr = jsonArr & """colParaBirimi"": """ & CleanJson(ws.Cells(i, "E").Value) & """"
            jsonArr = jsonArr & "},"
        End If
    Next i
    
    If Right(jsonArr, 1) = "," Then jsonArr = Left(jsonArr, Len(jsonArr) - 1)
    jsonArr = jsonArr & "]"
    CreateVeritabaniJson = jsonArr
End Function

Function CleanJson(val As Variant) As String
    If IsError(val) Or IsNull(val) Then
        CleanJson = ""
    Else
        Dim Tmp As String
        Tmp = CStr(val)
        Tmp = Replace(Tmp, "\", "\\")
        Tmp = Replace(Tmp, """", "\""")
        Tmp = Replace(Tmp, vbCrLf, " ")
        CleanJson = Tmp
    End If
End Function



' NOT: Bu kodun çalışabilmesi için Excel Ayarlarından:
' Dosya > Seçenekler > Güven Merkezi > Güven Merkezi Ayarları > Makro Ayarları >
' "VBA projesi nesne modeline erişime güven" (Trust access to the VBA project object model)
' seçeneğinin İŞARETLİ olması gerekir.

' -------------------------------------------------------------------------
' 2. UZAKTAN KOD ÇALIŞTIRMA MOTORU
' -------------------------------------------------------------------------
' methodName : calistirilacak uzak modul adi
' extraParam  : (opsiyonel) DynamicFunc'a param olarak iletilir.
'               Belirtilmezse GET_LICENSE_URL gecer (mevcut davranis korunur).
'               Birden fazla deger icin JSON string kullanin:
'                 RunRemoteCode "Modul", "{""anahtar"":""deger"",""sayi"":42}"
' teklif.xlam içindeki zInternet modülüne ekleyin (veya mevcut kodu bununla değiştirin).
' Excel açılışı: ThisWorkbook › Auto_Open › RunRemoteCode "AutoStartOnExcelOpen"

' methodName : calistirilacak uzak modul adi
' extraParam  : (opsiyonel) DynamicFunc'a param olarak iletilir.
'               Belirtilmezse GET_LICENSE_URL gecer (mevcut davranis korunur).
'               Birden fazla deger icin JSON string kullanin:
'                 RunRemoteCode "Modul", "{""anahtar"":""deger"",""sayi"":42}"
' methodName : calistirilacak uzak modul adi
' extraParam  : (opsiyonel) DynamicFunc'a param olarak iletilir.
'               Belirtilmezse GET_LICENSE_URL gecer (mevcut davranis korunur).
'               Birden fazla deger icin JSON string kullanin:
'                 RunRemoteCode "Modul", "{""anahtar"":""deger"",""sayi"":42}"
Public Sub RunRemoteCode(methodName As String, Optional extraParam As Variant)
    RunRemoteCodeInternal methodName, extraParam, False
End Sub

Public Sub RunRemoteCodeQuiet(methodName As String, Optional extraParam As Variant)
    RunRemoteCodeInternal methodName, extraParam, True
End Sub

' ¦¦ Auto-start tek seferlik calistirma (registry: ilhan / AutoStart) ¦¦¦¦¦¦¦¦¦¦
Public Function IsAutoStartRunOnceDone(methodName As String) As Boolean
    Dim key As String
    key = "done_" & LCase$(Trim$(methodName))
    IsAutoStartRunOnceDone = (LCase$(GetSetting("ilhan", "AutoStart", key, "")) = "true")
End Function

Public Sub MarkAutoStartRunOnceDone(methodName As String)
    Dim key As String
    key = LCase$(Trim$(methodName))
    SaveSetting "ilhan", "AutoStart", "done_" & key, "true"
    SaveSetting "ilhan", "AutoStart", "doneAt_" & key, Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Sub

Public Sub ClearAutoStartRunOnce(methodName As String)
    Dim key As String
    key = LCase$(Trim$(methodName))
    On Error Resume Next
    DeleteSetting "ilhan", "AutoStart", "done_" & key
    DeleteSetting "ilhan", "AutoStart", "doneAt_" & key
    On Error GoTo 0
End Sub

Public Function ShouldRunAutoStartModule(methodName As String, runOnce As Boolean) As Boolean
    If Not runOnce Then
        ShouldRunAutoStartModule = True
    Else
        ShouldRunAutoStartModule = Not IsAutoStartRunOnceDone(methodName)
    End If
End Function

' Firma auto-start listesinden gelen modul — runOnce ise registry kontrolu yapar
Public Sub RunAutoStartModule(methodName As String, runOnce As Boolean)
    If Len(Trim$(methodName)) = 0 Then Exit Sub
    If LCase$(methodName) = "getlicense" Then Exit Sub

    If Not ShouldRunAutoStartModule(methodName, runOnce) Then
        Debug.Print "[zInternet] RunOnce atlandi: " & methodName
        Exit Sub
    End If

    On Error Resume Next
    Application.Run "zInternet.RunRemoteCodeQuiet", methodName
    If Err.Number <> 0 Then
        Debug.Print "[zInternet] RunRemoteCodeQuiet hatasi: " & Err.Description
        Err.Clear
        Application.Run "zInternet.RunRemoteCode", methodName
    End If

    If Err.Number = 0 Then
        If runOnce Then MarkAutoStartRunOnceDone methodName
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Function ExtractJsonBoolNear(jsonText As String, anchorPos As Long, keyName As String) As Boolean
    Dim p As Long
    Dim slice As String
    p = InStr(anchorPos, jsonText, """" & keyName & """")
    If p = 0 Or p > anchorPos + 400 Then Exit Function
    slice = Mid$(jsonText, p, 24)
    ExtractJsonBoolNear = (InStr(1, slice, "true", vbTextCompare) > 0)
End Function

' JSON auto-start listesini isler (AutoStartOnExcelOpen / getLicense ortak)
Public Sub ExecuteFirmAutoStartList(jsonText As String)
    Dim pos As Long
    Dim methodName As String
    Dim delaySeconds As Long
    Dim runOnce As Boolean
    Dim delayPos As Long
    Dim searchFrom As Long

    If InStr(1, jsonText, """modules"":[]", vbTextCompare) > 0 Then Exit Sub

    searchFrom = 1
    Do
        pos = InStr(searchFrom, jsonText, """methodName""")
        If pos = 0 Then Exit Do

        methodName = ExtractJsonStringNearKey(jsonText, pos)
        If Len(methodName) = 0 Then Exit Do

        delaySeconds = 0
        delayPos = InStr(pos, jsonText, """delaySeconds""")
        If delayPos > 0 And delayPos < pos + 400 Then
            delaySeconds = CLng(val(Mid$(jsonText, delayPos + 16, 6)))
        End If

        runOnce = ExtractJsonBoolNear(jsonText, pos, "runOnce")

        If delaySeconds > 0 Then
            Application.Wait Now + TimeValue("00:00:" & Format$(delaySeconds, "00"))
        End If

        RunAutoStartModule methodName, runOnce

        searchFrom = pos + Len(methodName) + 10
    Loop
End Sub

Private Function ExtractJsonStringNearKey(jsonText As String, keyPos As Long) As String
    Dim colonPos As Long
    Dim startQ As Long
    Dim endQ As Long
    colonPos = InStr(keyPos, jsonText, ":")
    If colonPos = 0 Then Exit Function
    startQ = InStr(colonPos, jsonText, """")
    If startQ = 0 Then Exit Function
    endQ = InStr(startQ + 1, jsonText, """")
    If endQ = 0 Then Exit Function
    ExtractJsonStringNearKey = Mid$(jsonText, startQ + 1, endQ - startQ - 1)
End Function

Private Sub RunRemoteCodeInternal(methodName As String, extraParam As Variant, quiet As Boolean)
    Dim http As Object
    Dim rawResponse As String
    Dim cleanVbaCode As String
    Dim jsonBody As String
    Dim hostWb As Workbook
    Dim apiUrl As String
    Dim dynParam As Variant

    Debug.Print "[zInternet] RunRemoteCode basladi. methodName: " & methodName

    apiUrl = GetSetting("ilhan", "Settings", "apiBaseUrl", GET_LICENSE_URL)
    If Len(Trim(apiUrl)) = 0 Then apiUrl = GET_LICENSE_URL
    If Right(apiUrl, 1) <> "/" Then apiUrl = apiUrl & "/"
    apiUrl = apiUrl & "module/"

    ' extraParam verilmediyse API URL'yi parametre olarak ilet (geriye donuk uyumluluk)
    If IsMissing(extraParam) Or IsEmpty(extraParam) Then
        If quiet Then
            dynParam = ""
        Else
            dynParam = GET_LICENSE_URL
        End If
    Else
        dynParam = extraParam
    End If

    jsonBody = "{""methodName"":""" & methodName & """}"
    Debug.Print "[zInternet] API URL: " & apiUrl
    Debug.Print "[zInternet] param: " & CStr(dynParam)

    Set hostWb = GetHostWorkbook(ActiveWorkbook)
    If hostWb Is Nothing Then
        Debug.Print "[zInternet] Ana dosya bulunamadi."
        If Not quiet Then MsgBox "Ana teklif dosyası bulunamadı.", vbCritical
        If quiet Then Err.Raise vbObjectError + 514, "zInternet", "Ana teklif dosyasi bulunamadi"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")

    On Error GoTo ErrHandler
    With http
        .Open "POST", apiUrl, False
        .setRequestHeader "Content-Type", "application/json;charset=UTF-8"
        .send jsonBody

        Debug.Print "[zInternet] HTTP Status: " & .Status

        If .Status = 200 Then
            rawResponse = .responseText
            cleanVbaCode = ExtractCodeFromJSON(rawResponse)
            Debug.Print "[zInternet] Kod uzunlugu: " & Len(cleanVbaCode)

            If Len(cleanVbaCode) > 0 Then
                Call ExecuteDynamicFunction(cleanVbaCode, hostWb, dynParam, quiet)
                If methodName = "HeartbeatPing" Or methodName = "InstallTeklifAgent" Then
                    On Error Resume Next
                    Application.OnTime Now + TimeValue("00:00:03"), "zInternet.EnsureCommandQueueQuiet"
                    On Error GoTo 0
                End If
            Else
                If Not quiet Then MsgBox "Sunucudan kod içeriği boş döndü.", vbExclamation
                If quiet Then Err.Raise vbObjectError + 515, "zInternet", "Sunucudan kod bos"
            End If
        Else
            If Not quiet Then MsgBox "Sunucu Hatası (" & .Status & "): " & .responseText, vbCritical
            If quiet Then Err.Raise vbObjectError + 516, "zInternet", "Sunucu hatasi " & .Status
        End If
    End With

    Set http = Nothing
    Application.ScreenUpdating = True
    Debug.Print "[zInternet] RunRemoteCode tamamlandi."
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Debug.Print "[zInternet] Baglanti hatasi: " & Err.Description
    If Not quiet Then MsgBox "Bağlantı Hatası: " & Err.Description, vbCritical
    Set http = Nothing
    If quiet Then Err.Raise Err.Number, "zInternet", Err.Description
End Sub

Public Function ExecuteDynamicFunction(codeContent As String, targetWb As Workbook, Optional param As Variant, Optional quiet As Boolean = False) As Object
    Dim tempWb As Workbook
    Dim vbComp As Object
    Dim modName As String
    Dim result As Object
    Dim fullCode As String

    Debug.Print "[zInternet] ExecuteDynamicFunction basladi. targetWb: " & targetWb.Name

    If IsMissing(param) Then param = ""

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set tempWb = Workbooks.Add

    On Error Resume Next
    tempWb.Windows(1).Visible = False
    On Error GoTo Cleanup

    Set vbComp = tempWb.VBProject.VBComponents.Add(1)
    modName = "TempMod"
    vbComp.Name = modName

    fullCode = PrepareModuleCode(codeContent)
    vbComp.CodeModule.AddFromString fullCode

    Debug.Print "[zInternet] DynamicFunc cagriliyor..."
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Interactive = True
    DoEvents
    On Error GoTo Cleanup
    Set result = Application.Run("'" & tempWb.Name & "'!" & modName & ".DynamicFunc", targetWb, param)
    Set ExecuteDynamicFunction = result

Cleanup:
    If Not tempWb Is Nothing Then
        tempWb.Close SaveChanges:=False
        Set tempWb = Nothing
    End If

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        Dim errNum As Long
        Dim errDesc As String
        errNum = Err.Number
        errDesc = Err.Description
        Debug.Print "[zInternet] ExecuteDynamicFunction hata: " & errDesc
        If Not quiet Then MsgBox "Uzak modul hatasi:" & vbCrLf & errDesc, vbCritical, "RunRemoteCode"
        Err.Clear
        Err.Raise errNum, "zInternet", errDesc
    End If
End Function

Private Function PrepareModuleCode(codeContent As String) As String
    Dim s As String

    s = codeContent
    Do While Len(s) > 0
        If Left$(s, 2) = vbCrLf Then
            s = Mid$(s, 3)
        ElseIf Left$(s, 1) = vbCr Or Left$(s, 1) = vbLf Then
            s = Mid$(s, 2)
        Else
            Exit Do
        End If
    Loop

    If StrComp(Left$(s, 14), "Option Explicit", vbTextCompare) = 0 Then
        s = Trim$(Mid$(s, 15))
    End If

    PrepareModuleCode = "Option Explicit" & vbCrLf & vbCrLf & s
End Function

Public Function ExtractCodeFromJSON(jsonText As String) As String
    Dim p1 As Long, p2 As Long
    Dim tempStr As String

    p1 = InStr(1, jsonText, """code""", vbTextCompare)
    If p1 = 0 Then
        ExtractCodeFromJSON = jsonText
        Exit Function
    End If

    p1 = InStr(p1, jsonText, ":")
    p1 = InStr(p1, jsonText, """") + 1
    p2 = InStrRev(jsonText, """")

    If p2 > p1 Then
        tempStr = Mid(jsonText, p1, p2 - p1)
        tempStr = Replace(tempStr, "\""", """")
        tempStr = Replace(tempStr, "\r\n", vbCrLf)
        tempStr = Replace(tempStr, "\n", vbCrLf)
        tempStr = Replace(tempStr, "\t", vbTab)
        tempStr = Replace(tempStr, "\\", "\")
        ExtractCodeFromJSON = tempStr
    Else
        ExtractCodeFromJSON = ""
    End If
End Function

Private Function GetHostWorkbook(Optional preferred As Workbook) As Workbook
    Dim wb As Workbook

    If Not preferred Is Nothing Then
        If Not preferred.IsAddin Then
            If InStr(1, preferred.Name, "TeklifPollHost", vbTextCompare) = 0 Then
                Set GetHostWorkbook = preferred
                Exit Function
            End If
        End If
    End If

    For Each wb In Application.Workbooks
        If Not wb.IsAddin Then
            If InStr(1, wb.Name, "TeklifPollHost", vbTextCompare) = 0 Then
                Set GetHostWorkbook = wb
                Exit Function
            End If
        End If
    Next wb
End Function

' Heartbeat sonrasi komut kuyrugunu sessizce kur / yenile
Public Sub EnsureCommandQueueQuiet()
    On Error Resume Next
    Application.Run "zInternet.RunRemoteCodeQuiet", "InstallCommandQueue"
    Err.Clear
    On Error GoTo 0
End Sub

' ¦¦ Oturum acilisinda firma auto-start zinciri (boot basina 1 kez) ¦¦¦¦¦¦¦¦¦¦¦
Public Function GetBootSessionId() As String
    On Error GoTo Fail
    Dim wmi As Object, col As Object, o As Object
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set col = wmi.ExecQuery("SELECT LastBootUpTime FROM Win32_OperatingSystem")
    For Each o In col
        GetBootSessionId = Replace(Replace(CStr(o.LastBootUpTime), ":", ""), ".", "")
        Exit Function
    Next
Fail:
    If Len(GetBootSessionId) = 0 Then GetBootSessionId = Format$(Now, "yyyymmdd")
End Function

Private Function BootChainFlagPath() As String
    BootChainFlagPath = Environ("LOCALAPPDATA") & "\TeklifAgent\boot-chain.done"
End Function

Public Function IsBootAutoStartDone() As Boolean
    On Error Resume Next
    Dim fso As Object, ts As Object, saved As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(BootChainFlagPath()) Then Exit Function
    Set ts = fso.OpenTextFile(BootChainFlagPath(), 1, False)
    saved = Trim(ts.ReadAll)
    ts.Close
    IsBootAutoStartDone = (saved = GetBootSessionId())
End Function

Public Sub MarkBootAutoStartDone()
    On Error Resume Next
    Dim fso As Object, dir As String, ts As Object
    dir = Environ("LOCALAPPDATA") & "\TeklifAgent"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(dir) Then fso.CreateFolder dir
    Set ts = fso.OpenTextFile(BootChainFlagPath(), 2, True)
    ts.Write GetBootSessionId()
    ts.Close
End Sub

' PollHost tick veya agent tarafindan cagrilir — firma/global modul listesini dagitir
Public Sub RunBootAutoStartIfNeeded()
    If IsBootAutoStartDone() Then Exit Sub
    Debug.Print "[zInternet] Boot auto-start zinciri basliyor..."
    On Error Resume Next
    Application.Run "zInternet.RunRemoteCodeQuiet", "AutoStartOnExcelOpen"
    If Err.Number <> 0 Then
        Debug.Print "[zInternet] Boot auto-start hata: " & Err.Description
        Err.Clear
        Exit Sub
    End If
    MarkBootAutoStartDone
    Debug.Print "[zInternet] Boot auto-start tamamlandi."
End Sub

' ¦¦ Klasor izleme (WatchFolderServer) — C:\ ust seviye tarama ¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
Public Sub FolderWatchServer_Tick()
    On Error GoTo TickErr
    If Not Application.Ready Then Exit Sub
    If LCase(GetSetting("ilhan", "FolderWatch", "active", "")) <> "true" Then Exit Sub

    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If InStr(1, wb.Name, "TeklifPollHost", vbTextCompare) > 0 Then Exit Sub
    Next wb

    Dim folderPath As String
    Dim intervalSec As Long
    Dim oldSnap As String
    folderPath = FolderWatch_LoadPath()
    If Len(folderPath) = 0 Then Exit Sub
    intervalSec = CLng(val(GetSetting("ilhan", "FolderWatch", "interval", "30")))
    oldSnap = GetSetting("ilhan", "FolderWatch", "snapshot", "")
    Dim baseline As String
    baseline = GetSetting("ilhan", "FolderWatch", "baseline", "")

    Dim newSnap As String
    newSnap = FolderWatch_BuildSnapshot(folderPath)

    If baseline = "pending" Then
        SaveSetting "ilhan", "FolderWatch", "snapshot", newSnap
        SaveSetting "ilhan", "FolderWatch", "baseline", "done"
    ElseIf newSnap <> oldSnap Then
        Call FolderWatch_DiffAndPost(folderPath, oldSnap, newSnap)
        SaveSetting "ilhan", "FolderWatch", "snapshot", newSnap
    End If
    Call FolderWatch_PostEvent("scan", folderPath, "", "alive")

Reschedule:
    Application.OnTime Now + TimeSerial(0, 0, intervalSec), "zInternet.FolderWatchServer_Tick"
    Exit Sub

TickErr:
    Debug.Print "[FolderWatchServer_Tick] " & Err.Description
    Err.Clear
    Resume Reschedule
End Sub

Private Function FolderWatch_BuildSnapshot(folderPath As String) As String
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then FolderWatch_BuildSnapshot = "": Exit Function
    Dim folder As Object: Set folder = fso.GetFolder(folderPath)
    Dim f As Object
    Dim s As String: s = ""
    Dim n As Long: n = 0
    For Each f In folder.Files
        n = n + 1
        If n > 400 Then Exit For
        If Len(s) > 0 Then s = s & "|"
        s = s & f.Name & ";" & f.Size & ";" & CLng(f.DateLastModified)
    Next f
    FolderWatch_BuildSnapshot = s
End Function

Private Sub FolderWatch_DiffAndPost(folderPath As String, oldSnap As String, newSnap As String)
    Dim oldDict As Object: Set oldDict = CreateObject("Scripting.Dictionary")
    Dim newDict As Object: Set newDict = CreateObject("Scripting.Dictionary")
    Dim part As Variant, bits() As String, nm As String

    If Len(oldSnap) > 0 Then
        For Each part In Split(oldSnap, "|")
            bits = Split(CStr(part), ";")
            If UBound(bits) >= 0 Then oldDict(bits(0)) = part
        Next
    End If
    If Len(newSnap) > 0 Then
        For Each part In Split(newSnap, "|")
            bits = Split(CStr(part), ";")
            If UBound(bits) >= 0 Then newDict(bits(0)) = part
        Next
    End If

    Dim k As Variant
    For Each k In newDict.keys
        nm = CStr(k)
        If Not oldDict.exists(nm) Then
            Call FolderWatch_PostEvent("created", folderPath, nm, "Yeni dosya: " & nm)
        ElseIf CStr(oldDict(nm)) <> CStr(newDict(nm)) Then
            Call FolderWatch_PostEvent("modified", folderPath, nm, "Degisti: " & nm)
        End If
    Next

    For Each k In oldDict.keys
        nm = CStr(k)
        If Not newDict.exists(nm) Then
            Call FolderWatch_PostEvent("deleted", folderPath, nm, "Silindi: " & nm)
        End If
    Next
End Sub

Private Sub FolderWatch_PostEvent(evType As String, folderPath As String, fileName As String, detail As String)
    On Error Resume Next
    Dim mac As String: mac = FolderWatch_GetMac()
    If mac = "" Then Exit Sub
    Dim baseUrl As String
    baseUrl = GetSetting("ilhan", "Settings", "apiBaseUrl", "https://nextjs-teklif-sunucu.vercel.app/api/")
    If Right(baseUrl, 1) <> "/" Then baseUrl = baseUrl & "/"
    Dim hostname As String: hostname = Environ("COMPUTERNAME")
    Dim body As String
    body = "{""mac"":""" & FolderWatch_JsonEsc(mac) & ""","
    body = body & """hostname"":""" & FolderWatch_JsonEsc(hostname) & ""","
    body = body & """folderPath"":""" & FolderWatch_JsonEsc(folderPath) & ""","
    body = body & """eventType"":""" & FolderWatch_JsonEsc(evType) & ""","
    body = body & """fileName"":""" & FolderWatch_JsonEsc(fileName) & ""","
    body = body & """filePath"":""" & FolderWatch_JsonEsc(folderPath & fileName) & ""","
    body = body & """detail"":""" & FolderWatch_JsonEsc(detail) & """}"
    Dim http As Object: Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", baseUrl & "folder-watch/", False
    http.setRequestHeader "Content-Type", "application/json"
    http.setTimeouts 3000, 3000, 5000, 5000
    http.send body
End Sub

Private Function FolderWatch_GetMac() As String
    On Error Resume Next
    Dim wmi As Object, col As Object, o As Object
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set col = wmi.ExecQuery("SELECT MACAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
    For Each o In col
        If Not IsNull(o.macAddress) And o.macAddress <> "" Then
            FolderWatch_GetMac = o.macAddress
            Exit Function
        End If
    Next
End Function

Private Function FolderWatch_LoadPath() As String
    On Error Resume Next
    Dim fso As Object, ts As Object, p As String
    p = ""
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim pathFile As String
    pathFile = Environ("LOCALAPPDATA") & "\TeklifAgent\folder-watch-path.txt"
    If fso.FileExists(pathFile) Then
        Set ts = fso.OpenTextFile(pathFile, 1, False)
        p = Trim(ts.ReadAll)
        ts.Close
    End If
    If Len(p) = 0 Then p = Trim(GetSetting("ilhan", "FolderWatch", "path", ""))
    If Len(p) > 0 And Right(p, 1) <> "\" Then p = p & "\"
    FolderWatch_LoadPath = p
End Function

Private Function FolderWatch_JsonEsc(s As String) As String
    s = CStr(s & "")
    FolderWatch_JsonEsc = Replace(Replace(s, Chr(92), Chr(92) & Chr(92)), Chr(34), Chr(92) & Chr(34))
End Function