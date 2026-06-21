Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Debug.Print "[LisansKontrolVeGuncelleme] DynamicFunc basladi."
    Debug.Print "[LisansKontrolVeGuncelleme] targetWb: " & targetWb.Name & " | IsAddin: " & targetWb.IsAddin

    Dim tempPath As String
    Dim hostWb As Workbook

    tempPath = DownloadAddinToTemp(param)
    If Len(tempPath) = 0 Then
        Debug.Print "[LisansKontrolVeGuncelleme] Indirme basarisiz, cikiliyor."
        Set DynamicFunc = Nothing
        Exit Function
    End If

    Debug.Print "[LisansKontrolVeGuncelleme] Indirme tamam: " & tempPath

    Set hostWb = ResolveHostWorkbook(targetWb)
    If hostWb Is Nothing Then
        Debug.Print "[LisansKontrolVeGuncelleme] Ana xlsm bulunamadi!"
        MsgBox "Ana teklif dosyası (xlsm) bulunamadı." & vbCrLf & _
               "TeklifBootstrap modülünü ana dosyaya ekleyin.", vbCritical
        Set DynamicFunc = Nothing
        Exit Function
    End If

    Debug.Print "[LisansKontrolVeGuncelleme] Ana dosya: " & hostWb.Name
    Debug.Print "[LisansKontrolVeGuncelleme] TeklifBootstrap.ApplyPendingAddinUpdate cagriliyor..."

    Application.Run "'" & hostWb.Name & "'!TeklifBootstrap.ApplyPendingAddinUpdate", _
        tempPath, _
        Application.UserLibraryPath & "teklif.xlam", _
        "teklif", _
        "teklif.xlam"

    Debug.Print "[LisansKontrolVeGuncelleme] Kurulum cagrisi tamamlandi."
    Set DynamicFunc = Nothing
End Function

Private Function DownloadAddinToTemp(Optional apiBaseUrl As Variant) As String
    Const YEREL_DOSYA_ADI As String = "teklif.xlam"
    Const MODUL_ADI As String = "LisansKontrolVeGuncelleme"

    Dim mac As String
    Dim fullUrl As String
    Dim httpReq As Object
    Dim stream As Object
    Dim tempPath As String
    Dim jsonBody As String
    Dim baseUrl As String

    Debug.Print "[" & MODUL_ADI & "] DownloadAddinToTemp basladi."
    DownloadAddinToTemp = vbNullString
    mac = GetFirstMACAddress()
    Debug.Print "[" & MODUL_ADI & "] MAC: " & mac

    If Len(mac) < 10 Or Left(mac, 5) = "HATA:" Or mac = "MAC_BULUNAMADI" Then
        Debug.Print "[" & MODUL_ADI & "] Gecersiz MAC, indirme iptal."
        MsgBox "MAC adresi alınamadı.", vbCritical
        Exit Function
    End If

    baseUrl = ResolveApiBaseUrl(apiBaseUrl)
    fullUrl = baseUrl & "download/teklif/"
    tempPath = Environ("TEMP") & "\" & YEREL_DOSYA_ADI
    jsonBody = "{""macAdresi"": """ & mac & """}"

    Debug.Print "[" & MODUL_ADI & "] URL: " & fullUrl
    Debug.Print "[" & MODUL_ADI & "] Temp: " & tempPath

    Set httpReq = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    On Error GoTo BaglantiHatasi
    httpReq.Open "POST", fullUrl, False
    httpReq.setRequestHeader "Content-Type", "application/json"
    httpReq.send jsonBody

    Debug.Print "[" & MODUL_ADI & "] HTTP Status: " & httpReq.Status

    If httpReq.Status = 200 Then
        Set stream = CreateObject("ADODB.Stream")
        stream.Open
        stream.Type = 1
        stream.Write httpReq.responseBody
        stream.SaveToFile tempPath, 2
        stream.Close
        DownloadAddinToTemp = tempPath
        Debug.Print "[" & MODUL_ADI & "] Dosya kaydedildi: " & tempPath
    Else
        Debug.Print "[" & MODUL_ADI & "] Sunucu hatasi: " & httpReq.responseText
        MsgBox "İşlem Başarısız!" & vbCrLf & _
               "Hata Kodu: " & httpReq.Status & vbCrLf & _
               "Sunucu Mesajı: " & httpReq.responseText, vbExclamation
    End If

    Set httpReq = Nothing
    Set stream = Nothing
    Exit Function

BaglantiHatasi:
    Debug.Print "[" & MODUL_ADI & "] Baglanti hatasi: " & Err.Description
    MsgBox "Sunucuya bağlanırken hata oluştu." & vbCrLf & _
           "Hata: " & Err.Description, vbCritical
End Function

Private Function ResolveHostWorkbook(fallback As Workbook) As Workbook
    Dim wb As Workbook

    Debug.Print "[LisansKontrolVeGuncelleme] ResolveHostWorkbook basladi."

    If Not fallback Is Nothing Then
        Debug.Print "[LisansKontrolVeGuncelleme] fallback: " & fallback.Name & " | IsAddin: " & fallback.IsAddin
        If Not fallback.IsAddin Then
            Set ResolveHostWorkbook = fallback
            Debug.Print "[LisansKontrolVeGuncelleme] fallback ana dosya olarak secildi."
            Exit Function
        End If
    End If

    For Each wb In Application.Workbooks
        If Not wb.IsAddin Then
            Set ResolveHostWorkbook = wb
            Debug.Print "[LisansKontrolVeGuncelleme] Bulunan ana dosya: " & wb.Name
            Exit Function
        End If
    Next wb

    Debug.Print "[LisansKontrolVeGuncelleme] Ana dosya bulunamadi."
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

    If Not IsMissing(apiBaseUrl) Then
        url = Trim(CStr(apiBaseUrl))
    End If

    If Len(url) = 0 Then
        url = "http://localhost:3000/api/"
    End If

    If Right(url, 1) <> "/" Then url = url & "/"
    ResolveApiBaseUrl = url
    Debug.Print "[LisansKontrolVeGuncelleme] API Base URL: " & url
End Function
