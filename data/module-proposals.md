# RunRemoteCode — Gelişmiş Modül Önerileri

Her modül `DynamicFunc(targetWb, param)` imzasını kullanır.  
Eklemek için: `POST /api/modules` → `{ methodName, description, category, code, active }`

---

## 1. BİLGİSAYAR & DONANIM BİLGİLERİ

| # | MethodName | Açıklama |
|---|-----------|----------|
| 1 | GetComputerName | Bilgisayar adını döndürür |
| 2 | GetWindowsVersion | Windows sürüm + build numarası (WMI) |
| 3 | GetCpuInfo | CPU adı, çekirdek sayısı, maks frekans (GHz) |
| 4 | GetRamInfo | Toplam / Kullanılan / Boş RAM (GB) |
| 5 | GetDiskInfo | Tüm sürücüler: boyut, boş alan, dosya sistemi |
| 6 | GetMacAddress | Aktif NIC'in MAC adresini döndürür |
| 7 | GetIpAddress | IPv4 + IPv6 yerel adresler |
| 8 | GetPublicIp | api.ipify.org üzerinden dış IP |
| 9 | GetLoggedInUser | Aktif Windows kullanıcısı |
| 10 | GetDomainName | Domain / Workgroup adı |
| 11 | GetBiosInfo | BIOS sürümü, tarih, üretici |
| 12 | GetMotherboardInfo | Anakart modeli ve seri numarası |
| 13 | GetGpuInfo | Ekran kartı adı, VRAM, sürücü versiyonu |
| 14 | GetScreenResolution | Ekran genişliği × yüksekliği (piksel) |
| 15 | GetAllNetworkAdapters | Tüm NIC'leri (IP, MAC, durum) sayfaya yazar |
| 16 | GetSystemUptime | Son açılıştan itibaren gün/saat/dakika |
| 17 | GetTimeZone | Saat dilimi adı ve UTC offseti |
| 18 | GetInstalledSoftwareList | Tüm yüklü yazılım listesini sayfaya döker |
| 19 | GetRunningProcesses | CPU/Bellek kullanımı ile process listesi |
| 20 | GetWindowsActivationStatus | Windows aktivasyon durumu (slmgr /xpr çıktısı) |
| 21 | GetBatteryInfo | Pil şarj yüzdesi ve kalan süre (laptop) |
| 22 | GetPrinterList | Yüklü yazıcılar ve varsayılan yazıcı |
| 23 | GetUsbDevices | Bağlı USB aygıtları listeler |
| 24 | GetAudioDevices | Ses aygıtları (varsayılan giriş/çıkış) |
| 25 | GetSystemLocale | Sistem dili, klavye düzeni, para birimi biçimi |

---

## 2. KAYIT DEFTERİ (REGISTRY) İŞLEMLERİ

| # | MethodName | Açıklama |
|---|-----------|----------|
| 26 | ReadRegistryValue | param=`HKCU\...\Key` ile değer okur, hücreye yazar |
| 27 | WriteRegistryValue | param=`{"path":"...","name":"...","value":"..."}` ile yazar |
| 28 | DeleteRegistryKey | Belirtilen anahtarı ve alt anahtarları siler |
| 29 | ListRegistryKeys | Bir path altındaki tüm anahtar/değerleri listeler |
| 30 | ExportRegistrySection | Seçili bölümü `.reg` dosyasına aktarır |
| 31 | ImportRegistryFile | `.reg` dosyasını sessizce içe aktarır |
| 32 | CheckRegistryKeyExists | Anahtarın varlığını true/false döndürür |
| 33 | BackupVbaSettings | VBA GetSetting değerlerini JSON dosyasına yedekler |
| 34 | RestoreVbaSettings | JSON yedeği VBA SaveSetting ile geri yükler |
| 35 | GetAllVbaSettings | ilhan/scngnr/sercan bölümlerini sayfaya döker |
| 36 | SetRunOnceCommand | HKCU RunOnce'a komut ekler |
| 37 | RemoveRunOnceCommand | RunOnce kaydını siler |
| 38 | GetStartupPrograms | HKCU/HKLM Run anahtarlarındaki başlangıç programları |
| 39 | AddStartupProgram | Başlangıca program ekler (HKCU Run) |
| 40 | RemoveStartupProgram | Başlangıç programı kaydını kaldırır |

---

## 3. DOSYA / KLASÖR İŞLEMLERİ

| # | MethodName | Açıklama |
|---|-----------|----------|
| 41 | CreateFolder | Parametre ile klasör oluşturur (recursive) |
| 42 | DeleteFolder | Klasörü içeriğiyle birlikte siler |
| 43 | CopyFolder | Klasörü hedefe kopyalar |
| 44 | MoveFolder | Klasörü taşır |
| 45 | ListFolderContents | Ad, boyut, tarih, uzantı bilgisiyle listeler |
| 46 | CopyFile | Dosya kopyalar, üzerine yazma seçeneği |
| 47 | MoveFile | Dosya taşır |
| 48 | DeleteFile | Dosya siler (geri dönüşüm kutusu bypass) |
| 49 | RenameFile | Dosyayı yeniden adlandırır |
| 50 | FileExists | Dosyanın varlığını kontrol eder |
| 51 | GetFileSize | Dosya boyutu (B/KB/MB/GB otomatik birim) |
| 52 | GetFileHashMd5 | Dosyanın MD5 hash'ini hesaplar (bütünlük kontrolü) |
| 53 | ReadTextFile | UTF-8 metin dosyasını okur, hücreye yazar |
| 54 | WriteTextFile | Hücre içeriğini metin dosyasına yazar |
| 55 | AppendToTextFile | Metin dosyasına satır ekler |
| 56 | ReadCsvToSheet | CSV dosyasını aktif sayfaya aktarır |
| 57 | WriteSheetToCsv | Aktif sayfayı CSV olarak dışa aktarır |
| 58 | ZipFolder | Shell üzerinden klasörü ZIP'ler |
| 59 | UnzipToFolder | ZIP dosyasını hedef klasöre açar |
| 60 | SearchFilesByPattern | `*.xlsx` gibi pattern ile özyinelemeli arama |
| 61 | GetFolderSize | Alt klasörler dahil toplam disk kullanımı |
| 62 | OpenFileWithDefaultApp | Varsayılan uygulamada açar (Shell) |
| 63 | CleanTempFolder | `%TEMP%` klasörünü temizler |
| 64 | BackupFileWithTimestamp | Dosyayı `ad_YYYYMMDD_HHMMSS.bak` olarak kopyalar |
| 65 | WatchFolderForNewFile | Klasöre yeni dosya düşene kadar bekler |
| 66 | CompareFilesIdentical | İki dosyanın byte-by-byte aynı olup olmadığını kontrol eder |
| 67 | ReplaceTextInFile | Metin dosyasında bul-değiştir işlemi yapar |
| 68 | GetNewestFileInFolder | Son değiştirilen dosyayı bulur ve tam yolunu döndürür |
| 69 | GetFileAttributes | Gizli/Salt-okunur/Sistem özniteliklerini okur |
| 70 | SetFileAttribute | Dosya özniteliğini (Gizli vb.) değiştirir |

---

## 4. İNTERNET / HTTP İŞLEMLERİ

| # | MethodName | Açıklama |
|---|-----------|----------|
| 71 | HttpGetJson | URL'den JSON indirir, parse eder, sayfaya yazar |
| 72 | HttpPostJson | JSON body ile POST; yanıtı döndürür |
| 73 | HttpDownloadFile | Dosyayı ADODB.Stream ile kaydeder (binary) |
| 74 | HttpGetText | Düz metin yanıt alır |
| 75 | HttpPatchJson | PATCH isteği gönderir |
| 76 | HttpDeleteRequest | DELETE isteği gönderir |
| 77 | CheckUrlReachable | URL'ye HEAD isteği atarak erişilebilirliği test eder |
| 78 | GetExchangeRate | TCMB/fixer.io'dan anlık döviz kuru çeker |
| 79 | GetGoldPrice | Altın fiyatı API'sinden güncel fiyat alır |
| 80 | SendSlackMessage | Slack Incoming Webhook'a mesaj gönderir |
| 81 | SendTeamsMessage | Teams Adaptive Card webhook gönderir |
| 82 | SendTelegramMessage | Telegram Bot API ile mesaj gönderir |
| 83 | UploadFileToBlobStorage | Azure Blob / S3 uyumlu REST API'ye dosya yükler |
| 84 | CheckInternetConnection | `connectivity-check` üzerinden bağlantı testi |
| 85 | PingHost | WMI Win32_PingStatus ile ping, ms cinsinden |
| 86 | GetLatestModuleVersion | Sunucudan modül versiyon numarası sorgular |
| 87 | CheckForUpdate | Mevcut versiyon < sunucu versiyonu ise güncelleme önerir |
| 88 | DownloadAndOpenExcel | Excel dosyasını indirir, açar, değişken sayfa adına gider |
| 89 | SendErrorReportToServer | Hata stack trace'ini JSON ile sunucuya gönderir |
| 90 | FetchAndFillForm | API'den gelen JSON'u Excel UserForm alanlarına doldurur |
| 91 | WebScrapeSimple | MSXML2 ile sayfa HTML'ini indirir, belirli etiket içeriğini döndürür |
| 92 | SubmitFormData | HTML form gibi application/x-www-form-urlencoded gönderir |
| 93 | GetRedirectedUrl | Yönlendirme zincirinin son URL'sini bulur |
| 94 | BasicAuthGet | Basic Auth başlıklı GET isteği yapar |
| 95 | BearerTokenGet | Bearer token ile korumalı endpoint'ten veri alır |

---

## 5. POWERSHELL / KOMUT SATIRI

| # | MethodName | Açıklama |
|---|-----------|----------|
| 96 | RunPsCommand | PS komutu çalıştırır, stdout'u döndürür |
| 97 | RunPsScript | .ps1 dosyasını çalıştırır, çıktıyı hücreye yazar |
| 98 | RunCmdCommand | cmd.exe /c komutunu çalıştırır |
| 99 | GetPsOutputToSheet | PS çıktısını satır satır sayfaya yazar |
| 100 | SetPsExecutionPolicy | ExecutionPolicy ayarlar (RemoteSigned vb.) |
| 101 | GetWindowsUpdateList | Bekleyen Windows güncellemelerini listeler |
| 102 | InstallWindowsUpdates | PS ile güncelleme başlatır (PSWindowsUpdate modülü) |
| 103 | GetEventLogErrors | Son N Application/System hatasını çeker |
| 104 | FlushDnsCache | `ipconfig /flushdns` çalıştırır |
| 105 | ResetNetworkAdapter | Bağdaştırıcıyı devre dışı bırakıp tekrar etkinleştirir |
| 106 | GetNetworkConfig | IP, DNS, Gateway, DHCP bilgilerini sayfaya yazar |
| 107 | SetStaticIp | PS ile statik IP/DNS atar |
| 108 | EnableRemoteDesktop | RDP kaydını ve servisi aktif eder |
| 109 | GetFirewallRules | Aktif güvenlik duvarı kurallarını listeler |
| 110 | AddFirewallRule | İnbound/Outbound kural ekler |
| 111 | GetInstalledDrivers | Yüklü sürücüleri (InfName, Sürüm, Tarih) listeler |
| 112 | RestartWindowsService | Servis adına göre yeniden başlatır |
| 113 | GetServiceStatus | Servis durumunu (Running/Stopped) döndürür |
| 114 | KillProcessByName | İsme göre process sonlandırır |
| 115 | GetDiskHealthStatus | SMART verisini PS üzerinden alır |
| 116 | RunAsAdmin | Komutu yükseltilmiş (admin) PS ile çalıştırır |
| 117 | GetEnvironmentVariables | Tüm env değişkenlerini sayfaya yazar |
| 118 | SetEnvironmentVariable | Kullanıcı env değişkeni atar |
| 119 | CreateScheduledTask | Zamanlanmış görev oluşturur (XML ile) |
| 120 | RemoveScheduledTask | Zamanlanmış görevi siler |

---

## 6. EXCEL / WORKBOOK / RAPORLAMA

| # | MethodName | Açıklama |
|---|-----------|----------|
| 121 | SaveAllWorkbooks | Tüm açık dosyaları kaydeder |
| 122 | ExportSheetAsPdf | Aktif sayfayı PDF olarak dışa aktarır |
| 123 | ExportAllSheetsAsPdf | Her sayfayı ayrı PDF'e aktarır, isimleri sayfa adı |
| 124 | ImportSheetFromFile | Başka dosyadan sayfa kopyalar |
| 125 | ProtectAllSheets | Tüm sayfaları parola ile korur |
| 126 | UnprotectAllSheets | Tüm sayfa korumalarını kaldırır |
| 127 | RefreshAllPivotTables | Tüm pivot tabloları yeniler |
| 128 | RefreshAllConnections | Dış veri bağlantılarını yeniler |
| 129 | ConvertFormulasToValues | Seçilen aralıktaki formülleri değerle değiştirir |
| 130 | RemoveDuplicateRows | Yinelenen satırları siler, kaç adet silindiğini raporlar |
| 131 | SortSheetsByName | Sayfaları alfabetik sıralar |
| 132 | BatchRenameSheets | JSON parametresindeki ad eşleştirmesiyle toplu yeniden adlandırır |
| 133 | MergeMultipleFiles | Birden fazla Excel dosyasını tek sayfada birleştirir |
| 134 | SplitSheetByColumn | Sayfayı sütun değerine göre ayrı dosyalara böler |
| 135 | AutoFitAllColumns | Tüm sütun genişliklerini içeriğe göre ayarlar |
| 136 | AddWatermarkToSheet | Sayfa arka planına metin filigran ekler |
| 137 | SendWorkbookByEmail | MAPI / Outlook COM ile dosyayı e-posta ekine ekler |
| 138 | CreateSummarySheet | Tüm sayfaların A1 değerlerini özet sayfada toplar |
| 139 | CompressAllImages | Sayfadaki resimlerin sıkıştırma kalitesini düşürür |
| 140 | TableToJsonAndPost | Seçili tabloyu JSON'a çevirip API'ye gönderir |

---

## 7. GÜVENLİK & LİSANS

| # | MethodName | Açıklama |
|---|-----------|----------|
| 141 | CheckLicenseStatus | Registry'den lisans okur; tab'ı gösterir/gizler |
| 142 | ValidateMacWithServer | MAC + HWID'yi sunucuya doğrulatır |
| 143 | CheckFileIntegrity | Dosya MD5 hash'ini beklenen değerle karşılaştırır |
| 144 | DetectVirtualMachine | VM ortamında çalışılıp çalışılmadığını tespit eder |
| 145 | GenerateHardwareId | CPU SerialNumber + MAC → benzersiz 32 hex ID |
| 146 | EncryptTextXor | XOR + Base64 ile metin şifreleme |
| 147 | DecryptTextXor | XOR + Base64 ile metin çözme |
| 148 | CheckAdminRights | Yönetici haklarıyla çalışılıp çalışılmadığını kontrol eder |
| 149 | LockWorkbookOnExpiry | Lisans süresi dolduysa dosyayı kilitler + sunucuya bildirir |
| 150 | DetectCopyAndSelfDestruct | startingAddin kontrolü; ihlal varsa cleanup tetikler |
| 151 | AuditLogAction | Kullanıcı eylemini timestamp + mac + detail ile sunucuya loglar |
| 152 | ObfuscateSheetFormulas | Formülleri xlVeryHidden sayfalara taşıyarak gizler |
| 153 | CheckDebuggerAttached | VBA debugger'ın çalışıp çalışmadığını tespit eder |
| 154 | EncryptCellRange | Seçili hücreleri RC4 algoritmasıyla şifreler |
| 155 | DecryptCellRange | RC4 ile şifrelenmiş hücreleri çözer |

---

## 8. BİLDİRİM & KULLANICI ARABİRİMİ

| # | MethodName | Açıklama |
|---|-----------|----------|
| 156 | ShowToastNotification | Windows 10/11 baloncuk bildirimi (PS BurntToast) |
| 157 | ShowProgressBar | Özel UserForm ile %0→%100 ilerleme çubuğu |
| 158 | ShowCustomInputForm | Çok alanlı giriş formu; sonuçları JSON döndürür |
| 159 | ShowYesNoCancelDialog | Üç seçenekli dialog, seçimi string döndürür |
| 160 | PlaySystemSound | Windows ses teması çalar (Asterisk, Critical vb.) |
| 161 | OpenUrlInBrowser | Varsayılan tarayıcıda URL açar |
| 162 | ShowSystemTrayBalloon | WScript.Shell PopUp baloncuğu |
| 163 | SetExcelTitleBar | Excel başlık çubuğunu özelleştirir |
| 164 | ShowStatusBarProgress | Durum çubuğunda % göstergesiyle uzun işlem |
| 165 | FlashTaskbarIcon | Görev çubuğu simgesini yanıp söndürür (dikkat çekme) |

---

## 9. VERİ DÖNÜŞÜM & ENTEGRASYON

| # | MethodName | Açıklama |
|---|-----------|----------|
| 166 | ConvertSheetToJson | Aktif sayfayı JSON array string'e çevirir |
| 167 | ImportJsonToSheet | JSON array'i başlık + satırlarıyla sayfaya aktarır |
| 168 | ConvertXmlToSheet | XML dökümanını sayfaya çevirir (MSXML2.DOMDocument) |
| 169 | ExportSheetToServer | Sayfayı JSON olarak sunucu API'sine POST eder |
| 170 | NormalizePhoneNumbers | TR format telefon numaralarını `+90 5xx xxx xx xx`'e çevirir |
| 171 | ValidateTCKimlik | TC kimlik algoritma doğrulaması |
| 172 | FormatCurrencyColumn | Para birimi sütununu `₺ 1.234,56` formatına çevirir |
| 173 | MergeJsonFiles | Birden fazla JSON dosyasını tek array'de birleştirir |
| 174 | SplitCsvByColumn | CSV dosyasını sütun değerine göre ayrı dosyalara böler |
| 175 | SqliteQueryToSheet | SQLite dosyasını ADO üzerinden sorgular, sonucu sayfaya yazar |

---

## 10. ZAMANLANMIŞ & OTOMATİK

| # | MethodName | Açıklama |
|---|-----------|----------|
| 176 | ScheduleTaskOnce | Belirli tarih/saatte tek seferlik görev oluşturur |
| 177 | ScheduleTaskDaily | Her gün belirli saatte tekrarlayan görev |
| 178 | ScheduleTaskOnLogin | Windows girişinde çalışan görev |
| 179 | RemoveScheduledTask | İsme göre zamanlanmış görevi siler |
| 180 | AutoSaveWorkbook | Her N dakikada otomatik kayıt + sürümlü yedek |
| 181 | MonitorFolderTrigger | Klasöre dosya eklenince belirtilen modülü tetikler |
| 182 | SendDailyEmailReport | Belirlenen saatte sayfayı PDF'e çevirerek e-posta atar |
| 183 | CleanOldBackups | N günden eski yedek dosyalarını temizler |
| 184 | AutoUpdateModules | Sunucudan modül listesini çekip DB ile karşılaştırır, günceller |
| 185 | HeartbeatPing | Her N dakikada MAC + versiyon ile sunucuya sinyal gönderir |

---

## 11. GELİŞMİŞ / UZMAN MODÜLLER

| # | MethodName | Açıklama |
|---|-----------|----------|
| 186 | SelfUpdateAddin | Sunucudan yeni teklif.xlam indirir, mevcut sürümü değiştirir |
| 187 | InjectVbaModule | Çalışma zamanında hedef workbook'a VBA modülü enjekte eder |
| 188 | RemoveVbaModule | Workbook'tan modülü programatik olarak siler |
| 189 | RunMacroInWorkbook | Parametre workbook adındaki makroyu çalıştırır |
| 190 | CallDllFunction | `Declare`/`LoadLibrary` ile native DLL fonksiyonu çağırır |
| 191 | ReadWriteNamedPipe | Windows adlandırılmış boru (Named Pipe) üzerinden IPC |
| 192 | SendKeystrokes | `SendKeys` veya UI Automation ile tuş dizisi gönderir |
| 193 | CaptureScreenshot | PrintScreen + clipboard + kaydedilmiş PNG |
| 194 | ReadQrCode | QR kod resim dosyasını ZXing COM ile okur |
| 195 | GenerateBarcode | Code128 barkod formülü ile hücreye barkod yazar |
| 196 | SignPdfWithCertificate | iTextSharp COM ile PDF'e dijital imza atar |
| 197 | ConnectToSqlServer | ADO üzerinden SQL Server sorgusu çalıştırır |
| 198 | ConnectToMySql | MySQL ODBC connector ile sorgu |
| 199 | ReadFromExcelOneDrive | SharePoint/OneDrive URL'den Excel dosyasını okur |
| 200 | WatchClipboard | Pano değişimini izler, metin kopyalanınca modül tetikler |

---

## DynamicFunc Şablonları

### Basit modül
```vba
Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim deger As String
    deger = CStr(param)
    MsgBox deger, vbInformation
    Set DynamicFunc = Nothing
End Function
```

### HTTP GET + JSON ayrıştırma
```vba
Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim http As Object
    Dim url As String
    url = CStr(param)                     ' param = URL

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.setTimeouts 5000, 10000, 30000, 30000
    http.send

    If http.Status = 200 Then
        Dim ws As Worksheet
        Set ws = targetWb.Sheets(1)
        ws.Range("A1").Value = http.responseText
    End If
    Set DynamicFunc = Nothing
End Function
```

### Çoklu parametre (JSON)
```vba
' Çağrı: RunRemoteCode "ModulAdi", "{""key1"":""val1"",""key2"":42}"
Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String
    p = CStr(param)

    ' Basit JSON değer çekme
    Dim val1 As String
    val1 = ExtractJsonValue(p, "key1")    ' "val1"

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
        p1 = p1 + 1
        p2 = InStr(p1, json, """")
        ExtractJsonValue = Mid(json, p1, p2 - p1)
    Else
        p2 = p1
        Do While InStr(",}]" & Chr(13) & Chr(10), Mid(json, p2, 1)) = 0 : p2 = p2 + 1 : Loop
        ExtractJsonValue = Trim(Mid(json, p1, p2 - p1))
    End If
End Function
```

### PowerShell çıktısı alma
```vba
Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim cmd As String
    Dim wsh As Object
    Dim exec As Object
    Dim output As String

    cmd = "powershell -NonInteractive -Command """ & CStr(param) & """"
    Set wsh = CreateObject("WScript.Shell")
    Set exec = wsh.Exec(cmd)

    Do While exec.Status = 0 : Application.Wait Now + TimeValue("00:00:01") : Loop
    output = exec.StdOut.ReadAll

    targetWb.Sheets(1).Range("A1").Value = Trim(output)
    Set DynamicFunc = Nothing
End Function
```
