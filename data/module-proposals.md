# RunRemoteCode — Modül Önerileri

Her satır: `[KATEGORİ] MethodName — Açıklama`  
Geliştirip eklemek için `/api/modules` (POST) kullanılır.

---

## 1. BİLGİSAYAR BİLGİLERİ

| # | MethodName | Açıklama |
|---|-----------|----------|
| 1 | GetComputerName | Bilgisayar adını MsgBox ile gösterir |
| 2 | GetWindowsVersion | Windows sürüm bilgisini döndürür (WMI) |
| 3 | GetCpuInfo | İşlemci adı, çekirdek sayısı (Win32_Processor) |
| 4 | GetRamInfo | Toplam ve boş RAM miktarı (GB) |
| 5 | GetDiskInfo | Tüm sürücülerin boyut / boş alan listesi |
| 6 | GetMacAddress | Aktif ağ kartının MAC adresini döndürür |
| 7 | GetIpAddress | Yerel IP adresini döndürür |
| 8 | GetPublicIp | İnternet üzerinden dış IP sorgular |
| 9 | GetLoggedInUser | Oturum açmış kullanıcı adı |
| 10 | GetDomainName | Bağlı olunan domain/workgroup adı |
| 11 | GetBiosVersion | BIOS sürüm ve tarih bilgisi |
| 12 | GetMotherboardInfo | Anakart üretici ve model bilgisi |
| 13 | GetGpuInfo | Ekran kartı adı ve sürücü sürümü |
| 14 | GetScreenResolution | Ekran çözünürlüğü (piksel) |
| 15 | GetInstalledSoftwareList | Yüklü yazılım listesini Excel sayfasına yazar |
| 16 | GetRunningProcesses | Çalışan process listesini sayfaya yazar |
| 17 | GetStartupPrograms | Başlangıç programları listesi (registry + WMI) |
| 18 | GetSystemUptime | Sistem çalışma süresi (gün/saat/dakika) |
| 19 | GetTimeZone | Sistem saat dilimi |
| 20 | GetWindowsLicenseKey | Windows ürün anahtarını registry'den okur |

---

## 2. KAYIT DEFTERİ (REGİSTRY) İŞLEMLERİ

| # | MethodName | Açıklama |
|---|-----------|----------|
| 21 | ReadRegistryValue | Parametre ile verilen anahtar değerini okur |
| 22 | WriteRegistryValue | Parametre ile verilen anahtara değer yazar |
| 23 | DeleteRegistryKey | Belirtilen registry anahtarını siler |
| 24 | ListRegistryKeys | Bir path altındaki tüm anahtarları listeler |
| 25 | ExportRegistrySection | Bir bölümü .reg dosyasına aktarır |
| 26 | ImportRegistryFile | .reg dosyasını sisteme uygular |
| 27 | CheckRegistryKeyExists | Anahtarın var olup olmadığını kontrol eder |
| 28 | BackupRegistrySection | VBA SaveSetting bölümünü JSON'a yedekler |
| 29 | RestoreRegistrySection | JSON yedekten VBA SaveSetting geri yükler |
| 30 | GetAllVbaSettings | Tüm ilhan/scngnr/sercan registry değerlerini sayfaya döker |

---

## 3. DOSYA / KLASÖR İŞLEMLERİ

| # | MethodName | Açıklama |
|---|-----------|----------|
| 31 | CreateFolder | Parametre ile verilen yolu oluşturur |
| 32 | DeleteFolder | Klasörü içeriğiyle birlikte siler |
| 33 | CopyFolder | Klasörü başka konuma kopyalar |
| 34 | MoveFolder | Klasörü taşır |
| 35 | ListFolderContents | Klasör içeriğini sayfaya yazar (ad, boyut, tarih) |
| 36 | CopyFile | Dosyayı hedef konuma kopyalar |
| 37 | MoveFile | Dosyayı taşır |
| 38 | DeleteFile | Dosyayı siler |
| 39 | RenameFile | Dosyayı yeniden adlandırır |
| 40 | FileExists | Dosyanın var olup olmadığını kontrol eder |
| 41 | GetFileSize | Dosya boyutunu KB/MB olarak döndürür |
| 42 | GetFileCreationDate | Dosya oluşturma tarihini döndürür |
| 43 | GetFileModifiedDate | Son değiştirilme tarihini döndürür |
| 44 | ReadTextFile | Metin dosyasını okur, hücreye yazar |
| 45 | WriteTextFile | Hücre içeriğini metin dosyasına yazar |
| 46 | AppendToTextFile | Metin dosyasının sonuna satır ekler |
| 47 | ReadCsvFile | CSV dosyasını aktif sayfaya aktarır |
| 48 | WriteCsvFile | Aktif sayfayı CSV olarak kaydeder |
| 49 | ZipFolder | Klasörü ZIP dosyasına sıkıştırır (Shell) |
| 50 | UnzipFile | ZIP dosyasını açar (Shell) |
| 51 | SearchFilesInFolder | Klasörde pattern ile dosya arar |
| 52 | GetFolderSize | Tüm alt klasörlerle birlikte toplam boyut |
| 53 | OpenFileWithDefaultApp | Dosyayı varsayılan uygulamada açar |
| 54 | GetDesktopPath | Masaüstü yolunu döndürür |
| 55 | GetDocumentsPath | Belgeler klasörü yolunu döndürür |
| 56 | CleanTempFolder | Windows TEMP klasörünü temizler |
| 57 | CountFilesInFolder | Klasördeki dosya sayısını sayar |
| 58 | GetNewestFileInFolder | Klasörde en son değiştirilen dosyayı bulur |
| 59 | BackupFileWithTimestamp | Dosyayı tarihli isimle yedek olarak kopyalar |
| 60 | WatchFolderChange | Bir klasörün değişimini bekler (WScript.Shell) |

---

## 4. İNTERNET / HTTP İŞLEMLERİ

| # | MethodName | Açıklama |
|---|-----------|----------|
| 61 | HttpGetJson | Parametre URL'den JSON indirir, hücreye yazar |
| 62 | HttpPostJson | JSON body ile POST isteği yapar |
| 63 | HttpDownloadFile | URL'den dosya indirir, ADODB.Stream ile kaydeder |
| 64 | HttpGetText | URL'den düz metin çeker |
| 65 | CheckUrlReachable | URL'nin erişilebilir olup olmadığını kontrol eder |
| 66 | GetExchangeRate | Döviz kuru API'sinden anlık kur çeker |
| 67 | GetWeather | Hava durumu API'sinden veri alır |
| 68 | SendWebhook | Slack/Teams webhook'una mesaj gönderir |
| 69 | UploadFileToServer | Dosyayı multipart/form-data ile sunucuya yükler |
| 70 | CheckInternetConnection | İnternet bağlantısını test eder |
| 71 | PingHost | Host'a ping atar, ms cinsinden döndürür |
| 72 | FetchHtmlPage | Web sayfasının HTML içeriğini indirir |
| 73 | ParseJsonResponse | JSON string'den anahtar değer çeker |
| 74 | GetLatestModuleVersion | Sunucudan modül versiyon numarası sorgular |
| 75 | SendErrorReport | Hata raporunu sunucuya POST eder |
| 76 | CheckForUpdate | Mevcut versiyon ile sunucu versiyonunu karşılaştırır |
| 77 | DownloadAndOpenExcel | Excel dosyasını indirir ve açar |
| 78 | SendEmailViaApi | REST API üzerinden e-posta gönderir |
| 79 | GetServerStatus | Sunucunun sağlık durumunu kontrol eder |
| 80 | SyncDataWithServer | Yerel veriyi sunucu ile senkronize eder |

---

## 5. POWERSHELL / KOMUT SATIRI İŞLEMLERİ

| # | MethodName | Açıklama |
|---|-----------|----------|
| 81 | RunPowerShellCommand | Parametre ile PS komutu çalıştırır, çıktısını döndürür |
| 82 | RunPowerShellScript | .ps1 dosyasını çalıştırır |
| 83 | RunCmdCommand | cmd.exe ile komut çalıştırır |
| 84 | GetPsOutput | PS komutunun stdout çıktısını hücreye yazar |
| 85 | SetExecutionPolicy | PS execution policy'yi günceller |
| 86 | InstallWindowsFeature | PS ile Windows özelliği yükler |
| 87 | GetWindowsUpdateList | Bekleyen Windows güncellemelerini listeler |
| 88 | InstallWindowsUpdates | PS üzerinden Windows güncellemelerini başlatır |
| 89 | GetEventLog | Windows olay günlüğünden son kayıtları çeker |
| 90 | FlushDnsCache | `ipconfig /flushdns` komutunu çalıştırır |
| 91 | ResetNetworkAdapter | Ağ bağdaştırıcısını sıfırlar |
| 92 | GetNetworkConfig | IP, DNS, Gateway bilgilerini sayfaya yazar |
| 93 | SetStaticIp | PS ile statik IP atar |
| 94 | EnableRemoteDesktop | RDP'yi PS ile aktif eder |
| 95 | GetFirewallRules | Windows Firewall kurallarını listeler |
| 96 | AddFirewallRule | Yeni firewall kuralı ekler |
| 97 | GetInstalledDrivers | Yüklü sürücü listesini sayfaya yazar |
| 98 | RestartService | Windows servisini yeniden başlatır |
| 99 | StopService | Windows servisini durdurur |
| 100 | StartService | Windows servisini başlatır |

---

## 6. EXCEL / WORKBOOK İŞLEMLERİ

| # | MethodName | Açıklama |
|---|-----------|----------|
| 101 | SaveAllWorkbooks | Tüm açık dosyaları kaydeder |
| 102 | CloseAllWorkbooks | Kaydetmeden tüm dosyaları kapatır |
| 103 | RefreshAllPivotTables | Tüm pivot tabloları yeniler |
| 104 | RefreshAllConnections | Veri bağlantılarını yeniler |
| 105 | ExportSheetAsPdf | Aktif sayfayı PDF olarak dışa aktarır |
| 106 | ExportAllSheetsAsPdf | Tüm sayfaları ayrı PDF'lere aktarır |
| 107 | ImportSheetFromFile | Başka bir dosyadan sayfa kopyalar |
| 108 | ProtectAllSheets | Tüm sayfaları parola ile korur |
| 109 | UnprotectAllSheets | Tüm sayfaların korumasını kaldırır |
| 110 | SendWorkbookByEmail | MAPI ile dosyayı e-posta ekine ekler |
| 111 | ClearAllNamedRanges | Tüm isimli aralıkları temizler |
| 112 | ListAllFormulas | Sayfadaki tüm formülleri listeler |
| 113 | ConvertFormulasToValues | Formülleri değerle değiştirir |
| 114 | RemoveDuplicateRows | Yinelenen satırları siler |
| 115 | SortSheetsByName | Sayfaları alfabetik sıralar |
| 116 | HideEmptyColumns | Boş sütunları gizler |
| 117 | AutoFitAllColumns | Tüm sütunları otomatik genişletir |
| 118 | SetPrintArea | Baskı alanını aktif seçime göre ayarlar |
| 119 | CompressImages | Sayfadaki resimleri sıkıştırır |
| 120 | BatchRenameSheets | Sayfaları parametre ile toplu yeniden adlandırır |

---

## 7. GÜVENLİK / KİMLİK DOĞRULAMA

| # | MethodName | Açıklama |
|---|-----------|----------|
| 121 | CheckLicenseStatus | Registry'den lisans durumunu okur, tab gösterir/gizler |
| 122 | ValidateMacWithServer | MAC adresini sunucuya doğrulatır |
| 123 | CheckFileIntegrity | Dosyanın MD5 hash'ini kontrol eder |
| 124 | DetectVirtualMachine | VM ortamında çalışıp çalışmadığını kontrol eder |
| 125 | LockWorkbookOnExpiry | Lisans süresi dolduysa dosyayı kilitler |
| 126 | GenerateHardwareId | CPU+MAC kombinasyonundan benzersiz ID üretir |
| 127 | EncryptText | Metni XOR ile şifreler |
| 128 | DecryptText | XOR şifrelenmiş metni çözer |
| 129 | CheckAdminRights | Uygulamanın yönetici haklarıyla çalışıp çalışmadığını kontrol eder |
| 130 | LogUserAction | Kullanıcı eylemini sunucuya loglar |

---

## 8. BİLDİRİM / KULLANICI ARABİRİMİ

| # | MethodName | Açıklama |
|---|-----------|----------|
| 131 | ShowToastNotification | Windows bildirim baloncuğu gösterir (PowerShell) |
| 132 | ShowProgressBar | Özel UserForm ile işlem çubuğu gösterir |
| 133 | ShowCustomInputForm | Özel parametreli giriş formu açar |
| 134 | ShowYesNoDialog | Evet/Hayır dialog kutusu açar |
| 135 | PlayBeepSound | Windows beep sesi çalar |
| 136 | OpenUrlInBrowser | URL'yi varsayılan tarayıcıda açar |
| 137 | ShowSystemTrayMessage | Görev çubuğu bildirim baloncuğu (WScript.Shell) |
| 138 | ShowPrintDialog | Baskı iletişim kutusunu açar |
| 139 | SetExcelTitle | Excel başlık çubuğunu günceller |
| 140 | ShowStatusBarMessage | Excel durum çubuğuna mesaj yazar |

---

## 9. ZAMANLANMIŞ / OTOMATİK İŞLEMLER

| # | MethodName | Açıklama |
|---|-----------|----------|
| 141 | ScheduleTaskOnce | Windows Task Scheduler'a tek seferlik görev ekler |
| 142 | ScheduleTaskDaily | Günlük tekrarlayan görev oluşturur |
| 143 | RemoveScheduledTask | Zamanlanmış görevi siler |
| 144 | RunOnNextLogin | RunOnce registry ile bir sonraki girişte çalışır |
| 145 | AutoSaveEvery5Min | Workbook'u her 5 dakikada otomatik kaydeder |
| 146 | MonitorFolderAndRun | Klasöre dosya düşünce modül tetikler |
| 147 | SendDailyReport | Her sabah otomatik rapor e-posta atar |
| 148 | CleanOldLogsDaily | 30 günden eski logları günlük siler |
| 149 | AutoUpdateModules | Sunucudan yeni modülleri otomatik çeker ve DB'ye yazar |
| 150 | HeartbeatPing | Her 10 dakikada sunucuya canlılık sinyali gönderir |

---

## 10. VERİ / DÖNÜŞÜM İŞLEMLERİ

| # | MethodName | Açıklama |
|---|-----------|----------|
| 151 | ConvertExcelToJson | Aktif sayfayı JSON string'e çevirir |
| 152 | ImportJsonToSheet | JSON array'i sayfaya aktarır |
| 153 | ConvertXmlToSheet | XML dosyasını sayfaya aktarır |
| 154 | ExportSheetAsJson | Sayfayı sunucuya JSON olarak yükler |
| 155 | ConvertTextToColumns | Metin veriyi sütunlara böler |
| 156 | MergeMultipleFiles | Birden fazla Excel dosyasını tek sayfada birleştirir |
| 157 | SplitSheetByColumn | Sayfayı sütun değerine göre ayrı dosyalara böler |
| 158 | NormalizePhoneNumbers | Türkiye formatına göre telefon numaralarını normalleştirir |
| 159 | ValidateTCKimlik | TC kimlik numarası algoritma doğrulaması |
| 160 | FormatCurrencyColumn | Para birimi sütununu standart formata çevirir |

---

## Kullanım Notları

```vba
' Parametresiz:
Application.Run "zInternet.RunRemoteCode", "MethodName"

' Tek parametre:
Application.Run "zInternet.RunRemoteCode", "MethodName", "deger"

' Çoklu parametre (JSON):
Dim p As String
p = "{""key1"":""val1"",""key2"":""val2""}"
Application.Run "zInternet.RunRemoteCode", "MethodName", p
```

### DynamicFunc şablonu

```vba
Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ' param: tek deger veya JSON string
    Dim deger As String
    deger = CStr(param)
    
    ' ... işlemler ...
    
    Set DynamicFunc = Nothing
End Function
```
