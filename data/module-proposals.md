# RunRemoteCode — Gelişmiş Modül Önerileri

Her modül `DynamicFunc(targetWb, param)` imzasını kullanır.  
Eklemek için: `POST /api/modules` → `{ methodName, description, category, code, active }`

> **Durum**: ✅ = Uygulandı · ⬜ = Planlandı

---

## 1. BİLGİSAYAR & DONANIM BİLGİLERİ

| # | MethodName | Açıklama | Durum |
|---|-----------|----------|-------|
| 1 | GetComputerName | Bilgisayar adını döndürür | ✅ |
| 2 | GetWindowsVersion | Windows sürüm + build numarası (WMI) | ✅ |
| 3 | GetCpuInfo | CPU adı, çekirdek sayısı, maks frekans (GHz) | ✅ |
| 4 | GetRamInfo | Toplam / Kullanılan / Boş RAM (GB) | ✅ |
| 5 | GetDiskInfo | Tüm sürücüler: boyut, boş alan, dosya sistemi | ✅ |
| 6 | GetMacAddress | Aktif NIC'in MAC adresini döndürür | ✅ |
| 7 | GetIpAddress | IPv4 + IPv6 yerel adresler | ✅ |
| 8 | GetPublicIp | api.ipify.org üzerinden dış IP | ✅ |
| 9 | GetLoggedInUser | Aktif Windows kullanıcısı | ✅ |
| 10 | GetDomainName | Domain / Workgroup adı | ✅ |
| 11 | GetBiosInfo | BIOS sürümü, tarih, üretici | ✅ |
| 12 | GetMotherboardInfo | Anakart modeli ve seri numarası | ✅ |
| 13 | GetGpuInfo | Ekran kartı adı, VRAM, sürücü versiyonu | ✅ |
| 14 | GetScreenResolution | Ekran genişliği × yüksekliği (piksel) | ✅ |
| 15 | GetAllNetworkAdapters | Tüm NIC'leri (IP, MAC, durum) sayfaya yazar | ✅ |
| 16 | GetSystemUptime | Son açılıştan itibaren gün/saat/dakika | ✅ |
| 17 | GetTimeZone | Saat dilimi adı ve UTC offseti | ✅ |
| 18 | GetInstalledSoftwareList | Tüm yüklü yazılım listesini sayfaya döker | ✅ |
| 19 | GetRunningProcesses | CPU/Bellek kullanımı ile process listesi | ✅ |
| 20 | GetWindowsActivationStatus | Windows aktivasyon durumu (slmgr /xpr çıktısı) | ✅ |
| 21 | GetBatteryInfo | Pil şarj yüzdesi ve kalan süre (laptop) | ✅ |
| 22 | GetPrinterList | Yüklü yazıcılar ve varsayılan yazıcı | ✅ |
| 23 | GetUsbDevices | Bağlı USB aygıtları listeler | ✅ |
| 24 | GetAudioDevices | Ses aygıtları (varsayılan giriş/çıkış) | ✅ |
| 25 | GetSystemLocale | Sistem dili, klavye düzeni, para birimi biçimi | ✅ |
| 26 | GetCpuUsage | Anlık CPU kullanım yüzdesi (WMI LoadPercentage) | ⬜ |
| 27 | GetVirtualMemoryInfo | Sayfa dosyası boyutu ve kullanımı | ⬜ |
| 28 | GetNetworkSpeed | Adaptör bant genişliği ve anlık kullanım (Mbps) | ⬜ |
| 29 | GetHardwareSerial | Bilgisayar seri numarası (Win32_ComputerSystemProduct) | ⬜ |
| 30 | GetWindowsProductKey | Kayıt defterinden Windows ürün anahtarı okuma | ⬜ |

---

## 2. KAYIT DEFTERİ (REGISTRY) İŞLEMLERİ

| # | MethodName | Açıklama | Durum |
|---|-----------|----------|-------|
| 31 | ReadRegistryValue | param=`HKCU\...\Key` ile değer okur, hücreye yazar | ⬜ |
| 32 | WriteRegistryValue | param=`{"path":"...","name":"...","value":"..."}` ile yazar | ⬜ |
| 33 | DeleteRegistryKey | Belirtilen anahtarı ve alt anahtarları siler | ⬜ |
| 34 | ListRegistryKeys | Bir path altındaki tüm anahtar/değerleri listeler | ⬜ |
| 35 | ExportRegistrySection | Seçili bölümü `.reg` dosyasına aktarır | ⬜ |
| 36 | ImportRegistryFile | `.reg` dosyasını sessizce içe aktarır | ⬜ |
| 37 | CheckRegistryKeyExists | Anahtarın varlığını true/false döndürür | ⬜ |
| 38 | BackupVbaSettings | VBA GetSetting değerlerini JSON dosyasına yedekler | ⬜ |
| 39 | RestoreVbaSettings | JSON yedeği VBA SaveSetting ile geri yükler | ⬜ |
| 40 | GetAllVbaSettings | ilhan/scngnr/sercan bölümlerini sayfaya döker | ⬜ |
| 41 | SetRunOnceCommand | HKCU RunOnce'a komut ekler | ⬜ |
| 42 | RemoveRunOnceCommand | RunOnce kaydını siler | ⬜ |
| 43 | GetStartupPrograms | HKCU/HKLM Run anahtarlarındaki başlangıç programları | ⬜ |
| 44 | AddStartupProgram | Başlangıca program ekler (HKCU Run) | ⬜ |
| 45 | RemoveStartupProgram | Başlangıç programı kaydını kaldırır | ⬜ |
| 46 | MonitorRegistryChange | Belirli bir registry anahtarını değişiklik için izler | ⬜ |
| 47 | GetInstalledAppPaths | HKLM App Paths'tan uygulama exe yollarını listeler | ⬜ |
| 48 | CompareRegistrySnapshot | İki farklı zamandaki registry farkını raporlar | ⬜ |

---

## 3. DOSYA / KLASÖR İŞLEMLERİ

| # | MethodName | Açıklama | Durum |
|---|-----------|----------|-------|
| 49 | CreateFolder | Parametre ile klasör oluşturur (recursive) | ⬜ |
| 50 | DeleteFolder | Klasörü içeriğiyle birlikte siler | ⬜ |
| 51 | CopyFolder | Klasörü hedefe kopyalar | ⬜ |
| 52 | MoveFolder | Klasörü taşır | ⬜ |
| 53 | ListFolderContents | Ad, boyut, tarih, uzantı bilgisiyle listeler | ⬜ |
| 54 | CopyFile | Dosya kopyalar, üzerine yazma seçeneği | ⬜ |
| 55 | MoveFile | Dosya taşır | ⬜ |
| 56 | DeleteFile | Dosya siler (geri dönüşüm kutusu bypass) | ⬜ |
| 57 | RenameFile | Dosyayı yeniden adlandırır | ⬜ |
| 58 | FileExists | Dosyanın varlığını kontrol eder | ⬜ |
| 59 | GetFileSize | Dosya boyutu (B/KB/MB/GB otomatik birim) | ⬜ |
| 60 | GetFileHashMd5 | Dosyanın MD5 hash'ini hesaplar (bütünlük kontrolü) | ⬜ |
| 61 | ReadTextFile | UTF-8 metin dosyasını okur, hücreye yazar | ⬜ |
| 62 | WriteTextFile | Hücre içeriğini metin dosyasına yazar | ⬜ |
| 63 | AppendToTextFile | Metin dosyasına satır ekler | ⬜ |
| 64 | ReadCsvToSheet | CSV dosyasını aktif sayfaya aktarır | ⬜ |
| 65 | WriteSheetToCsv | Aktif sayfayı CSV olarak dışa aktarır | ⬜ |
| 66 | ZipFolder | Shell üzerinden klasörü ZIP'ler | ⬜ |
| 67 | UnzipToFolder | ZIP dosyasını hedef klasöre açar | ⬜ |
| 68 | SearchFilesByPattern | `*.xlsx` gibi pattern ile özyinelemeli arama | ⬜ |
| 69 | GetFolderSize | Alt klasörler dahil toplam disk kullanımı | ⬜ |
| 70 | OpenFileWithDefaultApp | Varsayılan uygulamada açar (Shell) | ⬜ |
| 71 | CleanTempFolder | `%TEMP%` klasörünü temizler | ⬜ |
| 72 | BackupFileWithTimestamp | Dosyayı `ad_YYYYMMDD_HHMMSS.bak` olarak kopyalar | ⬜ |
| 73 | WatchFolderForNewFile | Klasöre yeni dosya düşene kadar bekler | ⬜ |
| 74 | CompareFilesIdentical | İki dosyanın byte-by-byte aynı olup olmadığını kontrol eder | ⬜ |
| 75 | ReplaceTextInFile | Metin dosyasında bul-değiştir işlemi yapar | ⬜ |
| 76 | GetNewestFileInFolder | Son değiştirilen dosyayı bulur ve tam yolunu döndürür | ⬜ |
| 77 | GetFileAttributes | Gizli/Salt-okunur/Sistem özniteliklerini okur | ⬜ |
| 78 | SetFileAttribute | Dosya özniteliğini (Gizli vb.) değiştirir | ⬜ |
| 79 | SyncFolderToServer | Yerel klasörü REST API ile uzak sunucuya senkronize eder | ⬜ |
| 80 | ConvertPdfToText | COM ile PDF metnini çıkartır, sayfaya yazar | ⬜ |

---

## 4. İNTERNET / HTTP İŞLEMLERİ

| # | MethodName | Açıklama | Durum |
|---|-----------|----------|-------|
| 81 | HttpGetJson | URL'den JSON indirir, parse eder, sayfaya yazar | ⬜ |
| 82 | HttpPostJson | JSON body ile POST; yanıtı döndürür | ⬜ |
| 83 | HttpDownloadFile | Dosyayı ADODB.Stream ile kaydeder (binary) | ⬜ |
| 84 | HttpGetText | Düz metin yanıt alır | ⬜ |
| 85 | HttpPatchJson | PATCH isteği gönderir | ⬜ |
| 86 | HttpDeleteRequest | DELETE isteği gönderir | ⬜ |
| 87 | CheckUrlReachable | URL'ye HEAD isteği atarak erişilebilirliği test eder | ⬜ |
| 88 | GetExchangeRate | TCMB/fixer.io'dan anlık döviz kuru çeker | ⬜ |
| 89 | GetGoldPrice | Altın fiyatı API'sinden güncel fiyat alır | ⬜ |
| 90 | SendSlackMessage | Slack Incoming Webhook'a mesaj gönderir | ⬜ |
| 91 | SendTeamsMessage | Teams Adaptive Card webhook gönderir | ⬜ |
| 92 | SendTelegramMessage | Telegram Bot API ile mesaj gönderir | ⬜ |
| 93 | UploadFileToBlobStorage | Azure Blob / S3 uyumlu REST API'ye dosya yükler | ⬜ |
| 94 | CheckInternetConnection | connectivity-check üzerinden bağlantı testi | ⬜ |
| 95 | PingHost | WMI Win32_PingStatus ile ping, ms cinsinden | ⬜ |
| 96 | GetLatestModuleVersion | Sunucudan modül versiyon numarası sorgular | ⬜ |
| 97 | CheckForUpdate | Mevcut versiyon < sunucu versiyonu ise güncelleme önerir | ⬜ |
| 98 | DownloadAndOpenExcel | Excel dosyasını indirir, açar, değişken sayfa adına gider | ⬜ |
| 99 | SendErrorReportToServer | Hata stack trace'ini JSON ile sunucuya gönderir | ⬜ |
| 100 | FetchAndFillForm | API'den gelen JSON'u Excel UserForm alanlarına doldurur | ⬜ |
| 101 | WebScrapeSimple | MSXML2 ile sayfa HTML'ini indirir, belirli etiket içeriğini döndürür | ⬜ |
| 102 | SubmitFormData | HTML form gibi application/x-www-form-urlencoded gönderir | ⬜ |
| 103 | GetRedirectedUrl | Yönlendirme zincirinin son URL'sini bulur | ⬜ |
| 104 | BasicAuthGet | Basic Auth başlıklı GET isteği yapar | ⬜ |
| 105 | BearerTokenGet | Bearer token ile korumalı endpoint'ten veri alır | ⬜ |
| 106 | UploadExcelToSharePoint | SharePoint REST API ile Excel dosyasını yükler | ⬜ |
| 107 | GetWeatherData | OpenWeatherMap API ile hava durumu bilgisi çeker | ⬜ |
| 108 | FetchCurrencyHistory | Son 30 günlük döviz kuru tarihçesini sayfaya yazar | ⬜ |
| 109 | OAuthGetToken | OAuth 2.0 Client Credentials ile access token alır | ⬜ |
| 110 | WebhookListener | Belirli endpoint'i dinleyip gelen veriyi sayfaya yazar | ⬜ |

---

## 5. POWERSHELL / KOMUT SATIRI

| # | MethodName | Açıklama | Durum |
|---|-----------|----------|-------|
| 111 | RunPsCommand | PS komutu çalıştırır, stdout'u döndürür | ⬜ |
| 112 | RunPsScript | .ps1 dosyasını çalıştırır, çıktıyı hücreye yazar | ⬜ |
| 113 | RunCmdCommand | cmd.exe /c komutunu çalıştırır | ⬜ |
| 114 | GetPsOutputToSheet | PS çıktısını satır satır sayfaya yazar | ⬜ |
| 115 | SetPsExecutionPolicy | ExecutionPolicy ayarlar (RemoteSigned vb.) | ⬜ |
| 116 | GetWindowsUpdateList | Bekleyen Windows güncellemelerini listeler | ⬜ |
| 117 | InstallWindowsUpdates | PS ile güncelleme başlatır (PSWindowsUpdate modülü) | ⬜ |
| 118 | GetEventLogErrors | Son N Application/System hatasını çeker | ⬜ |
| 119 | FlushDnsCache | `ipconfig /flushdns` çalıştırır | ⬜ |
| 120 | ResetNetworkAdapter | Bağdaştırıcıyı devre dışı bırakıp tekrar etkinleştirir | ⬜ |
| 121 | GetNetworkConfig | IP, DNS, Gateway, DHCP bilgilerini sayfaya yazar | ⬜ |
| 122 | SetStaticIp | PS ile statik IP/DNS atar | ⬜ |
| 123 | EnableRemoteDesktop | RDP kaydını ve servisi aktif eder | ⬜ |
| 124 | GetFirewallRules | Aktif güvenlik duvarı kurallarını listeler | ⬜ |
| 125 | AddFirewallRule | İnbound/Outbound kural ekler | ⬜ |
| 126 | GetInstalledDrivers | Yüklü sürücüleri (InfName, Sürüm, Tarih) listeler | ⬜ |
| 127 | RestartWindowsService | Servis adına göre yeniden başlatır | ⬜ |
| 128 | GetServiceStatus | Servis durumunu (Running/Stopped) döndürür | ⬜ |
| 129 | KillProcessByName | İsme göre process sonlandırır | ⬜ |
| 130 | GetDiskHealthStatus | SMART verisini PS üzerinden alır | ⬜ |
| 131 | RunAsAdmin | Komutu yükseltilmiş (admin) PS ile çalıştırır | ⬜ |
| 132 | GetEnvironmentVariables | Tüm env değişkenlerini sayfaya yazar | ⬜ |
| 133 | SetEnvironmentVariable | Kullanıcı env değişkeni atar | ⬜ |
| 134 | GetHostsFile | `C:\Windows\System32\drivers\etc\hosts` içeriğini okur | ⬜ |
| 135 | AddHostsEntry | hosts dosyasına yeni satır ekler | ⬜ |
| 136 | GetWifiProfiles | Kayıtlı Wi-Fi profilleri ve şifrelerini listeler | ⬜ |
| 137 | ConnectWifi | Belirtilen SSID'ye bağlan | ⬜ |
| 138 | GetBitLockerStatus | Sürücülerin BitLocker durumunu kontrol eder | ⬜ |
| 139 | GetShadowCopies | Volume Shadow Copy listesini döndürür | ⬜ |
| 140 | CreateShadowCopy | Belirtilen sürücü için shadow copy oluşturur | ⬜ |

---

## 6. EXCEL / WORKBOOK / RAPORLAMA

| # | MethodName | Açıklama | Durum |
|---|-----------|----------|-------|
| 141 | SaveAllWorkbooks | Tüm açık dosyaları kaydeder | ⬜ |
| 142 | ExportSheetAsPdf | Aktif sayfayı PDF olarak dışa aktarır | ⬜ |
| 143 | ExportAllSheetsAsPdf | Her sayfayı ayrı PDF'e aktarır, isimleri sayfa adı | ⬜ |
| 144 | ImportSheetFromFile | Başka dosyadan sayfa kopyalar | ⬜ |
| 145 | ProtectAllSheets | Tüm sayfaları parola ile korur | ⬜ |
| 146 | UnprotectAllSheets | Tüm sayfa korumalarını kaldırır | ⬜ |
| 147 | RefreshAllPivotTables | Tüm pivot tabloları yeniler | ⬜ |
| 148 | RefreshAllConnections | Dış veri bağlantılarını yeniler | ⬜ |
| 149 | ConvertFormulasToValues | Seçilen aralıktaki formülleri değerle değiştirir | ⬜ |
| 150 | RemoveDuplicateRows | Yinelenen satırları siler, kaç adet silindiğini raporlar | ⬜ |
| 151 | SortSheetsByName | Sayfaları alfabetik sıralar | ⬜ |
| 152 | BatchRenameSheets | JSON parametresindeki ad eşleştirmesiyle toplu yeniden adlandırır | ⬜ |
| 153 | MergeMultipleFiles | Birden fazla Excel dosyasını tek sayfada birleştirir | ⬜ |
| 154 | SplitSheetByColumn | Sayfayı sütun değerine göre ayrı dosyalara böler | ⬜ |
| 155 | AutoFitAllColumns | Tüm sütun genişliklerini içeriğe göre ayarlar | ⬜ |
| 156 | AddWatermarkToSheet | Sayfa arka planına metin filigran ekler | ⬜ |
| 157 | SendWorkbookByEmail | MAPI / Outlook COM ile dosyayı e-posta ekine ekler | ⬜ |
| 158 | CreateSummarySheet | Tüm sayfaların A1 değerlerini özet sayfada toplar | ⬜ |
| 159 | CompressAllImages | Sayfadaki resimlerin sıkıştırma kalitesini düşürür | ⬜ |
| 160 | TableToJsonAndPost | Seçili tabloyu JSON'a çevirip API'ye gönderir | ⬜ |
| 161 | CreateNamedRangeFromSelection | Seçimi adlandırılmış bölge olarak tanımlar | ⬜ |
| 162 | GeneratePivotFromData | Parametre ile belirtilen aralıktan otomatik pivot oluşturur | ⬜ |
| 163 | ApplyConditionalFormatting | JSON kurallarına göre koşullu biçimlendirme ekler | ⬜ |
| 164 | InsertChartFromData | Belirtilen veri aralığından grafik ekler | ⬜ |
| 165 | ConvertSheetToHtmlTable | Sayfayı HTML tablo formatında dışa aktarır | ⬜ |
| 166 | BulkFindAndReplace | Tüm çalışma kitabında toplu bul-değiştir | ⬜ |
| 167 | CopySheetToNewWorkbook | Sayfayı yeni bir çalışma kitabına kopyalar | ⬜ |
| 168 | AutoNumberRows | Seçili sütunda otomatik numara verir | ⬜ |
| 169 | InsertRowAboveSelected | Seçili satırın üstüne N adet satır ekler | ⬜ |
| 170 | FreezeFirstRowAndColumn | İlk satır ve sütunu dondurur | ⬜ |

---

## 7. GÜVENLİK & LİSANS

| # | MethodName | Açıklama | Durum |
|---|-----------|----------|-------|
| 171 | CheckLicenseStatus | Registry'den lisans okur; tab'ı gösterir/gizler | ⬜ |
| 172 | ValidateMacWithServer | MAC + HWID'yi sunucuya doğrulatır | ⬜ |
| 173 | CheckFileIntegrity | Dosya MD5 hash'ini beklenen değerle karşılaştırır | ⬜ |
| 174 | DetectVirtualMachine | VM ortamında çalışılıp çalışılmadığını tespit eder | ⬜ |
| 175 | GenerateHardwareId | CPU SerialNumber + MAC → benzersiz 32 hex ID | ⬜ |
| 176 | EncryptTextXor | XOR + Base64 ile metin şifreleme | ⬜ |
| 177 | DecryptTextXor | XOR + Base64 ile metin çözme | ⬜ |
| 178 | CheckAdminRights | Yönetici haklarıyla çalışılıp çalışılmadığını kontrol eder | ⬜ |
| 179 | LockWorkbookOnExpiry | Lisans süresi dolduysa dosyayı kilitler + sunucuya bildirir | ⬜ |
| 180 | DetectCopyAndSelfDestruct | startingAddin kontrolü; ihlal varsa cleanup tetikler | ⬜ |
| 181 | AuditLogAction | Kullanıcı eylemini timestamp + mac + detail ile sunucuya loglar | ⬜ |
| 182 | ObfuscateSheetFormulas | Formülleri xlVeryHidden sayfalara taşıyarak gizler | ⬜ |
| 183 | CheckDebuggerAttached | VBA debugger'ın çalışıp çalışmadığını tespit eder | ⬜ |
| 184 | EncryptCellRange | Seçili hücreleri RC4 algoritmasıyla şifreler | ⬜ |
| 185 | DecryptCellRange | RC4 ile şifrelenmiş hücreleri çözer | ⬜ |
| 186 | BlacklistMacAddress | Kara listedeki MAC'leri sunucudan çekip erişimi engeller | ⬜ |
| 187 | TimeLimitedAccess | Belirli saat aralığı dışında dosyayı kilitler | ⬜ |
| 188 | IpWhitelistCheck | İzin verilen IP aralığında olup olmadığını kontrol eder | ⬜ |
| 189 | AntiScreenCapture | PrintScreen ve Snipping Tool'u geçici olarak engeller | ⬜ |
| 190 | WatermarkVisibleOnPrint | Baskı önizlemesinde görünür filigran ekler | ⬜ |

---

## 8. BİLDİRİM & KULLANICI ARABİRİMİ

| # | MethodName | Açıklama | Durum |
|---|-----------|----------|-------|
| 191 | ShowToastNotification | Windows 10/11 baloncuk bildirimi (PS BurntToast) | ⬜ |
| 192 | ShowProgressBar | Özel UserForm ile %0→%100 ilerleme çubuğu | ⬜ |
| 193 | ShowCustomInputForm | Çok alanlı giriş formu; sonuçları JSON döndürür | ⬜ |
| 194 | ShowYesNoCancelDialog | Üç seçenekli dialog, seçimi string döndürür | ⬜ |
| 195 | PlaySystemSound | Windows ses teması çalar (Asterisk, Critical vb.) | ⬜ |
| 196 | OpenUrlInBrowser | Varsayılan tarayıcıda URL açar | ⬜ |
| 197 | ShowSystemTrayBalloon | WScript.Shell PopUp baloncuğu | ⬜ |
| 198 | SetExcelTitleBar | Excel başlık çubuğunu özelleştirir | ⬜ |
| 199 | ShowStatusBarProgress | Durum çubuğunda % göstergesiyle uzun işlem | ⬜ |
| 200 | FlashTaskbarIcon | Görev çubuğu simgesini yanıp söndürür (dikkat çekme) | ⬜ |
| 201 | ShowRibbonCustomGroup | Dinamik olarak Ribbon'a özel grup ekler | ⬜ |
| 202 | HideRibbonCustomGroup | Eklenen özel Ribbon grubunu kaldırır | ⬜ |
| 203 | ShowFloatingToolbar | Ekranın üstünde kayan mini araç çubuğu oluşturur | ⬜ |
| 204 | ShowCountdownTimer | Geri sayım sayacı gösteren modeless UserForm | ⬜ |
| 205 | DisplayQrCode | Değeri QR kod resmi olarak hücreye ekler | ⬜ |
| 206 | AnimateStatusMessage | Durum çubuğunda kayan yazı efekti | ⬜ |
| 207 | ShowDarkModeUserForm | Koyu tema uygulanmış modern UserForm | ⬜ |
| 208 | ShowCalendarPicker | Takvim widget'ı gösteren UserForm, tarih seçtirir | ⬜ |

---

## 9. VERİ DÖNÜŞÜM & ENTEGRASYON

| # | MethodName | Açıklama | Durum |
|---|-----------|----------|-------|
| 209 | ConvertSheetToJson | Aktif sayfayı JSON array string'e çevirir | ⬜ |
| 210 | ImportJsonToSheet | JSON array'i başlık + satırlarıyla sayfaya aktarır | ⬜ |
| 211 | ConvertXmlToSheet | XML dökümanını sayfaya çevirir (MSXML2.DOMDocument) | ⬜ |
| 212 | ExportSheetToServer | Sayfayı JSON olarak sunucu API'sine POST eder | ⬜ |
| 213 | NormalizePhoneNumbers | TR format telefon numaralarını `+90 5xx xxx xx xx`'e çevirir | ⬜ |
| 214 | ValidateTCKimlik | TC kimlik algoritma doğrulaması | ⬜ |
| 215 | FormatCurrencyColumn | Para birimi sütununu `₺ 1.234,56` formatına çevirir | ⬜ |
| 216 | MergeJsonFiles | Birden fazla JSON dosyasını tek array'de birleştirir | ⬜ |
| 217 | SplitCsvByColumn | CSV dosyasını sütun değerine göre ayrı dosyalara böler | ⬜ |
| 218 | SqliteQueryToSheet | SQLite dosyasını ADO üzerinden sorgular, sonucu sayfaya yazar | ⬜ |
| 219 | ConvertBase64ToFile | Base64 string'i dosyaya çözer (ADODB.Stream) | ⬜ |
| 220 | ConvertFileToBase64 | Dosyayı Base64 string'e kodlar | ⬜ |
| 221 | CleanHtmlToText | HTML içeriğinden etiketleri temizler, düz metin döndürür | ⬜ |
| 222 | SheetToPivotJson | Sayfa verisini pivot-ready JSON formatına dönüştürür | ⬜ |
| 223 | CsvToJsonApi | CSV dosyasını okuyup her satırı API'ye POST eder | ⬜ |
| 224 | NormalizeIbanFormat | IBAN numarasını formatlı ve doğrulamalı yazar | ⬜ |
| 225 | ExtractEmailsFromSheet | Sayfa genelinde e-posta adreslerini bulup listeler | ⬜ |
| 226 | ConvertDateFormats | Farklı formatlardaki tarihleri standart ISO'ya çevirir | ⬜ |
| 227 | DeduplicateByKey | Belirtilen sütuna göre tekrarlayan satırları kaldırır | ⬜ |
| 228 | MergeColumnsWithSeparator | Birden fazla sütunu ayraçla birleştirir | ⬜ |

---

## 10. ZAMANLANMIŞ & OTOMATİK

| # | MethodName | Açıklama | Durum |
|---|-----------|----------|-------|
| 229 | ScheduleTaskOnce | Belirli tarih/saatte tek seferlik görev oluşturur | ✅ |
| 230 | ScheduleTaskDaily | Her gün belirli saatte tekrarlayan görev | ✅ |
| 231 | ScheduleTaskOnLogin | Windows girişinde çalışan görev | ✅ |
| 232 | RemoveScheduledTask | İsme göre zamanlanmış görevi siler | ✅ |
| 233 | AutoSaveWorkbook | Her N dakikada otomatik kayıt + sürümlü yedek | ✅ |
| 234 | MonitorFolderTrigger | Klasöre dosya eklenince belirtilen modülü tetikler | ✅ |
| 235 | SendDailyEmailReport | Belirlenen saatte sayfayı PDF'e çevirerek e-posta atar | ✅ |
| 236 | CleanOldBackups | N günden eski yedek dosyalarını temizler | ✅ |
| 237 | AutoUpdateModules | Sunucudan modül listesini çekip DB ile karşılaştırır, günceller | ✅ |
| 238 | HeartbeatPing | Her N dakikada MAC + versiyon ile sunucuya sinyal gönderir | ✅ |
| 239 | ScheduleTaskWeekly | Haftanın belirli günü ve saatinde tekrarlayan görev | ⬜ |
| 240 | ScheduleTaskMonthly | Ayın belirli günü ve saatinde tekrarlayan görev | ⬜ |
| 241 | SelfHealingCheck | Modüllerin hash'ini kontrol edip bozulanları yeniden indirir | ⬜ |
| 242 | WakeOnLanSchedule | WOL magic packet gönderip belirli saatte bilgisayar uyanmasını sağlar | ⬜ |
| 243 | RecurringDataSync | Belirli aralıkta Excel verisini sunucuya senkronize eder | ⬜ |
| 244 | AutoArchiveOldRows | N günden eski satırları arşiv sayfasına taşır | ⬜ |
| 245 | TriggerOnCellChange | Belirli hücre değiştiğinde uzak modül tetikler | ⬜ |
| 246 | DailyDatabaseBackup | Her gün belirtilen saatte veritabanını yedekler | ⬜ |

---

## 11. GELİŞMİŞ / UZMAN MODÜLLER

| # | MethodName | Açıklama | Durum |
|---|-----------|----------|-------|
| 247 | SelfUpdateAddin | Sunucudan yeni teklif.xlam indirir, mevcut sürümü değiştirir | ⬜ |
| 248 | InjectVbaModule | Çalışma zamanında hedef workbook'a VBA modülü enjekte eder | ⬜ |
| 249 | RemoveVbaModule | Workbook'tan modülü programatik olarak siler | ⬜ |
| 250 | RunMacroInWorkbook | Parametre workbook adındaki makroyu çalıştırır | ⬜ |
| 251 | CallDllFunction | `Declare`/`LoadLibrary` ile native DLL fonksiyonu çağırır | ⬜ |
| 252 | ReadWriteNamedPipe | Windows adlandırılmış boru (Named Pipe) üzerinden IPC | ⬜ |
| 253 | SendKeystrokes | `SendKeys` veya UI Automation ile tuş dizisi gönderir | ⬜ |
| 254 | CaptureScreenshot | PrintScreen + clipboard + kaydedilmiş PNG | ⬜ |
| 255 | ReadQrCode | QR kod resim dosyasını ZXing COM ile okur | ⬜ |
| 256 | GenerateBarcode | Code128 barkod formülü ile hücreye barkod yazar | ⬜ |
| 257 | SignPdfWithCertificate | iTextSharp COM ile PDF'e dijital imza atar | ⬜ |
| 258 | ConnectToSqlServer | ADO üzerinden SQL Server sorgusu çalıştırır | ⬜ |
| 259 | ConnectToMySql | MySQL ODBC connector ile sorgu | ⬜ |
| 260 | ReadFromExcelOneDrive | SharePoint/OneDrive URL'den Excel dosyasını okur | ⬜ |
| 261 | WatchClipboard | Pano değişimini izler, metin kopyalanınca modül tetikler | ⬜ |
| 262 | RemoteDesktopSession | RDP oturumu başlatır, uzak bilgisayarda komut çalıştırır | ⬜ |
| 263 | GeneratePdfReport | iTextSharp ile çok sayfalı PDF raporu üretir | ⬜ |
| 264 | EmbedImageInCell | URL'den indirilen resmi hücreye gömer | ⬜ |
| 265 | LoadPluginFromServer | Sunucudan DLL indirir, çalışma zamanında yükler | ⬜ |

---

## 12. VERİTABANI & SORGU İŞLEMLERİ *(Yeni)*

| # | MethodName | Açıklama | Durum |
|---|-----------|----------|-------|
| 266 | AdoQueryToSheet | ADODB.Connection ile herhangi bir ODBC/OLEDB kaynağını sorgular | ⬜ |
| 267 | ConnectToPostgres | PostgreSQL ODBC ile sorgu çalıştırır | ⬜ |
| 268 | ConnectToOracle | Oracle ODP.NET / ODBC ile sorgu çalıştırır | ⬜ |
| 269 | BulkInsertToSqlServer | Sayfadaki veriyi SQL Server'a toplu ekler (BULK INSERT) | ⬜ |
| 270 | ExportQueryResultToExcel | SQL sorgu sonucunu yeni Excel dosyasına aktarır | ⬜ |
| 271 | SyncSheetWithDatabase | Sayfa ile veritabanı tablosunu iki yönlü senkronize eder | ⬜ |
| 272 | CallStoredProcedure | SQL Server saklı yordamını parametrelerle çağırır | ⬜ |
| 273 | GetDatabaseSchema | Tablo listesi ve sütun bilgilerini sayfaya döker | ⬜ |
| 274 | ExecuteTransactionalUpdate | BEGIN/COMMIT ile atomik güncelleme işlemi yapar | ⬜ |
| 275 | MongoDbRestQuery | MongoDB Data API (REST) ile koleksiyon sorgular | ⬜ |

---

## 13. EMAİL & KOMÜNİKASYON *(Yeni)*

| # | MethodName | Açıklama | Durum |
|---|-----------|----------|-------|
| 276 | SendEmailWithOutlook | Outlook COM ile konu+alıcı+ek ile e-posta gönderir | ⬜ |
| 277 | SendEmailSmtp | SMTP/CDO.Message ile e-posta gönderir | ⬜ |
| 278 | ReadInboxEmails | Outlook InboxItems'ten son N e-postayı sayfaya listeler | ⬜ |
| 279 | ReplyToEmail | Belirtilen EntryID'li e-postayı yanıtlar | ⬜ |
| 280 | CreateOutlookAppointment | Takvim randevusu oluşturur (başlık, yer, tarih, süre) | ⬜ |
| 281 | CreateOutlookTask | Görev oluşturur, son tarih ve öncelik atar | ⬜ |
| 282 | SendBulkEmail | Sayfadaki alıcı listesine kişiselleştirilmiş toplu e-posta | ⬜ |
| 283 | ExportContactsToVcard | Outlook kişilerini .vcf formatında dışa aktarır | ⬜ |
| 284 | SendWhatsAppMessage | WhatsApp Business API ile mesaj gönderir | ⬜ |
| 285 | SendSmsViaTwilio | Twilio REST API ile SMS gönderir | ⬜ |

---

## 14. GÖRÜNTÜ & MEDYA İŞLEMLERİ *(Yeni)*

| # | MethodName | Açıklama | Durum |
|---|-----------|----------|-------|
| 286 | CaptureActiveWindow | Sadece Excel penceresinin ekran görüntüsünü alır | ⬜ |
| 287 | InsertImageFromUrl | URL'den resim indirip aktif hücreye ekler | ⬜ |
| 288 | ResizeAllImages | Sayfadaki tüm resimleri belirtilen boyuta küçültür | ⬜ |
| 289 | ExtractImagesFromSheet | Sayfadaki tüm resimleri klasöre kaydeder | ⬜ |
| 290 | GenerateQrCodeImage | QR kod PNG'sini oluşturup hücreye ekler | ⬜ |
| 291 | ConvertImageFormat | JPG/PNG/BMP dönüşümü (Shell ImageMagick veya WIA) | ⬜ |
| 292 | AddLogoToAllSheets | Tüm sayfalara şirket logosunu belirli konuma ekler | ⬜ |
| 293 | RecordMacroToGif | Makro çalışırken ekranı GIF olarak kaydeder | ⬜ |
| 294 | PlayAudioFile | .wav/.mp3 dosyasını çalar (Windows Media Player COM) | ⬜ |
| 295 | TextToSpeech | Metin okuma — SAPI.SpVoice ile Türkçe/İngilizce seslendirme | ⬜ |

---

## 15. GELİŞMİŞ OTOMASYONDELEGASYON *(Yeni)*

| # | MethodName | Açıklama | Durum |
|---|-----------|----------|-------|
| 296 | AutoFillFormFromApi | API'den gelen veriyle UserForm alanlarını otomatik doldurur | ⬜ |
| 297 | ChainModulesSequentially | Parametre listesiyle modülleri sıralı zincirleme çalıştırır | ⬜ |
| 298 | RunModuleOnAllWorkbooks | Açık tüm workbook'lara modülü uygular | ⬜ |
| 299 | ConditionalModuleRunner | Koşul sağlanırsa A modülü, yoksa B modülü çalıştırır | ⬜ |
| 300 | RetryOnFailure | Başarısız modülü N kez yeniden dener, başarısızsa loglar | ⬜ |
| 301 | ParallelModuleRunner | Birden fazla modülü sözde paralel (hızlı ardışık) çalıştırır | ⬜ |
| 302 | ModuleVersionControl | Modülün son çalışma tarihini ve versiyonunu registry'e yazar | ⬜ |
| 303 | AutoDocumentWorkbook | Tüm sayfa/formül/named-range bilgisini dokümana döker | ⬜ |
| 304 | WorkflowEngine | JSON tanımlı iş akışını adım adım çalıştırır | ⬜ |
| 305 | RemoteConfigLoader | Sunucudan yapılandırma JSON'u çekip registry'e yazar | ⬜ |

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

### WMI veri okuma
```vba
Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT * FROM Win32_SomeClass")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Alan"
    ws.Range("A1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        ws.Cells(r, 1).Value = obj.PropertyName
        r = r + 1
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
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
