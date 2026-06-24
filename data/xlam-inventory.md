# teklif.xlam — Modül Envanteri ve Geliştirme Planı

**Kaynak:** `C:\Users\onurm\AppData\Roaming\Microsoft\AddIns\teklif.xlam`  
**Son değişiklik:** 24.06.2026 · **Boyut:** ~1,3 MB · **vbaProject.bin:** ~3 MB (şifreli)

> Bu belge, eklentinin mevcut yapısını Neon uzak modül önerileri (`module-proposals.md`) ile birleştirir.

---

## Okuma durumu

| Yöntem | Sonuç |
|--------|--------|
| `vba_list_modules` (Excel COM) | ⚠️ Proje korumalı — liste alınamıyor |
| `vba_read_module` (Excel COM) | ⚠️ Çoğu modül: *project is protected* |
| **`oletools.olevba`** (dosyadan) | ✅ **54 modül kaynak kodu çıkarıldı** |
| ZIP içinden `customUI.xml` | ✅ Ribbon callback'leri |
| `data/xlam-backup/modules/` | ✅ Modül başına `.bas` / `.cls` / `.frm` |
| `data/xlam-backup/index.json` | ✅ Satır sayıları ve dosya listesi |

**Komut:** `npm run xlam:extract` → `data/files/teklif.xlam` dosyasından okur (Excel açık olması gerekmez).

> Excel COM ile canlı okuma/yazma için hâlâ: Güven Merkezi → *VBA proje nesne modeline erişime güven* + proje şifresi kaldırılmalı.

---

## Özet istatistikler

| Kategori | Adet | Toplam satır (yaklaşık) |
|----------|------|-------------------------|
| Standart modül | 18 | ~8.900 |
| Sınıf modülü | 2 | ~96 |
| UserForm | 30 | ~18.500 |
| Belge modülü (ThisWorkbook, Sayfa4) | 2 | ~281 |
| **Toplam** | **54** | **~27.800 satır VBA** |

En büyük bileşenler: **UFOPAN00** (2.820), **UFmy** (1.690), **zMrpi** (1.513), **UF2** (1.251), **zKisayol** (1.306), **UFKW** (1.214), **JsonConverter** (1.124), **zInternet** (1.015).

---

## Standart ve sınıf modülleri

| Modül | Satır | Rol (bilinen / tahmini) | registry-settings | Sunucu bağlantısı |
|-------|------:|-------------------------|:-----------------:|:-----------------:|
| **zInternet** | 1.015 | Lisans API, `RunRemoteCode`, eklenti güncelleme | ✅ | **Ana köprü** — `/api/license`, `/api/module`, `/api/download/teklif` |
| **zMrpi** | 1.513 | Teklif sunucuya gönderim (`mrpiSend` ribbon) | ⬜ | `PostDataToServer` benzeri — **incelenmeli** |
| **zKisayol** | 1.306 | Kısayol / ribbon yardımcıları | ⬜ | Muhtemelen yok |
| **JsonConverter** | 1.124 | JSON parse/stringify (Tim Hall) | ⬜ | zInternet bağımlılığı |
| **Module10** | 926 | Bilinmiyor | ⬜ | ? |
| **Module6** | 657 | PM satırları: işçilik / sarf / ambalaj / bara | ✅ | Yok |
| **Module2** | 508 | Ribbon callbacks II | ✅ | Dolaylı (Macro → form) |
| **Module3** | 478 | Yardımcı prosedürler | ✅ | Yok |
| **Module1** | 423 | Ribbon callbacks, `ribbonLoaded`, sekme görünürlüğü | ✅ | `RunRemoteCode("AutoStartOnExcelOpen")` |
| **zMalzemeListesiImport** | 351 | Malzeme listesi içe aktarma | ⬜ | Potansiyel: uzak import modülü |
| **Module8** | 278 | Bilinmiyor | ⬜ | ? |
| **Module7** | 238 | UFTH TreeView, dosya adı | ✅ | Yok |
| **zDosyaİslemleri** | 230 | İş programından teklif klasörü oluşturma | ✅ | `teklifDosyaOlustur` ribbon |
| **zLicense** | 230 | Lisans registry CRUD, VBA koruma | ✅ | `scngnr/Settings/license` |
| **Module5** | 159 | Teklif format dönüşümü / işçilik | ✅ | Yok |
| **zMailIslemleri** | 115 | E-posta işlemleri | ⬜ | Potansiyel: `SendEmail*` Neon modülleri |
| **Module9** | 99 | ListBox scroll, pano kopyala/yapıştır | ✅ | Yok |
| **Module11** | 86 | Bilinmiyor | ⬜ | ? |
| **clsSayfaIzleyici** | 76 | Sayfa izleme sınıfı | ⬜ | Potansiyel: `TriggerWebhookOnChange` |
| **zActiveWb** | 41 | Aktif workbook yolu → registry | ✅ | `nowOpenPropsFile` |
| **Module4** | 31 | UFOPAN11 görsel yardımcıları | ✅ | Yok |
| **EVN** | 20 | Olay sınıfı | ⬜ | ? |
| **ThisWorkbook** | 34 | `Auto_Open` → boot zinciri | ⬜ | `RunRemoteCode` tetikleyici |
| **Sayfa4** | 247 | `ayar` sayfası kodu | ⬜ | Yerel ayarlar |

---

## UserForm'lar (30 adet)

| Form | Satır | Not |
|------|------:|-----|
| UFOPAN00 | 2.820 | En büyük form — bakım riski yüksek |
| UFmy | 1.690 | |
| UF2 | 1.251 | Malzeme / teklif satır girişi |
| UFKW | 1.214 | Kur girişi |
| UFObara1 | 1.047 | Bara hesabı |
| UFTH | 934 | Şablonlar |
| UFDAD | 873 | Sayfa düzenleme |
| UFmd | 803 | |
| UFOPAN11 | 746 | |
| UFOSARF | 437 | Sarf |
| UserFormS1 | 595 | |
| UFDD | 630 | |
| UFFirma | 214 | Firma seçimi |
| DL1 | 137 | Malzeme dizini ayarı |
| … | … | Tam liste: `data/xlam-extract/` veya `vba_list_modules` çıktısı |

---

## Ribbon (customUI) — Sunucu ile ilişkili callback'ler

`onLoad="ribbonLoaded"` → **Module1**  
`getVisible="GetTabVisibility"` → lisans durumuna göre **C-prop** sekmesi

| Ribbon butonu | onAction | Muhtemel modül |
|---------------|----------|----------------|
| Teklifi Sunucuya Gönder | `mrpiSend` | **zMrpi** |
| Yeni Pano Teklif Dosyası | `teklifDosyaOlustur` | **zDosyaİslemleri** / Module1 |
| (açılış) | `ribbonLoaded` | **Module1** → `AutoStartOnExcelOpen` |
| Sekme görünürlüğü | `GetTabVisibility` | **Module1** + **zLicense** |

Diğer ~80+ callback (Macro32, Macro13, …) teklif iş akışına ait; Neon modülleriyle doğrudan bağlı değil.

---

## zInternet — Repo ile karşılaştırma

Canlı xlam'daki `zInternet` (~1.015 satır) ile repodaki `data/modules-source/zInternet-additions.bas` (~561 satır) **aynı değil**.

Additions dosyasında olup xlam'da **doğrulanması gereken** özellikler:

| Özellik | additions.bas | Öncelik |
|---------|---------------|---------|
| `RunRemoteCodeQuiet` | ✅ | Yüksek — arka plan agent |
| `GetHostWorkbook` (eklenti yerine ana dosya) | ✅ | Yüksek |
| `ExecuteFirmAutoStartList` / boot zinciri | ✅ | Yüksek |
| `EnsureCommandQueueQuiet` | ✅ | Yüksek |
| `FolderWatchServer_Tick` | ✅ | Orta |
| `apiBaseUrl` GetSetting desteği | Kısmi | **Kritik** — hâlâ `localhost:3000` sabiti var |
| `PrepareModuleCode` / gelişmiş JSON extract | ✅ | Orta |

**Bilinen zInternet sorunları** (`registry-settings.md`):

1. `GET_LICENSE_URL = "http://localhost:3000/api/"` — production'da Vercel URL kullanılmalı (`ilhan/Settings/apiBaseUrl`).
2. `zLicense.GetLicenseFromRegistry` boolean kontrolü — `"false"` string'i True sayılabilir.
3. Eklenti güncelleme geçici adı `kitap23.xlam` — hedef `teklif.xlam`.

---

## Geliştirme önerileri (xlam + Neon birleşik)

### Faz A — Altyapı (öncelik: kritik)

| # | Öneri | Hedef modül | Neon modülü / API |
|---|--------|-------------|-------------------|
| A1 | VBA proje kilidini geliştirme ortamında kaldır; `xlam:backup` script ekle | — | — |
| A2 | `zInternet-additions.bas` → canlı `zInternet` merge + redeploy | zInternet | — |
| A3 | `apiBaseUrl` registry'den oku; varsayılan Vercel production URL | zInternet | Tüm `/api/*` |
| A4 | Vercel `DATABASE_URL` + redeploy | — | Dashboard, `/cihazlar` |
| A5 | `zLicense` boolean düzeltmesi | zLicense | `/api/license` |

### Faz B — Sunucu entegrasyonu (ribbon + otomatik)

| # | Öneri | Hedef | Neon / API |
|---|--------|-------|------------|
| B1 | Ribbon'a **"Cihaz Bilgisi Gönder"** → `RunRemoteCode("CollectDeviceInfoServer")` | Module1 + customUI | `/api/device-info` |
| B2 | **HeartbeatPing** + **InstallTeklifAgent** auto-start listesine ekle | zInternet boot | `/api/heartbeat`, agent |
| B3 | `mrpiSend` → `PostDataToServer` ile aynı API sözleşmesini doğrula | zMrpi | `/api/license` POST |
| B4 | Komut kuyruğu UI ile Excel tarafını eşle (`InstallCommandQueue`) | zInternet | `/api/commands` |
| B5 | Klasör izleme: `FolderWatchServer_Tick` ↔ dashboard `/klasor-izleme` | zInternet | `/api/folder-watch` |

### Faz C — Operasyon / izleme

| # | Öneri | Neon modül örneği |
|---|--------|-------------------|
| C1 | Firma bazlı auto-start modülleri dashboard'dan yönet | `ExecuteFirmAutoStartList` |
| C2 | Modül çıktılarını ribbon'dan görüntüle linki | `ListModuleOutputs` |
| C3 | Lisans dışı cihazlarda sekme gizleme + uyarı mesajı iyileştirme | `TestInternetConnection` |
| C4 | Eklenti güncelleme: sürüm kontrolü + changelog | `LisansKontrolVeGuncelleme` |

### Faz D — Kod kalitesi (uzun vadeli)

| # | Öneri | Bileşen |
|---|--------|---------|
| D1 | **UFOPAN00** (2.820 satır) parçalara böl | UserForm + yardımcı modül |
| D2 | **zMrpi** / **zKisayol** / **Module10** dokümante et (şifre açılınca) | ⬜ modüller |
| D3 | Sabit `C:\Belgelerim\Cemex\...` yollarını registry'ye taşı | UF2, çeşitli |
| D4 | `JsonConverter` yerine minimal `EscJson` (sadece API katmanı) | zInternet |

---

## Neon modülleri ↔ xlam eşlemesi

| İş akışı | Excel giriş noktası | Neon modül |
|----------|---------------------|------------|
| Lisans kontrol | `TestLicenseCheck` | `getLicense` |
| Teklif verisi POST | `PostDataToServer` / `mrpiSend` | (inline zInternet) |
| Uzak modül çalıştır | `RunRemoteCode` | 686 modül DB'de |
| Cihaz envanteri | *(ribbon yok — eklenmeli)* | `CollectDeviceInfoServer` ✅ |
| Agent heartbeat | `AutoStartOnExcelOpen` zinciri | `HeartbeatPing` ✅ |
| Registry import | Manuel | `ImportRegistrySettings` ✅ |
| Klasör izleme | `FolderWatchServer_Tick` (additions) | `FolderWatchPoll` ✅ |

---

## Sonraki adımlar

1. VBA şifresini kaldırın veya geçici geliştirme kopyası (`teklif-dev.xlam`) oluşturun.
2. `node scripts/audit-modules.mjs` — Neon modül syntax (686 modül, 3 uyarı).
3. `zInternet-additions.bas` merge → AddIns'e kaydet → Excel'de test.
4. Ribbon'a B1 maddesi (Cihaz Bilgisi Gönder) ekleyin.
5. Bu belgeyi şifre açıldıkça modül modül güncelleyin.

**İlgili dosyalar:** `module-proposals.md` · `zInternet-additions.bas` · `vba-notes/registry-settings.md` · `data/xlam-extract/customUI/customUI.xml`
