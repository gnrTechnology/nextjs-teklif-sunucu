# teklif.xlam — Geliştirme Önerileri

> **Kaynak analiz:** `data/xlam-backup/modules/` (olevba, 54 bileşen, ~27.800 satır)  
> **İlgili belgeler:** [xlam-inventory.md](./xlam-inventory.md) · [zInternet-additions.bas](./modules-source/zInternet-additions.bas) · [ui-plan.md](./ui-plan.md) · [registry-settings.md](./vba-notes/registry-settings.md)

---

## Özet

`teklif.xlam` iki katmanlı bir eklenti: **yerel teklif iş akışı** (30 UserForm, ribbon, malzeme listeleri) ve **uzak sunucu köprüsü** (`zInternet` → Neon/Vercel). Sunucu tarafı olgunlaşırken xlam tarafında en kritik eksikler **kod çıkarma hatası**, **lisans boolean mantığı** ve **ribbon’dan operasyon modüllerine erişim**dir.

---

## P0 — Hemen yapılmalı (blokaj / üretim hatası)

### R1 · `ExtractCodeFromJSON` düzeltmesi

**Sorun:** Canlı `zInternet.bas` içinde `ExtractCodeFromJSON` yalnızca string replace kullanıyor; VBA kaynağındaki `\n` / `\t` kaçışları gerçek satır sonuna dönüşünce modül **syntax error** veriyor (`CollectDeviceInfoServer`, `CaptureScreenshot` vb.).

**Çözüm:** `data/modules-source/zInternet-additions.bas` içindeki sürümü merge edin:
- Önce `JsonConverter.ParseJson` ile `code` alanını oku
- Fallback: `ExtractCodeFromJSONLegacy` (sadece parse başarısız olursa)

**Doğrulama:** Excel Immediate → `Application.Run "zInternet.RunRemoteCode", "CollectDeviceInfoServer"`

---

### R2 · Lisans sekme görünürlüğü — yanlış boolean

**Sorun:** `zInternet.TestInternetConnection` içinde:

```vba
If zLicense.GetLicenseFromRegistry Then
```

`GetLicenseFromRegistry` **String** döndürür. `"false"` metni VBA’da **True** sayılır → pasif lisanslı cihazda ribbon açık kalabilir.

**Çözüm (tercih edilen):**

```vba
Public Function IsLicenseActive() As Boolean
    Dim v As String
    v = LCase$(Trim$(GetLicenseFromRegistry()))
    IsLicenseActive = (v = "true" Or v = "1" Or v = "active" Or v = "evet")
End Function
```

`TestInternetConnection` ve `GetTabVisibility` bu fonksiyonu kullanmalı.

---

### R3 · `CaptureScreenshot` — VBA satır devam limiti

**Sorun:** Uzun PowerShell string’i tek ifadede `& _` ile birleştirilince **Run-time error 40192: Too many line continuations**.

**Durum:** Neon modülü düzeltildi (PS betiği `%TEMP%` dosyasına satır satır yazılıyor). xlam’daki `zInternet` merge sonrası uzak modül otomatik güncellenir; yerel kopya varsa silin.

---

### R4 · Geliştirme ortamı erişimi

| Engel | Etki | Öneri |
|-------|------|--------|
| VBA proje şifresi | COM ile okuma/yazma yok | Geçici `teklif-dev.xlam` şifresiz kopya |
| “VBA proje nesne modeline erişime güven” kapalı | MCP / otomasyon çalışmaz | Excel Güven Merkezi’nde açın |
| Canlı AddIns dosyası | Yanlışlıkla bozulma riski | `npm run xlam:extract` ile yedek al, sonra deploy |

---

## P1 — Sunucu entegrasyonu (1–2 hafta)

### R5 · Ribbon’a operasyon butonları

Mevcut sunucu bağlantılı callback’ler:

| Ribbon | Callback | Modül | Hedef API |
|--------|----------|-------|-----------|
| Teklifi Sunucuya Gönder | `mrpiSend` | zMrpi | Harici MRP API (JWT) |
| Yeni Pano Teklif Dosyası | `teklifDosyaOlustur` | zDosyaİslemleri | Yerel klasör |
| (açılış) | `ribbonLoaded` | Module1 | `AutoStartOnExcelOpen` |

**Eksik — eklenmeli (`customUI.xml` + `Module1`):**

| Buton | onAction | Neon modül |
|-------|----------|------------|
| Cihaz Bilgisi Gönder | `sendDeviceInfo` | `CollectDeviceInfoServer` |
| Ekran Görüntüsü Al | `sendScreenshot` | `CaptureScreenshot` |
| Açık Pencereleri Listele | `listOpenWindows` | `EnumVisibleWindows` |
| Sunucu Durumu | `openDashboard` | Tarayıcıda `/` |

Örnek callback:

```vba
Public Sub sendDeviceInfo(control As IRibbonControl)
    Application.Run "zInternet.RunRemoteCode", "CollectDeviceInfoServer"
End Sub
```

---

### R6 · Auto-start zinciri standardizasyonu

`Module1.ribbonLoaded` sırası:

1. `zInternet.TestInternetConnection`
2. `RunRemoteCode("AutoStartOnExcelOpen")`
3. `zKisayol.InitializeAddin`

**Önerilen global zincir** (dashboard `/firma-modulleri` ile uyumlu):

```
HeartbeatPing → InstallTeklifAgent → InstallCommandQueue → (firma modülleri)
```

`runOnce` modüller için registry `ilhan/AutoStart/done_*` zaten `zInternet` içinde var — dashboard’daki firma listesi ile senkron tutun.

---

### R7 · `PostDataToServer` vs `mrpiSend` — iki farklı “gönder”

| Fonksiyon | Modül | Hedef | Veri |
|-----------|-------|-------|------|
| `PostDataToServer` | zInternet | `/api/license/` POST | Sayfa3/Sayfa1 JSON (teklif meta) |
| `MrpApi_SendWorkbookForServerBuild` | zMrpi | Harici MRP `TeklifWorkbookSubmit` | Tüm workbook base64 |

**Öneri:** Ribbon etiketlerini netleştirin (“Teklif meta verisi” vs “MRP’ye tam dosya”). İleride `PostDataToServer` için ayrı `/api/teklif-meta` endpoint düşünülebilir; şu an license POST’a gidiyor.

---

### R8 · Registry anahtarları tekilleştirme

İki farklı `GetSetting` ailesi kullanılıyor:

| AppName | Kullanım |
|---------|----------|
| `ilhan` | `apiBaseUrl`, `mac`, `mdip`, `malzemedizini`, `AutoStart`, `FolderWatch` |
| `scngnr` | `zLicense` → `license` anahtarı |

**Öneri:** Yeni ayarlar `ilhan/Settings` altında; `scngnr` için migration script (`ImportRegistrySettings` Neon modülü) veya `zLicense` içinde çift okuma (önce `ilhan`, yoksa `scngnr`).

---

## P2 — Mimari ve bakım (orta vadeli)

### R9 · `zInternet` modül boyutu (~1.015 satır)

Tek modülde birleşik sorumluluklar:

- Lisans GET/POST
- `RunRemoteCode` motoru
- Firma auto-start
- Klasör izleme tick
- Boot session bayrağı
- `PostDataToServer`

**Öneri:** Parçalama (xlam içi, Neon’a taşımadan):

| Yeni modül | Taşınacak |
|------------|-----------|
| `zRemote` | `RunRemoteCode*`, `ExecuteDynamicFunction`, `ExtractCodeFromJSON` |
| `zFolderWatch` | `FolderWatchServer_*` |
| `zInternet` | Lisans, `TestInternetConnection`, `PostDataToServer` |

---

### R10 · UserForm teknik borcu

| Form | Satır | Risk | Öneri |
|------|------:|------|--------|
| UFOPAN00 | 2.819 | Çok yüksek | Alt formlara böl (UFOPAN00P1/P2 zaten var — genişlet) |
| UFmy | 1.689 | Yüksek | İş mantığını `Module*` modüllerine taşı |
| UF2 | 1.249 | Orta | Malzeme API’sini `zMalzemeListesiImport` ile birleştir |

---

### R11 · Sabit dosya yolları

`Module1` ve galeri callback’leri:

```
C:\Belgelerim\Cemex\Liste Kapakları\
C:\Belgelerim\Cemex\Resimler\
```

**Öneri:** `GetSetting("ilhan", "Settings", "cemexRoot", ...)` + kurulum sihirbazı (`UFKur` veya `Sayfa4` ayarları).

---

### R12 · `ThisWorkbook` test subs temizliği

`ThisWorkbook.cls` içinde `Sub a()`, `av()`, `b()`, `d()` — geliştirme artıkları. Üretim build’inde kaldırın veya `Const DEBUG_MODE` ile sarın.

---

### R13 · `JsonConverter` bağımlılığı

1.124 satır Tim Hall kütüphanesi — gerekli ama ağır. API katmanında minimal `EscJson` / `ParseJson` wrapper yeterli olabilir; form tarafında JsonConverter kalsın.

---

## P3 — Operasyon ve gözlemlenebilirlik

### R14 · Modül çıktısı standardı

Uzak modüllerden web’e veri için ortak sözleşme:

```json
{
  "mac": "...",
  "moduleName": "ModulAdi",
  "hostname": "...",
  "output": { "type": "...", ... }
}
```

`CaptureScreenshot` → `output.type = "screenshot"`. Diğer modüller (`EnumVisibleWindows`, `GetRunningProcesses`) için `type: "table"` + `rows: [...]` önerilir.

---

### R15 · Hata görünürlüğü

`RunRemoteCodeQuiet` hataları yalnızca `Debug.Print`. **Öneri:**
- Kritik hatalarda `Application.OnTime` ile tek seferlik tray benzeri `MsgBox` (operatör modu)
- veya Neon `module_outputs` / activity log’a `PostModuleOutput` ile özet

---

### R16 · Agent ve PollHost

`GetHostWorkbook` TeklifPollHost’u filtreliyor — doğru. Agent kurulumu sonrası:
- `EnsureCommandQueueQuiet` her heartbeat’te çalışmalı
- Excel kapalıyken agent’ın komut kuyruğunu poll ettiğini doğrulayın

---

## P4 — Güvenlik

### R17 · JWT ve MRP tokenları

`zMrpi` içinde `MrpApi_Configure(baseUrl, jwtToken)` — token kaynak kodunda veya registry’de düz metin olmamalı. Windows Credential Manager veya şifreli registry (`EncryptTextXor` Neon modülü) değerlendirin.

---

### R18 · Uzak modül çalıştırma güveni

`ExecuteDynamicFunction` sunucudan gelen kodu geçici workbook’ta çalıştırıyor. **Öneri:**
- Neon’da modül imza / allowlist (yalnızca `DynamicFunc` içeren modüller)
- Dashboard’da modül bazlı “istemcide çalıştırılabilir” bayrağı

---

### R19 · VBA proje kilidi

Şifre üretim dağıtımını korur ama CI/CD engeller. **Öneri:** Geliştirme şifresiz + release pipeline’da şifreli paket; `npm run xlam:extract` ile diff kontrolü.

---

## Neon modül ↔ xlam eşleme tablosu

| İş | xlam girişi | Neon modül | Web |
|----|-------------|------------|-----|
| Lisans | `TestLicenseCheck` | API route | `/lisanslar` |
| Uzak çalıştır | `RunRemoteCode` | 686 modül DB | `/moduller` |
| Cihaz envanteri | *(ribbon ekle)* | `CollectDeviceInfoServer` | `/cihazlar` |
| Ekran görüntüsü | *(ribbon ekle)* | `CaptureScreenshot` | `/modul-ciktilari` |
| Açık pencereler | *(ribbon ekle)* | `EnumVisibleWindows` | `/modul-ciktilari` |
| Process listesi | — | `GetRunningProcesses` | `/modul-ciktilari` |
| Heartbeat | auto-start | `HeartbeatPing` | `/heartbeats` |
| Komut kuyruğu | agent | `InstallCommandQueue` | `/komutlar` |
| Klasör izleme | `WatchFolderServer` | `FolderWatchPoll` | `/klasor-izleme` |
| Registry import | manuel | `ImportRegistrySettings` | `/kurulum` |

---

## Uygulama yol haritası

```
Hafta 1 (P0)
  ├─ R1  ExtractCodeFromJSON merge
  ├─ R2  IsLicenseActive
  ├─ R3  CaptureScreenshot (Neon — tamam)
  └─ R4  teklif-dev.xlam + xlam:extract yedek

Hafta 2 (P1)
  ├─ R5  Ribbon: Cihaz / Screenshot / Pencereler
  ├─ R6  Auto-start zinciri dashboard ile hizala
  └─ R8  Registry migration planı

Ay 1 (P2–P3)
  ├─ R9  zInternet parçalama
  ├─ R10 UFOPAN00 refactor (kademeli)
  ├─ R11 Cemex yolları → registry
  └─ R14 Modül çıktısı şeması

Sürekli
  ├─ npm run xlam:extract → diff review
  ├─ node scripts/audit-modules.mjs
  └─ xlam-inventory.md güncelle
```

---

## Hızlı komutlar

```bash
# Modülleri dosyadan çıkar (Excel kapalı olabilir)
npm run xlam:extract

# Neon modül denetimi
npm run modules:audit

# Tek modül Neon güncelle
npm run modules:upsert-one -- CaptureScreenshot data/audit/CaptureScreenshot.bas
```

---

## Sonraki belge güncellemesi

`xlam-inventory.md` modül envanterini tutar; bu dosya **ne yapılmalı** odaklıdır. Merge veya ribbon değişikliği sonrası ilgili R* maddesini ✅ işaretleyin.
