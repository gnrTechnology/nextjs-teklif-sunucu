# Teklif Sunucu

Excel VBA istemcisi ile uyumlu lisans, teklif verisi ve uzak modül sunucusu.

## Kurulum

```bash
npm install
npm run dev
```

Sunucu varsayılan olarak `http://localhost:3000` adresinde çalışır.

VBA kodunuzdaki URL sabitlerini buna göre ayarlayın:

```vba
Public Const GET_LICENSE_URL As String = "http://localhost:3000/api/"
Public Const POST_LICENSE_URL As String = "http://localhost:3000/api/license/"
```

Üretim sunucusu için IP veya domain kullanın (ör. `http://31.57.154.123:3000/api/`).

## API Uç Noktaları

| Metot | URL | VBA Fonksiyonu | Açıklama |
|-------|-----|----------------|----------|
| GET | `/api/license/{mac}` | `GetLicenseStatus` | MAC ile lisans sorgula |
| POST | `/api/license/` | `RegisterLicense`, `PostDataToServer` | Kayıt / teklif verisi |
| POST | `/api/module` | `RunRemoteCode` | Uzak VBA kodu al |
| POST | `/api/download/teklif` | `LisansKontrolVeGuncelleme` | Eklenti indir |

## Veri Depolama

- Lisanslar: `data/licenses.json`
- Uzak VBA modülleri: `data/modules.json`
- Firma otomatik modüller: `data/firm-auto-modules.json`
- Eklenti dosyası: `data/files/teklif.xlam` (sizin yüklemeniz gerekir)

### Firma bazlı Excel açılış modülleri

`data/firm-auto-modules.json` dosyasında modül tanımları yapılır.

- `firmaAdi: "*"` → **tüm firmalara** uygulanan ortak modüller (`getLicense` vb.)
- Diğer kayıtlar → o firmaya **ek** modüller (tekrar yazmaya gerek yok)
- Sonuç: `*` modülleri + firma modülleri birleştirilir

```json
{ "firmaAdi": "*", "onExcelOpen": { "modules": [{ "methodName": "getLicense", "order": 1 }] } },
{ "firmaAdi": "EPRON", "onExcelOpen": { "modules": [{ "methodName": "LisansKontrolVeGuncelleme", "order": 2 }] } }
```

**Excel açılışı (teklif.xlam):**

1. `data/modules-source/zInternet-additions.bas` → `zInternet` modülüne ekleyin
2. `ThisWorkbook` modülünde:
   ```vba
   Public Sub Auto_Open()
       zInternet.RunRemoteCode "AutoStartOnExcelOpen"
   End Sub
   ```

`AutoStartOnExcelOpen` API modülü sunucudan modül listesini alır ve her biri için `zInternet.RunRemoteCode` çağırır.

### Yeni modül ekleme

`data/modules.json` dosyasına yeni kayıt ekleyin. VBA kaynakları `data/modules-source/` altında tutulabilir.

### Eklenti güncelleme (önemli)

`LisansKontrolVeGuncelleme` indirmeyi geçici kitapta yapar, kurulumu **ana teklif dosyasındaki** `TeklifBootstrap` modülüne devreder. Böylece `teklif.xlam` kapatılırken kod durmaz.

1. `data/modules-source/TeklifBootstrap.bas` dosyasını ana `.xlsm` dosyanıza import edin
2. `RunRemoteCode` ve `ExecuteDynamicFunction` kodlarını eklentiden kaldırıp ana dosyada kullanın
3. Eklenti butonundan çağırırken:
   ```vba
   Application.Run "'" & ActiveWorkbook.Name & "'!TeklifBootstrap.RunRemoteCode", "LisansKontrolVeGuncelleme"
   ```
   veya doğrudan ana dosyadan: `TeklifBootstrap.RunRemoteCode "LisansKontrolVeGuncelleme"`

```json
{
  "methodName": "SelamTest",
  "description": "Deneme modülü",
  "active": true,
  "code": "Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object\r\n    MsgBox \"selam\"\r\n    Set DynamicFunc = Nothing\r\nEnd Function"
}
```

Excel'de çağrı: `RunRemoteCode "SelamTest"`

Registry import (tüm SaveSetting değerleri):

```vba
' teklif.xlam içinden (sunucu + Next.js çalışıyor olmalı)
Application.Run "zInternet.RunRemoteCode", "ImportRegistrySettings"
```

JSON kaynağı: `GET /api/registry-settings` → `data/registry-settings-full.json`. Modül bara fiyatı, `TBveren` ve firma adını sorar; JSON'daki `EPRON` firma adıyla değiştirilir.

Kaynak: `data/modules-source/ImportRegistrySettings.bas` — `node scripts/sync-modules.js` ile `modules.json`'a aktarılır.

## Üretim

```bash
npm run build
npm start
```

Port değiştirmek için:

```bash
set PORT=3000 && npm start
```

## VBA Tarafında Dikkat Edilecekler

1. **Çift slash (RunRemoteCode)**  
   `GET_LICENSE_URL & "/module"` → `api//module` üretir. Düzeltme:
   ```vba
   .Open "POST", GET_LICENSE_URL & "module/", False
   ```

2. **GET lisans URL sonu**  
   `trailingSlash: true` ile GET isteğinin sonuna `/` ekleyin:
   ```vba
   requestUrl = GET_LICENSE_URL & "license/" & macAddress & "/"
   ```
   Alternatif: `next.config.ts` içinde `trailingSlash: false` yapıp `POST_LICENSE_URL` sonundaki `/` karakterini kaldırın.

3. **MAC URL kodlama**  
   Bazı ağ ortamlarında `:` karakteri sorun çıkarırsa:
   ```vba
   requestUrl = GET_LICENSE_URL & "license/" & Replace(macAddress, ":", "%3A") & "/"
   ```

4. **RegisterLicense sabit firma adı**  
   `Const FIRMA_ADI As String = "ABC"` yerine sayfadan okuyun (PostDataToServer gibi).

5. **Eklenti dosyası**  
   `data/files/teklif.xlam` dosyasını sunucuya yükleyin.

6. **Lisans değeri**  
   Sunucu yeni kayıtlarda `license: "true"` döner. `zLicense.GetLicenseFromRegistry` bunu Boolean olarak okuyabilmelidir.

## Uzak Modüller

Modüller `data/modules.json` dosyasından okunur. Dönen kod mutlaka `DynamicFunc(targetWb, param)` fonksiyonunu içermelidir.
