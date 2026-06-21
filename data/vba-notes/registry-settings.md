# VBA Registry (GetSetting / SaveSetting) Notları

Form ve modüller tek tek eklendikçe güncellenir.

Ana uygulama: `ilhan` / `Settings`. Ek namespace: `sercan` / `fileOpenWorkBooks` (**zActiveWb**, **zInternet**); `scngnr` / `Settings` (**zLicense** lisans).

---

## DL1 (UserForm)

| App | Section | Key | İşlem | Kullanıldığı yer | Açıklama |
|-----|---------|-----|-------|------------------|----------|
| `ilhan` | `Settings` | `malzemedizini` | **Get** | `UserForm_Initialize` | Malzeme/fiyat listeleri ana dizini → `fm1`, `TextBox1` |
| `ilhan` | `Settings` | `malzemedizini` | **Get** | `UserForm_Initialize` | `TextBox1` tekrar atanıyor (aynı key) |
| `ilhan` | `Settings` | `malzemedizini` | **Save** | `CommandButton2_Click` | `TextBox1` değeri kaydedilir |
| `ilhan` | `Settings` | `malzemedizini` | **Save** | `klasorsecm11` | `InputBox` ile seçilen `yolm1` kaydedilir |

### DL1 — Dizin yapısı (GetSetting sonrası kullanılan yollar)

```
{malzemedizini}\Malzeme Listeleri\1\*.xlsb
{malzemedizini}\Malzeme Listeleri\2\*.xlsb
{malzemedizini}\Malzeme Listeleri\3\*.xlsb
{malzemedizini}\Malzeme Listeleri\4\*.xlsb
{malzemedizini}\Otomatik Seçim\*.xlsb
```

### DL1 — Diğer notlar

- `fm1` modül seviyesinde; `Initialize` içinde `GetSetting` ile doldurulur.
- `CommandButton3`: `CBB1` işaretliyse `TextBox1\Malzeme Listeleri` → `fm1\Malzeme Listeleri` kopyalanır (`FileSystemObject.CopyFolder`).
- `ListViewp11` öğe `Key` = tam dosya yolu (`f & dosya`).
- `ListViewP11_dblClick`: dosya açar; `ListViewp11.Tag` doluysa `UFDD` formu açılır.
- `Label1`–`Label5`: ilgili alt klasörleri Explorer ile açar.

### DL1 — GetSetting özeti (tek satır)

```
GetSetting("ilhan", "Settings", "malzemedizini")
```

---

## UF2 (UserForm)

| App | Section | Key | İşlem | Kullanıldığı yer | Açıklama |
|-----|---------|-----|-------|------------------|----------|
| `ilhan` | `Settings` | `deposabitdosya` | **Get** | `Label72_Click` | Depo stok dosya adı (ör. `stok.xlsb`) |
| `ilhan` | `Settings` | `depodizini` | **Get** | `Label72_Click` | Depo dosyasının bulunduğu klasör; `yolm = depodizini & "\"` |
| `ilhan` | `Settings` | `dtxs` | **Get** | `verigirx` | Malzeme değişiminde sütun eşlemesi; `"A-C-F"` gibi `-` ile ayrılmış |

**UF2 içinde `SaveSetting` yok.**

### UF2 — GetSetting kullanım örnekleri

```vba
' Stok sorgusu (Label72)
d1m = GetSetting("ilhan", "Settings", "deposabitdosya")
yolm = GetSetting("ilhan", "Settings", "depodizini") & "\"
Workbooks.Open (yolm & d1m)

' Teklif dışı malzeme değişimi (verigirx)
dtxs = GetSetting("ilhan", "Settings", "dtxs")
ayır = Split(dtxs, "-")   ' ayır(0), ayır(1), ayır(2) → sütun harfleri
```

### UF2 — Sabit yollar (GetSetting değil, not)

| Yol | Kullanım |
|-----|----------|
| `C:\Belgelerim\Cemex\Ayarlar\Malzeme Linkleri\Malzeme Linkleri.txt` | `malzemelink` — marka link şablonları |
| `C:\Belgelerim\CEMEX\Resimler\{rs}\` | Ürün resimleri, logo (`form`, `TextBox8_Change`, `Toolbar2` vb.) |
| `C:\Belgelerim\CEMEX\Resimler\noimage.jpg` | Resim bulunamazsa varsayılan |

`rs` = malzeme listesi dosya adının ilk 3 karakteri (`Trim(Left(Replace(mlz, "+", ""), 3))`).

### UF2 — İlişkili formlar / değişkenler

- `mlz` — açık malzeme listesi kitap adı
- `dt` — ana teklif kitap adı
- `dtx` — değişim modu (0/1/2); `dtx=2` iken `verigirx` + `dtxs` devreye girer
- `UFMZ`, `UFDD`, `UF2P` — bağlı formlar
- `il` — lisans/hazırlayan kontrolü (`Sayfa3!I55555`)
- **Module9** — `Textbox_Paste` (`TextBox22`); `Textbox_Copy2` (`TextBoxPPT`); ListBox tekerlek → `HookListBoxScroll`

### UF2 — GetSetting özeti

```
GetSetting("ilhan", "Settings", "deposabitdosya")
GetSetting("ilhan", "Settings", "depodizini")
GetSetting("ilhan", "Settings", "dtxs")
```

---

## UF2P (UserForm)

**GetSetting / SaveSetting yok.**

UF2'deki ürün resmini büyük önizleme; çoklu resim (`_1` … `_10`) arasında tıklayarak gezinme.

### UF2P — Sabit yollar

| Yol | Kullanım |
|-----|----------|
| `C:\Belgelerim\CEMEX\Resimler\{rs}\{Caption}.jpg` | Ana resim (`rsileri` — `Tag=0`) |
| `C:\Belgelerim\CEMEX\Resimler\{rs}\{Caption}_{n}.jpg` | Ek resimler (`n` = 1…10; `Activate` sayar, `rsileri` gösterir) |
| `C:\Belgelerim\CEMEX\Resimler\noimage.jpg` | Hiçbiri yoksa (kod sonu, pratikte nadiren) |

`rs` = `Trim(Left(Replace(mlz, "+", ""), 3))` — UF2 ile aynı.

### UF2P — UF2 bağlantısı

| UF2 kaynağı | UF2P hedefi |
|-------------|-------------|
| `TextBox8` | `Caption` (dosya adı, `/` → boşluk) |
| `Image1.Picture` | İlk `Picture` |
| `LBR1.Tag` | `Tag` (başlangıç indeksi) |
| `Frame3_Click` | `UF2P.Show` ile açılır |

`TBRS01` = `"mevcut / toplam"` sayaç etiketi (`Tag + 1 & " / " & TBRS01.Tag`).

---

## UFADT (UserForm)

**GetSetting / SaveSetting yok.** Sabit dosya yolu yok.

Seçili satırlarda **E sütunu (5)** miktarını toplu değiştirir (`Selection` üzerinde döngü).

| Buton | İşlem |
|-------|--------|
| `CommandButton1` | `E` += 1 |
| `CommandButton2` | `E` -= 1 (seçimde herhangi bir değer ≤ 1 ise çıkış) |
| `CommandButton3` | `E` × `TBADET.Value` |
| `CommandButton4` | `E` ÷ `TBADET.Value` |

Hesaplama sırasında `ScreenUpdating = False`, `Calculation = Manual`.

---

## UFDAD (UserForm)

**GetSetting / SaveSetting yok.**

Teklif / malzeme listesi düzenleme paneli: malzeme kodları, gruplama, bölüm başlıkları, toplamlar, para birimi, imalat şablonu, çeviri dosyaları.

### UFDAD — Sabit dosya yolları

| Yol | Kullanıldığı yer | Açıklama |
|-----|------------------|----------|
| `C:\Belgelerim\Cemex\Ayarlar\Malzeme Yönetimi\Malzeme Kodları.txt` | `malzkod2`, `malzkod`, `Image442_MouseUp` | `;` ile ayrılmış kod listesi → `ListBoxMK` / `ListBoxG1` |
| `C:\Belgelerim\Cemex\Ayarlar\Sayfa Düzenleme\İmalat Dosyası.txt` | `imalatdosya1`, `Image44_MouseUp`, `Image441_MouseUp` | İmalat sütun başlıkları → `ListBox3` |
| `C:\Belgelerim\CEMEX\Çeviri Dosyaları\Kelime Dosyaları\` | `CommandButton31`, `CommandButton32` | `*.xls*` kelime dosyaları listesi / açma |
| `C:\Belgelerim\Cemex\Yeni Teklif Şablonları\Yeni Teklif-İmalat V1.2.xltx` | `CommandButton194_Click` | İmalat teklif şablonuna A:E veri aktarımı |

### UFDAD — Sayfa / kitap verisi (registry değil)

| Konum | İşlem | Açıklama |
|-------|-------|----------|
| `Sayfa3!C7` | **Read** (`Initialize`) / **Write** (`CommandButton192`) | Teklif numarası → `TextBox4` |
| `Sayfa3!I55555` | **Read** | `"Programı Hazırlayan: İlhan Şirin"` → lisanslı teklif modu |
| `Sayfa1` | Çoklu | Gruplama, format, formül, filtre, malzeme ekleme |

### UFDAD — Mod / sayfa tipi davranışı

| Koşul | Davranış |
|-------|----------|
| `CodeName` başı `TM` | `MultiPage1` sayfa 0 gizli; toolbar 1–2 gizli; `imalatdosya1` |
| `CodeName` başı `CML` | Malzeme listesi biçimlendirme (`stkurmzEK1/2`) |
| `CodeName` başı `OTM` | İsim tanımları `Sayfa3!` önekine çekilir; `Tpb = 1` |
| Lisanslı teklif (`I55555`) | `Frame41` kapalı; `formuller1`, `stkur`, `CommandButton194` aktif |

### UFDAD — Önemli alt rutinler

| Rutin | İşlev |
|-------|--------|
| `malzkod2` / `malzkod` | Malzeme kodları txt → listbox |
| `imalatdosya1` | İmalat başlık txt → `ListBox3` |
| `CommandButton198` + `adettopla1` | Sıralama, adet birleştirme, ürün grubu satırları |
| `baslıkformatlarıdüzelt` / `toplamformatlarıdüzelt` | Bölüm adı / toplam satır formatı |
| `formuller1` / `bicimler1` | Teklif satır formülleri ve biçim (`CBB1`–`CBB3`) |
| `CommandButton31`–`33` | Çeviri dosyasından bul-değiştir |
| `CommandButton194` | İmalat xltx şablonuna dönüştürme |
| `CommandButton195`–`196` | İmalat karşılaştırma renklendirme / temizleme |

---

## UFDAD0 (UserForm)

**GetSetting / SaveSetting yok.**

Seçili hücreler için hızlı araç çubuğu: para birimi formatı, metin dönüşümü, adet sayımı, aynı kod bulma, adet birleştirme.

### UFDAD0 — Toolbar butonları

| # | İşlem |
|---|--------|
| 1 | TL format `#,##0.00` |
| 2 | USD format `#,##0.00 [$$-C0C]` (kırmızı) |
| 3 | EUR format `#,##0.00 [$€-1]` (mavi) |
| 4 | Türkçe büyük harf (`ı→I`, `i→İ`) |
| 5 | `Application.Proper` |
| 6 | `Trim` + MsgBox |
| 7 | `adetsay1` — seçimde adet sayımı, txt dosyası |
| 8 | `aynıbul1` — aynı B kodunu renklendir |
| 9 | `adettopla1` — B sütununda aynı kodları birleştir, E topla |

### UFDAD0 — Toolbar menü (`ButtonMenu.Tag`)

| Tag | İşlem |
|-----|--------|
| 1 | `LOWER` |
| 2 | Karakter düzeltme (`¤→ğ`, `›→ı`, `‹→i`, `fl→ş`) |
| 3 | Satır sonu ve çift boşluk temizleme |
| 4 | `aynıbul2` — ikinci sütunla eşleştirerek aynı kodları bul |

### UFDAD0 — Dosya yolu (koşullu, registry değil)

`adetsay1` sonucu `Miktar.txt` dosyasına yazılır ve Explorer ile açılır:

| Koşul | Yol |
|-------|-----|
| Kitap kayıtlı (`ActiveWorkbook.path` dolu) | `{kitap klasörü}\Miktar.txt` |
| Kayıtsız kitap | `C:\Belgelerim\Cemex\Miktar.txt` |

Var olan dosya silinip yeniden oluşturulur.

### UFDAD0 — Notlar

- `adettopla1`: **B sütunu (2)** kod eşlemesi, **E sütunu (5)** miktar toplama; `BÖLÜM ADI/NO:` / `BÖLÜM TOPLAMI:` atlanır.
- `aynıbul1` / `aynıbul2`: eşleşen hücreler renk 38, seçili satır renk 40.
- UFDAD içindeki `adettopla1` ile aynı isimli ama farklı kapsam (UFDAD0 seçim tabanlı, UFDAD tüm sayfa).

---

## UFADA1 (UserForm)

**GetSetting / SaveSetting yok.**

Çeviri kelime dosyaları paneli — UFDAD içindeki `CommandButton31`–`33` ile aynı mantık, bağımsız form.

### UFADA1 — Sabit dosya yolları

| Yol | Kullanıldığı yer |
|-----|------------------|
| `C:\Belgelerim\CEMEX\Çeviri Dosyaları\Kelime Dosyaları\*.xls*` | `CommandButton31` — dosya listesi |
| `C:\Belgelerim\Cemex\Çeviri Dosyaları\Kelime Dosyaları\` | `Image2_MouseUp`, `CommandButton32` — klasör aç / dosya yükle |

*(Kodda `CEMEX` / `Cemex` yazımı karışık; Windows’ta genelde aynı klasör.)*

### UFADA1 — Butonlar

| Buton | İşlem |
|-------|--------|
| `CommandButton2` | `Unload Me` — formu kapat |
| `CommandButton31` | Kelime dosyalarını `ListViewp21`'e listele |
| `CommandButton32` | Seçili xls dosyasını aç; A/B sütunlarını `ListViewp31`'e yükle |
| `CommandButton33` | Aktif sayfada A→B bul-değiştir (`LookAt:=xlPart`) |
| `Image2_MouseUp` | Kelime dosyaları klasörünü Explorer'da aç |

### UFADA1 — UFDAD ilişkisi

UFDAD `MultiPage` içinde aynı çeviri akışı var; UFADA1 muhtemelen doğrudan açılan hafif sürüm.

---

## UFDD (UserForm)

Veri güncelleme: teklif dosyasına (`T3`) malzeme/çeviri listelerinden (`T2`) kod eşleştirerek fiyat ve isteğe bağlı formül aktarımı.

| App | Section | Key | İşlem | Kullanıldığı yer | Açıklama |
|-----|---------|-----|-------|------------------|----------|
| `ilhan` | `Settings` | `malzemedizini` | **Get** | `UserForm_Activate` | `fm1` — malzeme listeleri kök dizini |

**UFDD içinde `SaveSetting` yok.**

### UFDD — GetSetting kullanımı

```vba
fm1 = GetSetting("ilhan", "Settings", "malzemedizini")
```

### UFDD — Dizin yapısı (`fm1` tabanlı)

| Yol | Kullanıldığı yer |
|-----|------------------|
| `{fm1}\Malzeme Listeleri\1\*.xlsb` | `listeler` → `ListViewP1` |
| `{fm1}\Malzeme Listeleri\2\*.xlsb` | `listeler` → `ListViewP2` |
| `{fm1}\Malzeme Listeleri\3\*.xlsb` | `listeler` → `ListViewP3` |
| `{fm1}\Malzeme Listeleri\4\*.xlsb` | `listeler` → `ListViewP4` |
| `{fm1}\Malzeme Listeleri\{alt klasör}\{dosya}` | `mlzara` (Tag=1) |

### UFDD — Sabit dosya yolları

| Yol | Kullanıldığı yer |
|-----|------------------|
| `C:\Belgelerim\CEMEX\Çeviri Dosyaları\Yabancı Listeler\*.xls*` | `cevlisteler1` |
| `C:\Belgelerim\CEMEX\Çeviri Dosyaları\Yabancı Listeler\` | `dosyabicim`, `mlzara` (Tag=2) |

### UFDD — Ana akış

| Bileşen | Açıklama |
|---------|----------|
| `T2` / `T3` | Kaynak (yeni liste) / hedef (teklif) kitap adı |
| `CA1`–`CA4` | Sütun eşlemesi: kaynak kod/fiyat, hedef kod/fiyat |
| `ListViewp11.Tag` | `1` = malzeme listeleri (`B`/`F`); `2` = çeviri listeleri (`A`/`C`, metin modu) |
| `CommandButton1` | `Find` ile kod eşleştir; fiyat + renk; isteğe `formuller` |
| `CLK` | Kaynak kitabı işlem sonrası kapat |
| `CKF1` | Eşleşmede tam teklif formüllerini yaz (`formuller`) |
| `Toolbar2` Case 2 | `DL1.Show` — malzeme dizini seçimi |

### UFDD — İlişkili formlar

- **DL1** — `malzemedizini` ayarı (UFDD sadece okur; kayıt DL1'de)
- **UF2** — `ListViewp11.Tag` doluysa malzeme seçiminden açılır

### UFDD — GetSetting özeti

```
GetSetting("ilhan", "Settings", "malzemedizini")
```

---

## UFDD1 (UserForm)

**GetSetting / SaveSetting yok.**

Harici veri dosyasından (`T1`) yeni teklif şablonuna (`Yeni Teklif V1.2.xltx`) sütun eşlemesiyle aktarım; bölüm başlığı ve toplam satırları ekleme.

### UFDD1 — Sabit dosya yolu

| Yol | Kullanıldığı yer |
|-----|------------------|
| `C:\Belgelerim\CEMEX\Yeni Teklif Şablonları\Yeni Teklif V1.2.xltx` | `teklifaktar1` (`CommandButtonM1`) |

Açık kitap adı `Yeni Teklif V1.2.xltx` ise tekrar açılmaz (`WorkbookOpen`).

### UFDD1 — Sütun eşlemesi (`CATH1`–`CATH9`)

| Kontrol | Hedef sütun (yeni teklif) |
|---------|---------------------------|
| `CATH1` | B — Sipariş kodu |
| `CATH2` | C — Açıklama |
| `CATH3` | D — Üretici (zorunlu) |
| `CATH4` | E — Miktar |
| `CATH5`–`CATH7` + `CATKB1/2` | Bölüm başlığı satırı (`BÖLÜM ADI/NO:`) |
| `CATH8`–`CATH9` + `CATKT1/2` | Bölüm toplamı satırı (`BÖLÜM TOPLAMI:`) |

`CATKB` / `CATKT` değerleri: `BOŞ`, `DOLU`, `B`, `BAŞLIK`, `BÖLÜM`, `TOPLAM` vb.

### UFDD1 — Ana işlevler

| Rutin / Buton | İşlev |
|---------------|--------|
| `Toolbar1` Case 1 | Kaynak dosya seç (`T1`, `CATS1`–`CATS2`) |
| `CommandButtonA1` + `saltmalzeme1` | Üretici sütunundan marka listesi → `ListView1` (ürün grubu) |
| `ListView1_DblClick` | Marka için ürün grubu döngüsü (Şalt→Pano→Bara→İşçilik→…) |
| `CommandButtonM1` + `teklifaktar1` | Veriyi xltx şablonuna aktar; `ListView1` grubuna göre A/L formülleri |
| `CommandButtonM2` | Bölüm başlığı ve toplam satırlarını ekle (`ad_toplam_format`) |
| `CKF1` | Boş fiyat + tam teklif formülleri yaz |

### UFDD1 — UFDAD / UFDAD ilişkisi

UFDAD `CommandButton194` **İmalat** şablonu (`Yeni Teklif-İmalat V1.2.xltx`) kullanır; UFDD1 standart **Yeni Teklif V1.2.xltx** şablonunu kullanır.

---

## UFDD2 (UserForm)

> **Not:** Gönderilen kod `CommandButtonM1_Click` içinde kesilmiş (`Scripting.FileSyste…`). Aşağıdaki notlar yalnızca görünen parçaya dayanır; tam kod gelince güncellenecek.

### UFDD2 — Görünen kod (kısmi)

**GetSetting / SaveSetting:** görünen parçada yok.

| Olay | Davranış |
|------|----------|
| `CA12_Change` | `CheckBoxsa=False` iken `Range(CA12 & 1).Select`; `CA11=CA12` ise arka plan `&HC0C0FF`, değilse `&HDBE8DB` |
| `CommandButtonM1_Click` (başlangıç) | `Toolbar2.Tag = 2` ise `FileSystemObject` ile dosya işlemi (kod kesik) |

UFDD ile benzer sütun seçimi (`CA11`/`CA12`) ve `Toolbar2.Tag` modu (UFDD'de `ListViewp11.Tag=2` çeviri moduna karşılık gelebilir).

### UFDD2 — Eksik

Tam `CommandButtonM1_Click`, `UserForm_Activate`, toolbar, `GetSetting` ve sabit yol kullanımları için kodun geri kalanı gerekli.

---

## UFFirma (UserForm)

**GetSetting / SaveSetting yok.**

Firma ve ilgili kişi seçimi / kaydı; veri harici `xlsb` dosyasında tutulur (registry değil).

### UFFirma — Sabit dosya yolu

| Yol | Kullanıldığı yer |
|-----|------------------|
| `C:\Belgelerim\Cemex\Parametreler\Teklif Firma Bilgileri.xlsb` | `UserForm_Initialize` — yoksa aç, gizli pencere |

`fds = "Teklif Firma Bilgileri.xlsb"` — form kapanınca kitap kapatılır (`QueryClose`).

### UFFirma — Veri yapısı (`Teklif Firma Bilgileri.xlsb`)

| Sayfa | `ssno` | Kullanım |
|-------|--------|----------|
| `Sayfa1` | `0`, `1`, `3` | Firma listesi (`firmalar`) |
| `Sayfa2` | `2` | Alternatif firma listesi |

Satır yapısı (özet):
- Sütun 1: sıra no
- Sütun 2–8: `TBF1`–`TBF8` (firma adı, adres, tel, fax, ilgili, e-posta, tel2…)
- Sütun 16–22: `TBK1`–`TBK7` (kar oranları: pano, şalt, sarf, bara, işçi, ambalaj, nakliye)

`LBB1`–`LBB8`: Sayfa1 satır 1’den sütun başlıkları.

### UFFirma — `ssno` modları (`CommandButton1`)

| `ssno` | Hedef |
|--------|--------|
| `0` veya `1` | **UFTH** — işveren alanları (`TBisveren`, `TBisvadres`, `TBtno`, …); `CHK1=True` ise `Sayfa3` kar oranları (`Opano`, `Osalt`, `Osarf`, `Obara`, `Oisci`, `Oamb`, `Onak`) |
| Diğer | **UserFormS1** — `ListBoxFB` + etiketler |

`dt` — ana teklif kitabı adı (dışarıdan set).

### UFFirma — CRUD

| Buton | İşlem |
|-------|--------|
| `CommandButton22` | Kaydet — yeni firma / yeni ilgili / güncelle → `xlsb` Save |
| `CommandButton23` | Firma sil (tüm ilgililer) |
| `CommandButton21` | İlgili kişi sil |
| `TextBox22_KeyUp` | Firma adında arama (`B` sütunu) |
| `CommandButton24` | Form genişliği 720 ↔ 860 (kar oranları paneli) |

### UFFirma — İlişkili formlar

- **UFTH** — teklif başlığı / işveren
- **UserFormS1** — firma bilgisi listesi (alternatif mod)

---

## UFKur (UserForm)

**GetSetting / SaveSetting yok.** Sabit dosya yolu yok.

Aktif teklif kitabında **Sayfa3** döviz kurlarını düzenler.

### UFKur — Sayfa3 hücreleri

| Hücre / ad | İşlem | Açıklama |
|------------|-------|----------|
| `Eur` | Read (`Activate`) / Write (`CommandButton12`) | € kuru → `TextBoxEURO` |
| `Usd` | Read / Write | $ kuru → `TextBoxUSD` |
| `Tpbr` | Read | Teklif para birimi metni |
| `Tpb` | Write | `Tpbr` değerine göre: USD→`Usd`, EUR→`Eur`, TL→`1` |

### UFKur — Butonlar

| Buton | İşlem |
|-------|--------|
| `CommandButton12` | Kurları kaydet, `Tpb` güncelle, formu kapat |
| `CommandButton13` | Kaydetmeden kapat |

`TextBoxEURO` / `TextBoxUSD`: yalnızca rakam ve virgül (nokta otomatik virgüle çevrilir).

---

## UFKW (UserForm)

Otomatik seçim dosyasından (`sl1`) teklife (`dt`) ürün/grup aktarımı; resim önizleme **UFKWP** ile.

| App | Section | Key | İşlem | Kullanıldığı yer | Açıklama |
|-----|---------|-----|-------|------------------|----------|
| `ilhan` | `Settings` | `malzemedizini` | **Get** | `UserForm_Initialize` | `fm1` — DL1 ile aynı ana dizin (bu kodda doğrudan path birleştirmesi görünmüyor) |

**UFKW içinde `SaveSetting` yok.**

Lisans kontrolü: `Sayfa3!I55555` ≠ `"Programı Hazırlayan: İlhan Şirin"` ise `Initialize` çıkış.

### UFKW — Sabit dosya yolları

| Yol | Kullanım |
|-----|----------|
| `C:\Belgelerim\Cemex\Otomatik Seçim Dosyaları\{rs}\JPG\{ürün}.jpg` | `Resimler1`–`3` — otomatik seçim resimleri |
| `C:\Belgelerim\Cemex\Otomatik Seçim Dosyaları\{rs}\JPG\{ürün}_{n}.jpg` | Ek resimler (n=1…4) |
| `C:\Belgelerim\Cemex\Otomatik Seçim Dosyaları\{rs}\JPG\Logo.jpg` | Varsayılan logo |
| `C:\Belgelerim\CEMEX\Resimler\{rsy}\{prd}.jpg` | `Resimler4` — malzeme listesi resimleri |
| `C:\Belgelerim\CEMEX\Resimler\noimage.jpg` | Resim yoksa |

`rs` = `Left(sl1, 3)`; `rsy` = seçili ürün markasının ilk 3 harfi.

### UFKW — Ana değişkenler

| Değişken | Açıklama |
|----------|----------|
| `sl1` | Açık otomatik seçim kitabı adı; kapanınca `QueryClose` ile kapatılır |
| `dt` | Ana teklif kitabı |
| `UFKWP` | Resim paneli (`Toolbar1` buton 8) |
| `ListBoxMG` / `ListBoxMG2` | Ürün grubu / alt grup (güç vb.) |
| `ListViewA` | Seçilecek malzeme listesi |
| `LB` | Gruba eklenecek yardımcı ürünler |

### UFKW — Toolbar özeti

| # | İşlev |
|---|--------|
| 1 | Ana seçim sayfası |
| 2 | EK grup ürün seçimi (`LBdbg`) |
| 3 | `LB` → `ListViewA` gruba ekle |
| 4 | LİSTE modu (`LBdb3`) |
| 5 | Termik röle ekle/çıkar (`OTMM`) |
| 6–7 | Ürün aktif / pasif |
| 8 | `UFKWP` resim paneli aç/kapat |
| 9 | Form yüksekliği |
| 10 | Kapat |

Menü: grup adet modu, LİSTE-EK, grup çıkar, `sl1` dosyasını göster.

### UFKW — Teklife aktarım

| Rutin | Hedef |
|-------|--------|
| `gir_Click` | `ListViewA` → `dt` Sayfa1; `baslıkM1/M2`, `toplamM1`; tam formüller |
| `girEK_Click` | `LBdl3` tek satır ekleme |
| `grupekle` / `grupeksil` | Yardımcı ürünleri listeye ekle/çıkar |

Fiyatlar `Sayfa3` `Usd` / `Eur` ile TL'ye çevrilir (`toplamFiyat`, `LBGF`).

### UFKW — İlişkili formlar

- **UFKWP** — çoklu resim önizleme (`Resimler`, `Resimler1`–`4`)
- **DL1** — `fm1` kaynağı; otomatik seçim dosyası `sl1` genelde `{fm1}\Otomatik Seçim\` altından açılır (DL1 notları)

### UFKW — GetSetting özeti

```
GetSetting("ilhan", "Settings", "malzemedizini")
```

---

## UFKWP (UserForm)

**GetSetting / SaveSetting yok.**

UFKW resim paneli — otomatik seçim ve malzeme resimlerini büyük önizleme; PDF/DXF açma, yazdırma.

Resimler **UFKW** `Resimler1`–`4` tarafından yüklenir; UFKWP yalnızca gösterim ve etkileşim.

### UFKWP — Sabit dosya yolları

`rs` = `Trim(Left(sl1, 3))` — otomatik seçim kitap adının ilk 3 karakteri.

| Yol | Kullanıldığı yer |
|-----|------------------|
| `C:\Belgelerim\Cemex\Otomatik Seçim Dosyaları\{rs}\` | `Label61` — klasörü Explorer'da aç |
| `C:\Belgelerim\Cemex\Otomatik Seçim Dosyaları\{rs}\PDF\{TBRS01}.pdf` | `Label62` |
| `C:\Belgelerim\Cemex\Otomatik Seçim Dosyaları\{rs}\DXF\{TBRS01}.dxf` | `Label64` |

JPG resimleri UFKW tarafından `{rs}\JPG\` altından yüklenir (UFKW notları).

### UFKWP — Resim seçimi

| Kontrol | İşlev |
|---------|--------|
| `Image10`–`Image40` | Ana grup resimleri (LBRX1 = 1…4) |
| `Image1`–`Image4` | Alt resim / varyant seçimi → `Image0` büyük önizleme |
| `TBRS01` | Seçili ürün/resim adı (PDF/DXF dosya adı) |
| `resimboyut2` | Seçili grup kenarlık rengi + yenileme |
| `SBmk` + `resimboyut` | Form zoom (genişlik/yükseklik ölçekleme) |

### UFKWP — Diğer etiketler / butonlar

| Kontrol | İşlev |
|---------|--------|
| `Label63` | Yazıcı seç → `PrintForm` |
| `CommandButton11` | Varsayılan yazıcı `Microsoft Print to PDF` → `PrintForm` |
| `QueryClose` | UFKW toolbar buton 8 ikonunu sıfırla; `sUF = 1` |

### UFKWP — UFKW bağlantısı

UFKW `Toolbar1` buton 8 ile açılır/kapanır; `UFKW.Resimler` resimleri doldurur.

---

## UFmd (UserForm)

**GetSetting / SaveSetting yok.**

Tekliften **malzeme listesi** veya **pano listesi / pano icmal** sayfası üretir; gruplama `Malzeme Kodları.txt` ile.

### UFmd — Sabit dosya yolu

| Yol | Kullanıldığı yer |
|-----|------------------|
| `C:\Belgelerim\Cemex\Ayarlar\Malzeme Yönetimi\Malzeme Kodları.txt` | `malzkod`, `ToolbarP3` menü (Tag yok, doğrudan aç) |

`;` ile ayrılmış satırlar → `ListBoxMK`; `#` ile başlayan satırlar atlanır. `TBM1` = özel sıralama listesi (virgülle birleştirilmiş kodlar).

*(UFDAD ile aynı dosya — bkz. UFDAD notları.)*

### UFmd — Modlar

| Seçenek | Açıklama |
|---------|----------|
| `OptionButtonP31` | Teklifteki malzemeler (marka bazlı özet, `tümmalzeme` / `saltmalzeme`) |
| `OptionButtonP32` | Teklifteki panolar (`BÖLÜM ADI/NO:` satırları) |

`CodeName` başı `TM` → malzeme dosyası modu; `CheckBoxP31` gizli.

### UFmd — Filtreler (`tümmalzeme`)

| Checkbox | Filtre |
|----------|--------|
| `CheckBoxP33` | Şalt malzeme (PP- ve PM-MB hariç PM-M*) |
| `CheckBoxP34` | Pano (PP-) |
| İkisi birden | PM- (PM-MB hariç) |

`ListBoxP31` sütunları: marka/pano adı, adet, liste toplam, net toplam.

### UFmd — Aktarım (`CommandButtonP33`)

| Koşul | Rutin |
|-------|--------|
| P31 + seçili malzeme | `OMA` → `MA` → `mlzfullaktar` |
| P32 + pano listesi | `MA` → `Makro121` |
| P32 + pano icmal (`CheckBoxP321`) | `MA` → `Makro122` |

`MA` — seçime göre yeni sayfa: ad = marka önekleri veya `"Pano Listesi"` / `"Pano İcmal"`; `CodeName` → `Icmal_n`.

`mlzfullaktar` — `Sayfa1`'den marka eşleşen satırları icmal sayfasına; isteğe fiyat/iskonto (`CheckBoxP31`); `malzemegrup2` + `malzemebicim2`.

### UFmd — Diğer

- `dt` — ana teklif kitabı
- `CkarO` / `bfyt` — fiyat sütunu seçimi (UF2 ile aynı mantık)
- `ProgressBarP21` — uzun işlemlerde ilerleme çubuğu

---

## UFmy (UserForm)

İskonto yönetimi, malzeme listesi iskonto aktarımı, marka transferi, teklifte malzeme değişimi.

| App | Section | Key | İşlem | Kullanıldığı yer | Açıklama |
|-----|---------|-----|-------|------------------|----------|
| `ilhan` | `Settings` | `malzemedizini` | **Get** | `UserForm_Initialize` | `fm1` — malzeme listeleri kök dizini |

**UFmy içinde `SaveSetting` yok.**

### UFmy — `fm1` tabanlı yollar

| Yol | Kullanım |
|-----|----------|
| `{fm1}\Malzeme Listeleri\{1-4}\*.xlsb` | `malzisk`, `ComboBoxP22`, `markaaktar1/2`, `transferliste1`, `mlzdosyalar`, `CBMLZ1` |

### UFmy — Sabit dosya yolları

| Yol | Kullanım |
|-----|----------|
| `C:\Belgelerim\Cemex\Ayarlar\Malzeme Yönetimi\Malzeme Kodları.txt` | `malzkod`, `Image3`, toolbar menü |
| `C:\Belgelerim\Cemex\Transfer\` | `Image4`, `markaaktar1/2` |
| `C:\Belgelerim\Cemex\Transfer\Transfer.xlsb` | `malzemelisteleri1` (Markalar sayfası) |
| `C:\Belgelerim\Cemex\Transfer\{marka}.xlsb` | `malzemeserilisteleri1` (marka seri listesi) |
| `C:\Belgelerim\CEMEX\Parametreler\Montaj Fiyatları.xlsb` | `CommandButtonP24` — montaj fiyatı aktarımı |
| `C:\Belgelerim\CEMEX\Resimler\{ilk3}\Logo.jpg` / `{kod}.jpg` | Teklif/malzeme arama sekmesi resimleri |

### UFmy — Ana sekmeler (ToolbarP2)

| # | İşlev |
|---|--------|
| 1 | Teklif iskonto — marka/kod bazlı `ListViewP21`, `CommandButtonP23` aktarım |
| 2 | (degistir / tbmmalzemeler — malzeme listesi kapatma) |
| 3 | Malzeme arama — teklif + malzeme listesi yan yana değişim (`malzemedegisim`) |
| 4 | Malzeme listesi iskonto — `ComboBoxP22`, `CommandButtonP22`, `Frame23` |
| 5 | Malzeme kodları düzenleme (`ListBoxMK`) |
| 6–7 | Yükseklik / kapat |

### UFmy — İskonto aktarımı

| Mod | `Frame23` | Hedef |
|-----|-----------|--------|
| Teklif (`basla1`) | Gizli | `Sayfa1` sütun G — kod+marka eşleşmesi |
| Malzeme listesi (`basla2`) | Görünür | Açık `xlsb` Sayfa1 sütun G — K sütun kodu |

`CommandButtonP23` — `ListViewP21` → iskonto yaz; `SpinButtonP21` ile toplu düzenleme; `UserFormISK` çift tıklama.

### UFmy — Marka transferi (`CommandButtonP28`)

`Transfer.xlsb` veya `{marka}.xlsb` ile kod eşleştirme → malzeme listesinden fiyat/iskonto:

- `markaaktar1` — teklif `Sayfa1` güncelleme
- `markaaktar2` — icmal sayfasına ek sütunlar (karşılaştırmalı)

### UFmy — Diğer rutinler

| Rutin | İşlev |
|-------|--------|
| `teklifisk` | Teklifteki markaları listele (`ListBoxP21/P22`, `ListBoxP26`) |
| `iskaktar` / `CommandButtonP21` | Seçili marka iskontolarını `ListViewP21`'e yükle |
| `malzemegrup` | Kod → grup adı (`ListBoxMK`) |
| `malzemedegisim` | Teklif satırında malzeme değiştir (UF2 `verigir` benzeri) |
| `CommandButtonP24` | Montaj fiyatları xlsb → malzeme listesi H sütunu |

### UFmy — İlişkili formlar / değişkenler

- `dt` — ana teklif kitabı
- `UserFormISK` — iskonto detay
- `QueryClose` — açık malzeme listelerini kapat (`TextBoxP21`, `TBMLZ1`)

### UFmy — GetSetting özeti

```
GetSetting("ilhan", "Settings", "malzemedizini")
```

---

## UFMZ (UserForm)

Çoklu malzeme değişimi: teklif satırından malzeme listesi açma, **UF2** ile seçim, `LBLS1` kuyruğundan toplu değiştirme.

| App | Section | Key | İşlem | Kullanıldığı yer | Açıklama |
|-----|---------|-----|-------|------------------|----------|
| `ilhan` | `Settings` | `malzemedizini` | **Get** | `UserForm_Initialize` | `fm1` → `ComboBoxL` malzeme listeleri |

**UFMZ içinde `SaveSetting` yok.**

### UFMZ — `fm1` tabanlı yollar

| Yol | Kullanım |
|-----|----------|
| `{fm1}\Malzeme Listeleri\{1-4}\*.xlsb` | `ComboBoxL` — tam yol `List(i,1)` ile saklanır |

### UFMZ — Sabit yollar

| Yol | Kullanım |
|-----|----------|
| `C:\Belgelerim\CEMEX\Resimler\{marka ilk3}\{kod}.jpg` | `TextBox1_Change` |
| `C:\Belgelerim\CEMEX\Resimler\{marka ilk3}\Logo.jpg` | Varsayılan logo |
| `C:\Belgelerim\CEMEX\Resimler\noimage.jpg` | Hata yedeği |

`Label63` — üretici web linkleri (ABB, EATON, SIEMENS, SCHNEIDER, PHOENIX); registry değil.

### UFMZ — Ana akış

| Bileşen | Açıklama |
|---------|----------|
| `dt` | Ana teklif kitabı; `A1` rengi değişim modu (65535 / 16751103) |
| `mlz` | Açılan malzeme listesi kitabı |
| `Tbk` | Seçili teklif satırı |
| `ComboBoxL` / `CommandButton1` | Malzeme listesi aç → **UF2** göster |
| `LBLS1` | UF2 `CommandButton8` ile doldurulan değişim kuyruğu |
| `CommandButton3` + `tCdegistir` | Kuyruktaki malzemeleri teklife yaz (tek veya `CheckBoxLS1` ile tüm eşleşenler) |

`UF2.MultiPage1`: UFMZ `MultiPage1=1` ise UF2 sayfa 3, değilse sayfa 1.

### UFMZ — `tCdegistir`

`LBLS1` her satırı için teklif satırına A–I + formüller (J, K, L, … X); para birimi `LBLS1` sütun 9 (USD/EUR/YTL).

### UFMZ — İlişkili formlar

- **UF2** — malzeme seçim formu (`CommandButton8` → `LBLS1` doldurur)
- **QueryClose** — `mlz` kapat, `UF2` unload, `A1` rengi sıfırla

### UFMZ — GetSetting özeti

```
GetSetting("ilhan", "Settings", "malzemedizini")
```

---

## UFObara1 (UserForm)

Bakır bara hesabı: şalter/ana/nötr/toprak bara kg, akım kapasitesi seçimi, teklife `PM-MB` satırı aktarımı.

| App | Section | Key | İşlem | Kullanıldığı yer | Açıklama |
|-----|---------|-----|-------|------------------|----------|
| `ilhan` | `Settings` | `bara` | **Get** | `UserForm_Initialize` | Bakır birim fiyatı → `TextBoxbara` |
| `ilhan` | `Settings` | `bara` | **Save** | `CommandButton22_Click` | `TextBoxbara` kaydedilir |
| `ilhan` | `Settings` | `baralar` | **Save** | `CommandButton22_Click` | `TBSM1`–`TBSM13` değerleri `_` ile birleştirilir |
| `ilhan` | `Settings` | `baralar` | **Get** *(ölü kod)* | `Initialize` (atlanır) | `GoTo atla1` ile `baralar` okuma devre dışı |
| `ilhan` | `Settings` | `skdb` | **Get** | `baraliste1` | Şalter bara sipariş kodu öneki (`bkod`) |

### UFObara1 — Sabit dosya yolları

Kök: `C:\Belgelerim\Cemex\Ayarlar\Bakır Hesabı\`

| Dosya | Rutin |
|-------|--------|
| `Şalter Bara Ayak Uzunlukları.txt` | `salterayakuzunluk1` |
| `Şalter Bara Kesitleri.txt` | `salterbakırkesit1` |
| `Ana Bara Kesitleri.txt` | `anabakırkesit1` |
| `Nötr Bara Kesitleri.txt` | `Nbakırkesit1` |
| `Toprak Bara Kesitleri.txt` | `PEbakırkesit1` |
| `Bara Akım Taşıma Kapasiteleri.txt` | `Blist` |

`Image2_MouseUp` — klasörü Explorer'da açar.

### UFObara1 — MultiPage sekmeleri

| Sekme | İşlev |
|-------|--------|
| 0 | Şalter + ana/nötr/toprak bara kg (`salter`, `anabara`, `Nötr`, `PE`) → `TBGT` |
| 1 | Akım seçimi (`LBamper`, `LBB1/LBB2`) → `ListBoxBT` bara listesi |
| 2 | Kapasite tablosu (`Blist`) |

### UFObara1 — Teklife aktarım

| Buton | Rutin | Açıklama |
|-------|--------|----------|
| `CommandButton4` / `6` | `bakir_gir` / `baraliste1` | Seçili satıra bakır (`CKM1` → çoklu `baraliste1`) |
| `CBbakır` | — | Tüm `PM-MB` satırlarına güncel `TextBoxbara` fiyatı |

### UFObara1 — GetSetting / SaveSetting özeti

```vba
' Okuma
TextBoxbara = GetSetting("ilhan", "Settings", "bara")
bkod = GetSetting("ilhan", "Settings", "skdb")   ' baraliste1

' Kayıt (CommandButton22)
SaveSetting "ilhan", "Settings", "bara", bara
SaveSetting "ilhan", "Settings", "baralar", dizi1   ' TBSM1_TBSM2_..._TBSM13_
```

---

## UFOPAN00 (UserForm)

Pano oluşturma, pano çarpanları, montaj/sarf/ambalaj çarpanları; teklife `PP-` satırı ve `PM-MP/MS/MA/MB` aktarımı.

| App | Section | Key | İşlem | Kullanıldığı yer | Açıklama |
|-----|---------|-----|-------|------------------|----------|
| `ilhan` | `Settings` | `panomarka` | **Get** / **Save** | `Initialize` / `CBBPM1_Change` | Seçili pano markası indeksi (`CBBPM1.ListIndex`) |
| `ilhan` | `Settings` | `misi` | **Get** / **Save** | `Initialize` / `CommandButton10` | İşçilik-montaj markası (`TBimarka`) |
| `ilhan` | `Settings` | `msas` | **Get** / **Save** | `Initialize` / `CommandButton10` | Sarf markası |
| `ilhan` | `Settings` | `mbab` | **Get** / **Save** | `Initialize` / `CommandButton10` | Bara markası |
| `ilhan` | `Settings` | `mama` | **Get** / **Save** | `Initialize` / `CommandButton10` | Ambalaj markası |
| `ilhan` | `Settings` | `misia` | **Get** / **Save** | `Initialize` / `CommandButton10` | İşçilik açıklaması |
| `ilhan` | `Settings` | `msasa` | **Get** / **Save** | `Initialize` / `CommandButton10` | Sarf açıklaması |
| `ilhan` | `Settings` | `mbaba` | **Get** / **Save** | `Initialize` / `CommandButton10` | Bara açıklaması |
| `ilhan` | `Settings` | `mamaa` | **Get** / **Save** | `Initialize` / `CommandButton10` | Ambalaj açıklaması |
| `ilhan` | `Settings` | `skdi` | **Get** / **Save** | `Initialize` / `CommandButton10` | İşçilik sipariş kodu öneki |
| `ilhan` | `Settings` | `skds` | **Get** / **Save** | `Initialize` / `CommandButton10` | Sarf sipariş kodu öneki |
| `ilhan` | `Settings` | `skdb` | **Get** / **Save** | `Initialize` / `CommandButton10` | Bara sipariş kodu öneki (UFObara1 ile paylaşımlı) |
| `ilhan` | `Settings` | `skda` | **Get** / **Save** | `Initialize` / `CommandButton10` | Ambalaj sipariş kodu öneki |

Pano çarpanları ve montaj çarpanları çoğunlukla **workbook adları** (`Cphtv_{marka}`, `Cpptv_{marka}`, `Chtv`, `Cmtv`, `Cstv`, `Catv`) — registry değil.

Aşağıdaki anahtarlar **canlı export** (`data/registry-settings-export.json`, 2026-06-21) ile doğrulandı; **UFOPAN00** pano/montaj çarpan ekranlarından (`pano_gir1`, `CarpanlarPCTV`, `CarpanlarCMSTV` vb.) yazılır:

| Key | Örnek değer | Açıklama (tahmini) |
|-----|-------------|-------------------|
| `Cemex` | `12642` | Dahili sayaç / sürüm kodu |
| `mdip` | `EPRON` | Varsayılan pano marka kodu |
| `drcp` | `12` | Derinlik / çarpan parametresi |
| `drmi` | `600` | Derinlik minimum (mm?) |
| `PTA1`…`PTA12` | `DD`, `DH`, `DX`… | 12 pano tipi kısaltması |
| `ddcp` | `575` | DD tipi pano çarpanı |
| `dhcp` | `750` | DH |
| `dxcp` | `1375` | DX |
| `sdcp` | `575` | SD |
| `shcp` | `750` | SH |
| `sxcp` | `1375` | SX |
| `sacp` | `375` | SA |
| `kkcp` | `625` | KK |
| `tkcp` | `375` | TK |
| `ddf2`, `ddf3`, `ddf4` | `875`, `1125`, `1500` | DF tipi varyant çarpanları |
| `cpia`…`cpih` | `0`, `120`…`240` | Montaj **işçilik** harf A–H katsayıları |
| `cpsa`…`cpsh` | `0`, `120`…`200` | Montaj **sarf** harf A–H katsayıları |

`Chtv` / `Cmtv` workbook yorumlarına paralel; registry’de kalıcı yedek olarak duruyor olabilir.

### UFOPAN00 — Sabit dosya yolları

**Panolar** — `C:\Belgelerim\Cemex\Ayarlar\Panolar\`

| Dosya / klasör | Kullanım |
|----------------|----------|
| `{marka}\` | Marka alt klasörleri (`ListBoxMARKA`) |
| `{marka}\Pano Tip Tanımlamalar.txt` | `panotipparametreler`, `PcpanoCarpanver` |
| `{marka}\Pano Yapısal Parametreler.txt` | `panoyapıparametre` (yoksa `LOKAL\`) |
| `{marka}\{panotip}.txt` | `panodetay1` — panel seçim parametreleri |

**Montaj ve Sarf** — `C:\Belgelerim\Cemex\Ayarlar\Montaj ve Sarf\`

| Dosya | Kullanım |
|-------|----------|
| `Montaj ve Sarf Açıklamalar.txt` | `montajsarfacıklamalar` |
| `Montaj ve Sarf Çarpanlar.txt` | `Carpanlarmontajsarf`, `PcMontajCarpanver` |
| `Pano Tip Montaj ve Sarf Çarpanlar.txt` | `tipmontajsarfacıklamalar` |

**Resimler** — `C:\Belgelerim\CEMEX\Resimler\{marka veya ir}\*.gif`, `noimagepano.gif`

### UFOPAN00 — Ana sekmeler (Toolbar / MultiPage)

| Mod | Açıklama |
|-----|----------|
| Pano girişi (sayfa 0) | `isimyarat` → `gir1` — `PP-` satırı teklife |
| Pano çarpanları (sayfa 1) | `LBPC1/LBPC2`, `pano_gir1`, `CarpanlarPCTV` |
| Montaj/sarf (sayfa 2) | `LBMSC1/LBMSC2`, `iscilik_gir1`, `sarf_gir1`, `CarpanlarCMSTV` |
| Ayarlar (sayfa 3) | `CommandButton10` — PM marka/açıklama/kod kayıtları |
| Montaj listesi (`pfc=0`) | `ListBoxP1`…`P5` — malzeme listesinden pano seçimi |

### UFOPAN00 — İlişkili formlar / değişkenler

- `dt` — ana teklif; `pmlz` / `mlz` — malzeme listesi
- `UFOPAN00P1`, `UFOPAN00P2` — resim önizleme
- `UFOPAN01` — montaj detay
- `pfc` — 0 = montaj listesi modu, 1 = standart pano modu

### UFOPAN00 — GetSetting / SaveSetting özeti

```vba
' Initialize (CommandButton10 ile kaydedilenlerin hepsi)
PM1 = GetSetting("ilhan", "Settings", "panomarka")
TBimarka = GetSetting("ilhan", "Settings", "misi")
' ... msas, mbab, mama, misia, msasa, mbaba, mamaa, skdi, skds, skdb, skda

' Kayıt
SaveSetting "ilhan", "Settings", "panomarka", CBBPM1.ListIndex   ' CBBPM1_Change
' CommandButton10 → misi, msas, mbab, mama, misia, msasa, mbaba, mamaa, skdi, skds, skdb, skda
```

---

## UFOPAN00P1 (UserForm)

Malzeme listesi modunda (`UFOPAN00.pfc = 0`) pano önizleme penceresi — JPG resim ve PDF açma.

**Registry:** Yok (`GetSetting` / `SaveSetting` kullanılmıyor).

### UFOPAN00P1 — Sabit dosya yolları

| Yol | Kullanım |
|-----|----------|
| `C:\Belgelerim\CEMEX\Resimler\{ir}\{TBRS01}.jpg` | `TBRS01_Change` — sipariş koduna göre önizleme (`ir` = `Left(TBPM01, 3)`) |
| `C:\Belgelerim\CEMEX\Resimler\{ir}\{ir}.jpg` | Marka varsayılan görsel (fallback) |
| `C:\Belgelerim\CEMEX\Resimler\noimage.jpg` | Görsel yoksa |
| `C:\Belgelerim\CEMEX\PDF\{TBPM01}\{TBRS01}.pdf` | `Label62_Click` — PDF aç |

### UFOPAN00P1 — Davranış

| Olay | Açıklama |
|------|----------|
| `TBRS01_Change` | `UFOPAN00P1.Picture` güncellenir; `ir2` ile fallback sırası: sipariş JPG → marka JPG → `noimage.jpg` |
| `Label61_Click` | `C:\Belgelerim\CEMEX\Resimler\{ir}` klasörünü açar |
| `Label62_Click` | `ListBoxP5` seçili ve `TBRS01` doluysa PDF açar; yoksa MsgBox |
| `UserForm_QueryClose` | `UFOPAN00.Toolbar1` buton 6 görselini sıfırlar; `UFOPAN00.prs = 0` |

### UFOPAN00P1 — Üst form bağlantısı

- `UFOPAN00` Toolbar buton 6 ile açılır (`prs = 1`)
- `TBRS01` — `ListBoxP1` / `ListBoxP5` / `TBskd01` ile senkron (`UFOPAN00` içinden atanır)
- `UFOPAN00.TBPM01` — üretici/marka (PDF alt klasörü)
- `UFOPAN00.ListBoxP5` — PDF açmadan önce seçim kontrolü

---

## UFOPAN00P2 (UserForm)

Standart pano modunda (`UFOPAN00.pfc = 1`) çoklu JPG önizleme — pano tipi ve parametre seçimlerine göre 6 resim alanı (`R1`…`R6`).

**Registry:** Yok (`GetSetting` / `SaveSetting` kullanılmıyor).

### UFOPAN00P2 — Sabit dosya yolları

| Yol | Kullanım |
|-----|----------|
| `C:\Belgelerim\CEMEX\Resimler\{ir}\{TBn}.jpg` | `TB1_Change`…`TB6_Change` — `n=1..6`, seçime bağlı önizleme |
| `C:\Belgelerim\CEMEX\Resimler\{ir}` | `Label61_Click` — resim klasörünü açar |
| `C:\Belgelerim\CEMEX\PDF\{TBPM01}\{TBRS01}.pdf` | `Label62_Click` — PDF aç (P1 ile aynı kontrol) |

`ir` = `Left(CBBPM1, 3) & "\" & KDtip` — marka öneki + pano tip kodu alt klasörü.

### UFOPAN00P2 — Resim alanları eşlemesi

| Kontrol | Kaynak (`TBRS01_Change`) | UFOPAN00 listesi | Resim |
|---------|--------------------------|------------------|-------|
| `TB1` / `R1` | `TBtip` | `ListBoxPT` | Pano tipi |
| `TB2` / `R2` | `YPOEK1` | `ListBoxPTD1` | Ana tip detayı |
| `TB3` / `R3` | `ListBoxPTD2.Text` | `ListBoxPTD2` | Yükseklik / kapak tipi |
| `TB4` / `R4` | `KDok1` | `ListBoxPTD7` | Ön kapak |
| `TB5` / `R5` | `KDdyp1` & `KDdyp2` | `ListBoxPTD9` | Dip yan panel |
| `TB6` / `R6` | `KDiyp1` | `ListBoxPTD8` | İç yan panel |

`TPKOD` = `KDtip`; `TPKOD_Change` tüm `TB1`…`TB6` alanlarını temizler.

### UFOPAN00P2 — Davranış

| Olay | Açıklama |
|------|----------|
| `TBRS01_Change` | `UFOPAN00` pano alanlarından `TB1`…`TB6` ve `TPKOD` doldurulur (`Toolbar` buton 6, `TBskd` senkronu) |
| `TBn_Change` | İlgili listbox seçiliyse `Rn.Picture` = JPG; hata/boşta resim temizlenir |
| `UserForm_QueryClose` | `UFOPAN00.Toolbar1` buton 6 sıfırlanır; `UFOPAN00.prs = 0` |

### UFOPAN00P2 — Üst form bağlantısı

- `UFOPAN00` Toolbar buton 6, `pfc = 1` iken açılır
- `TBRS01` ↔ `UFOPAN00.TBskd` (`TBskd_Change` içinden atanır)
- P1’den farkı: tek sipariş JPG yerine pano parametrelerine göre 6 ayrı önizleme

---

## UFOPAN01 (UserForm)

Teklifteki pano satırlarında **montaj/sarf harfini** (I sütunu öneki) toplu veya seçili panolarda değiştirme.

**Registry:** Yok (`GetSetting` / `SaveSetting` kullanılmıyor).

**Sabit dosya yolu:** Yok — veri aktif teklif sayfasından (`Range`) ve `UFOPAN00` listelerinden okunur.

### UFOPAN01 — Açılış koşulu

`UFOPAN00.CBPSC1_Click` — `LBMSC2` seçili ve `LBPST1` dolu iken; `CBPSC1.BackColor = &HCAE3BF`.

### UFOPAN01 — Kontroller ve akış

| Bileşen | İşlev |
|---------|--------|
| `TBMH1` | Değiştirilecek mevcut montaj harfi (`UFOPAN00.TBMH1`) |
| `LBPSA5` | Hedef harf listesi — `UFOPAN00.LBMSC1` + `LBPSA1` yoğunluk/açıklama |
| `LBPSA3` | Etkilenecek pano tipleri (A/D/I sütunlarından) |
| `LBPSA4` | İlgili teklif satır numaraları (`LBPST1` / `TBPST1` noktalı liste) |
| `CBPC2` | Açık: çoklu seçim (`listedeara`); kapalı: tüm satırlar (`listetüm`) |
| `CommandButton14` | Harf değiştirmeyi uygula |
| `CommandButton13` | Formu kapat |

### UFOPAN01 — Ana prosedürler

| Prosedür | Açıklama |
|----------|----------|
| `ara1` | `UFOPAN00` durumunu yükler; satır modunda Caption = `LBPST1 & ". Satır Pano Harf Değiştirme"` |
| `panotipleri1` | `TBPST1` veya `LBPST1` satır listesinden `LBPSA3`/`LBPSA4` doldurur; `ListBoxKT1` ile pano tip adı eşleştirir |
| `listetüm` | `LBPSA4` tüm satırlarda `Range("I")` içinde `harf1` → `harf2` |
| `listedeara` | Yalnızca `LBPSA3`’te seçili ve A/D eşleşen satırlarda I sütunu güncellenir |
| `Initialize` | `LBPSA5` = tüm montaj çarpanları; `UFOPAN00.CBPSC1.BackColor` sıfırlanır |

Değişiklik sonrası `UFOPAN00.MontajCarpankontrol` çağrılır. Liste boşalırsa `UFOPAN00.LBMSC2.Object` güncellenir ve form kapanır.

### UFOPAN01 — Teklif sütunları

| Sütun | Kullanım |
|-------|----------|
| A | Pano referansı (`PP-…`, `-auto` kaldırılarak tip kodu) |
| D | Marka |
| I | Montaj harfi + alan (değiştirilen alan) |

### UFOPAN01 — Üst form bağlantısı

- `UFOPAN00.LBMSC1` / `LBMSC2` — montaj çarpan kaynağı
- `UFOPAN00.LBPSA1` — montaj açıklama/yoğunluk eşlemesi
- `UFOPAN00.ListBoxKT1` — pano tip tanım adları
- `UFOPAN00.LBPST1`, `TBPST1` — satır numarası listesi

---

## UFOPAN02 (UserForm)

Montajlı pano malzeme listesi (`.xlsb`) seçimi — seçilen dosya adı registry’ye yazılır ve `panogir0` çağrılır.

| App | Section | Key | İşlem | Kullanıldığı yer | Açıklama |
|-----|---------|-----|-------|------------------|----------|
| `ilhan` | `Settings` | `malzemedizini` | **Get** | `Initialize` | Malzeme kök dizini |
| `ilhan` | `Settings` | `panodizini` | **Save** | `panolistegir1` | Seçilen montajlı pano `.xlsb` dosya adı (`LBPML1`) |

`panodizini` için **Get**: `Module2.panogir0` (`UFOPAN02` ile **Save**).

### UFOPAN02 — Dosya yolu (dinamik)

| Yol | Kullanım |
|-----|----------|
| `{malzemedizini}\Malzeme Listeleri\4\*.xlsb` | `Initialize` — `*Montajlı Pano*` içeren dosyalar `LBPML1`’e eklenir |

### UFOPAN02 — Davranış

| Olay | Açıklama |
|------|----------|
| `Initialize` | `malzemedizini` okunur; montajlı pano listeleri listelenir |
| `CBPML1_Click` / `LBPML1_DblClick` | `panolistegir1` |
| `panolistegir1` | `SaveSetting panodizini` → `Call panogir0` → `Unload Me` |
| `CBPML2_Click` | Formu kapat |

### UFOPAN02 — GetSetting / SaveSetting özeti

```vba
fm1 = GetSetting("ilhan", "Settings", "malzemedizini")
SaveSetting "ilhan", "Settings", "panodizini", LBPML1.List(LBPML1.ListIndex)
```

---

## UFOPAN11 (UserForm)

Modül hesabı / pano yerleşim görselleştirme — teklif satırlarından QL/LS/M modüllerini okuyup örtü sacı ve malzeme yerleşimini çizer.

**Registry:** Yok (`GetSetting` / `SaveSetting` kullanılmıyor).

### UFOPAN11 — Sabit dosya yolları

| Yol | Kullanım |
|-----|----------|
| `C:\Belgelerim\Cemex\Ayarlar\Modül Hesabı\Pano Yüksekliği.txt` | `panoebatlar`, `Label284_Click` — standart pano yükseklikleri (`ListBoxPY` → `PDSTDBOY`) |
| `C:\Belgelerim\Cemex\Görünüşler\{malz_kod}.bmp` | `resimekle` — modül görseli |
| `C:\Belgelerim\Cemex\Görünüşler\` | `resimgen` — Shell metadata ile BMP boyutu (`rswidth`) |

### UFOPAN11 — Teklif sayfası verisi

`panoara` yalnızca `Sayfa3!I55555 = "Programı Hazırlayan: İlhan Şirin"` ise çalışır.

| Kaynak | Açıklama |
|--------|----------|
| Seçim / bölüm aralığı | `BÖLÜM ADI/NO:` … `BÖLÜM TOPLAMI:` satırları (`a`…`b`); `CBS1` açıksa doğrudan seçim |
| Sütun A, D, E, I | Malzeme kodu, marka, miktar/pad, tip (QL B00/B01, LS, M…) |
| `ListViewmzmtip` | Modüler ürünler (M ile başlayan I sütunu) |
| `QL00Dad/Yad`, `QL01Dad/Yad`, `LS00Dad` | Kompakt şalter ve akım trf adetleri |

`tekliftenpanoaktar` (`CommandButton1/3`) — listeyi temizleyip `panoara` + `boyutlandır` yeniden çalıştırır.

### UFOPAN11 — Ana hesap / çizim akışı

| Prosedür | Açıklama |
|----------|----------|
| `boyutlandır` | Sıfırlama; `moduler1`, kompakt/akım trf yerleştirme (`kompaktdikeyB00` vb.) |
| `yerlestir` | Örtü sacı (`sacekle`), modül resimleri (`resimekle`); çoklu pano (`EECT`, `PDGEN2`) |
| `panosil` / `yenile1` | Dinamik kontrolleri (`MLZ`, `EGOM`, `EGAOM`, `EGOD`) kaldırıp yeniden çizer |
| `formayarlama` | Form boyutu, `LabelPO` ölçü metni |
| `CBGN/CBGP` | Pano dış genişlik ±100 mm (`PDGEN1` / `PDGEN2`) |
| `CBN/CBP` | Seçili örtü sacı yüksekliği veya üst/alt boşluk ayarı (`CB1` düzenleme modu) |

Varsayılan ölçüler: `PDGEN1/2=600`, `QLG00=76.2`, `QLG01=106`, `QLG02=140`, `LS00=70`, `oran=3`.

### UFOPAN11 — Toolbar

| Buton | İşlev |
|-------|--------|
| 1 | `yenile1` |
| 2 | `FRP.Zoom` 100↔200 |
| 3 | `panoebatlar` — yükseklik listesi sekmesi |
| 4 | `yazdır` — `PrintForm` |
| 5 | `bilgi1` — boş modül yüzdesi |
| 6 | Kapat |

### UFOPAN11 — Dinamik kontroller

| Önek | Tür | Açıklama |
|------|-----|----------|
| `EGOM` / `EGAOM` | Image / Frame | Örtü sacı ve modül alanı |
| `EGOD` | Image | Düz örtü sacı (üst/alt boşluk) |
| `MLZ` | Image | Malzeme BMP (`Görünüşler`) |
| `FREPP1/2`, `FREPC1/2` | Frame | Pano dış/iç çerçeve |

`LBO` — seçili örtü sacı; `TBPMDB` — kalan modül sayısı.

### UFOPAN11 — Üst bağlantılar

- Teklif çalışma sayfasından (`Selection`) açılır; `UFOPAN02` → `panogir0` zinciriyle ilişkili olabilir
- `CommandButton2` — yedek modül (`YEDEK-MD-M10`) `ListViewmzmtip`'e ekler

---

## UFOSARF (UserForm)

Klemens ve kablo hesabı — bölümdeki modüler ürünlerden akım dağılımı çıkarır; kablo (`KB-auto`) ve klemens (`X1-auto`) satırlarını teklife ekler.

**Registry:** Yok (`GetSetting` / `SaveSetting` kullanılmıyor).

**Workbook adı (registry değil):** `CkarO` — liste/net fiyat kar formülü (`bfyt`); `Initialize` içinde yoksa oluşturulur.

### UFOSARF — Sabit dosya yolları

Kök: `C:\Belgelerim\Cemex\Ayarlar\Klemens ve Kablo Hesabı\`

| Yol | Kullanım |
|-----|----------|
| `Klemensler\*.txt` | `Initialize` — `ListBoxKM` marka/tip listesi |
| `Klemensler\{ListBoxKM}.txt` | `ListBoxKM_Click` — klemens tablosu (`ListBoxKT`, `TextBoxK1`…`8`) |
| `Kablolar\Kablo Kesitleri.txt` | `kablokesit` — `LabelPT` / `LabelPTA` (kesit, akım eşiği) |
| `Kablolar\Kablo Uzunluk.txt` | `kablometre` — `TextBoxPTM1`…`7` varsayılan metre |

`Image1`–`Image4` — ilgili txt dosya/klasörünü Shell ile açar.

### UFOSARF — Teklif verisi ve bölüm aralığı

`aralık` yalnızca `Sayfa3!I55555 = "Programı Hazırlayan: İlhan Şirin"` ise çalışır.

| Değişken | Açıklama |
|----------|----------|
| `a` | Bölüm başlangıç satırı (`BÖLÜM ADI/NO:` → sütun E) |
| `b` | Bölüm bitiş / ekleme satırı (`BÖLÜM TOPLAMI:` veya son satır+1) |
| `dt` | Ana teklif workbook |
| `sk1` | Klemens/kablo fiyat listesi workbook (`formuller` buradan eşleştirir) |
| `mdt` | Toplam modül adedi (`TextBoxMDT`) |

`modulara` — `a`…`b` arasında I sütunu `M` ile başlayan satırları akıma göre `TextBoxM1`…`M7` dağıtır (1A–25A → M1, 32A → M2, … 125A → M6).

### UFOSARF — Teklife aktarım

| Buton / olay | Prosedür | Teklif satırı |
|--------------|----------|---------------|
| `CBkablo_Click` | `gir1` | `A=KB-auto`, `B=LabelPTA.ControlTipText`, `E=PTM×M` |
| `CBklemens_Click` | `gir2` | `A=X1-auto`, `B=TextBoxK.ControlTipText`, `E=M` |
| `CommandButton1` | `aralık` + `modulara` + `kablometre` | Yeniden hesapla |
| `formuller` | `sk1` içinde `Find(kod)` ile fiyat/formül kopyası | C,D,F,G,H,I,J…X |

`gir1`/`gir2` yalnızca `mdt > 0` iken çalışır. Satırlar `b` konumuna `Insert` ile eklenir.

### UFOSARF — Diğer kontroller

| Kontrol | İşlev |
|---------|--------|
| `TextBoxM1`…`M8` | Modül/adet; `uf_modulara` toplamı günceller |
| `SBK1`…`SBK8` | Klemens tipi seçimi (`ListBoxKT`) |
| `SBmk` Spin | `TextBoxPTM1`…`7` ±0.2 m |
| `CommandButtonkd1` | `sk1` workbook’u görünür yapar |
| `QueryClose` | `sk1` kapatılır, `dt` aktif, `sl1` temizlenir |

### UFOSARF — Üst bağlantılar

- `UFOSARF.Caption` — bölüm adı (`aralık` içinden)
- UFOPAN11 ile aynı `Sayfa3!I55555` lisans/imza kontrolü
- Sarf/malzeme fiyatları `sk1` dosyasından; teklif `Sayfa1`’e yazılır

---

## UFTH (UserForm)

Teklif dosya yöneticisi — klasör ağacı (`TreeView1`), teklif meta verileri (`Sayfa3`), kopyala/kaydet, döviz kuru ve kar çarpanı uygulama.

| App | Section | Key | İşlem | Kullanıldığı yer | Açıklama |
|-----|---------|-----|-------|------------------|----------|
| `ilhan` | `Settings` | `teklifdizini` | **Get** / **Save** | `Initialize` / `klasorsec1`, `Toolbar1_ButtonMenuClick` (Tag 1–2) | Teklif `.xlsx` kök klasörü (`TextBox1`, `f`) |
| `ilhan` | `Settings` | `tbnoek` | **Get** / **Save** | `Initialize`, `toolbarbutton2` / `CommandButton3` | Teklif no öneki (`TBnoek1`) |
| `ilhan` | `Settings` | `sonteklif` | **Get** / **Save** | `Initialize` (`Labels1`) / `dosyaac` | Son açılan teklif tam yolu |
| `ilhan` | `Settings` | `TBveren` | **Get** / **Save** | `CommandButton10` / `CommandButton9` | Hazırlayan adı (`TBveren`) |

**Workbook adları (registry değil):** `CkarO` (liste/net kar), `Opsac` (PS- satırları için L sütunu).

### UFTH — Sabit / dinamik yollar

| Yol | Kullanım |
|-----|----------|
| `{teklifdizini}\` … | `TreeView1` — alt klasörler ve `.xlsx` teklifler |
| `C:\Belgelerim\Cemex\Parametreler\Teklif Firma Bilgileri.xlsb` | `CommandButton8` → `UFFirma` |
| `ActiveWorkbook.path` | Mevcut dosya konumu (`TByol`) |
| `Labels1` / `sonteklif` | Son dosya kısayolu (`Labels1_Click` ile aç) |

**Harici URL:** TCMB `https://www.tcmb.gov.tr/kurlar/today.xml` (`CBenpara`); `CommandButtontcmb` IE ile TCMB sayfası.

### UFTH — Sayfa3 alanları (`Veri_al` / `uygula_bilgiler`)

| Alan | Hücre / ad |
|------|------------|
| Proje adı, adres, işveren | C3–C5, E5 |
| Teklif no, tarih | C7–C8 |
| Hazırlayan, durum | C10, F10 |
| İlgili, tel, fax, e-posta | C11–C13, D12, F12 |
| Adam/saat, para birimi, kur | `Ads`, `Tpbr`, `Usd`, `Eur` |
| Kar çarpanı | `CkarO` |

`Sayfa3!I55555` imza kontrolü — Toolbar buton 2 ve `toolbarbutton2` teklif doğrulama.

### UFTH — Toolbar özeti

| Buton | İşlev |
|-------|--------|
| 1 | Klasör görünümü (`MultiPage1=0`), TreeView |
| 2 | `toolbarbutton2` → `Veri_al`, meta sekmesi |
| 3 | `uygula` — `uygula_bilgiler` + `kar_carpan` |
| 4 | `textBOS` — alanları temizle |
| 5 | Kopyala / Farklı kaydet (`menu21` / `menu22`) |
| 6 | Kaydet (`Save` veya `menu22`) |
| 7 | Seçili workbook kapat |
| 8 | Formu kapat |

`kar_carpan` — tüm `Sayfa1` satırlarında L sütunu: `PM-MP/MS/MA/MN/MB`, `PP-`, `PS-`, diğerleri `Osalt` ile `bfyt` formülü.

### UFTH — Diğer

| Öğe | Açıklama |
|-----|----------|
| `ssno` | `1` = tam TreeView; `2` = sadece meta sekmesi |
| `dt` | `TreeView1_DblClick` / `dosyaac` ile açılan teklif adı |
| `CommandButton1` | `zMrpi.MrpApi_TeklifNextNumberDdMmYy` — MRP teklif no |
| `CommandButton8` | MRP firma içe aktar + `UFFirma.Show` |
| `teklifnoal` | Hazırlayan baş harfleri + `tsaat` veya `tbnoek` öneki |
| `ssno=0` | `UFFirma` açılmadan önce |
| `flexgrid_rc` | **Module7** `CreateCmdBar` — TreeView sağ tık: Yeni Dizin / Yeni Teklif / Yeniden Adlandır / Sil (`menu1`–`menu4`) |

### UFTH — GetSetting / SaveSetting özeti

```vba
f = GetSetting("ilhan", "Settings", "teklifdizini")
TBnoek1.Text = GetSetting("ilhan", "Settings", "tbnoek")
Labels1 = GetSetting("ilhan", "Settings", "sonteklif")
TBveren.Text = GetSetting("ilhan", "Settings", "TBveren")   ' CommandButton10

SaveSetting "ilhan", "Settings", "teklifdizini", Yol
SaveSetting "ilhan", "Settings", "tbnoek", TBnoek1.Text
SaveSetting "ilhan", "Settings", "sonteklif", td
SaveSetting "ilhan", "Settings", "TBveren", TBveren.Text
```

---

## UserFormAD (UserForm)

Teklif sayfasına **proje** veya **bölüm başlık satırı** ekleyen küçük dialog.

**Registry:** Yok (`GetSetting` / `SaveSetting` kullanılmıyor).

**Sabit dosya yolu:** Yok — doğrudan aktif sayfaya satır ekler.

### UserFormAD — Modlar

| Koşul | B sütunu metni | UI |
|-------|----------------|-----|
| `Caption = "Proje Adı / No:"` veya `Frame4.BackColor = &H96A446` | `PROJE ADI/NO:` | `TextBox20`, `Label484`, `CBPC1` gizli; `TextBox21` = `"Proje Adı"` veya `"."` |
| Diğer (varsayılan bölüm modu) | `BÖLÜM ADI/NO:` | `TextBox20` adet (E); `CBPC1` açıksa A sütununa pano ref. |

### UserFormAD — `CommandButton12_Click` (ekle)

1. `TextBox19` boşsa çıkış
2. `Selection.EntireRow.Insert`
3. C sütunu = `TextBox19`
4. Biçim: kenarlık, kalın, yükseklik 12.75; `CodeName` `TM` ile başlıyorsa `A:J`, değilse geniş teklif aralığı
5. Proje modu: font rengi 54
6. Bölüm modu: font rengi 11; `TextBox20` → E (`# Adet`); `CBPC1` + `TextBox21` → A (`Pano Ref.` veya özel metin)
7. `Unload Me`

### UserFormAD — Diğer kontroller

| Kontrol | İşlev |
|---------|--------|
| `CommandButton13` | İptal — formu kapat |
| `CommandButton14` | `TextBox19` temizle, `TextBox20 = 1` |
| `TextBox20_KeyPress` | Yalnızca rakam |
| `UserForm_Activate` | Moda göre kontrol görünürlüğü / `TextBox21` metni |

### UserFormAD — Üst bağlantı

Çağıran kod `UserFormAD.Caption` veya `Frame4.BackColor` ile modu belirler; `UFOPAN11` / teklif `Sayfa1` bölüm yapısı (`BÖLÜM ADI/NO`, `BÖLÜM TOPLAMI`) ile uyumludur.

---

## UserForm2 (UserForm)

Yazdırma / önizleme onay dialogu — seçili sayfaları doğrudan yazdırır veya önizleme açar.

**Registry:** Yok (`GetSetting` / `SaveSetting` kullanılmıyor).

**Sabit dosya yolu:** Yok.

### UserForm2 — Kontroller

| Buton | İşlev |
|-------|--------|
| `CommandButton12` | Formu kapat → `ActiveWindow.SelectedSheets.PrintOut` |
| `CommandButton13` | İptal — `Unload Me` |
| `CommandButton14` | Önizleme: `UFmd` gizle, `PrintPreview`, `Sayfa1` seç, `UFmd` tekrar göster |

`UserForm_QueryClosexx` — yorum satırı / kullanılmıyor (`printreset` yorumlu).

### UserForm2 — Üst bağlantı

- `UFmd` — malzeme listesi / çıktı formu; önizleme sırasında geçici gizlenir
- `UserForm2` — `CommandButton14` içinde `Unload` (muhtemelen önizleme öncesi kapatma)

---

## UserFormS1 (UserForm)

Teklif / malzeme listesini **Excel şablonuna** (`Şablonlar`) aktararak müşteri teklif dosyası oluşturur (Kapak, İcmal, Detaylı Liste).

**Registry:** Yok (`GetSetting` / `SaveSetting` kullanılmıyor).

### UserFormS1 — Sabit dosya yolları

Kök: `C:\Belgelerim\Cemex\Şablonlar\`

| Yol | Kullanım |
|-----|----------|
| `Şablonlar\` (alt klasörler) | `Initialize` → `ListBoxDS` kategori listesi |
| `Şablonlar\{dd}\*.xltx` | `ListBoxDS_Click` → `ListViewP1` şablon listesi |
| `Şablonlar\{dd}\{şablon}.xltx` | `teklifyap1` / `teklifyap2` — şablon açılır (`ds1`) |
| `Şablonlar\{dd}\[{şablon}]` | `ListViewP1_Click`, `dosyabicim` — Excel4 makro ile `Veriler`, `Kapak` okuma |

Çıktı kaydı: `{kaynak klasörü}\{kaynak adı} - {ListBoxDS}.xlsx` (`ListViewP1_dblClick`).

### UserFormS1 — Veri kaynakları

| Sayfa | Kullanım |
|-------|----------|
| `Sayfa1` | `teklifyap1` — teklif satırları kopyalanır |
| `Sayfa3` | Para birimi (`Tpbr` → `p` format), `teklifveri1` meta |
| `Notlar` | `teklifveri1` / `mlzveri1` — not satırları `Veriler` sayfasına |
| Malzeme sayfası (`S.No` başlıklı) | `teklifyap2` — `ListBoxDS.listIndex > 0` |

### UserFormS1 — Ana akış

| Olay | Açıklama |
|------|----------|
| `ListBoxDS_Click` | Kategori seçimi; `MultiPage1` sekmesi; `ssno` 2/3 |
| `ListViewP1_Click` | Şablon seçimi (`LBSAD`); şablondan dil (`Veriler!R19C2`) |
| `ListViewP1_dblClick` | `teklifyap1` (index 0) veya `teklifyap2` (malzeme); `sayfasay`, `dilbul`; `SaveAs` |
| `CommandButtonsd1` | Seçili şablon dosyasını Shell ile aç |
| `CommandButton1` | `UFFirma.Show` |

### UserFormS1 — `teklifyap1` (teklif → şablon)

1. Şablon `.xltx` aç → `ds1`
2. `teklifveri1` — `Sayfa3` + `Notlar` → gizli `Veriler` sayfası
3. `DETAYLI LİSTE` — `Sayfa1` A:E, W:X kopyala
4. `düzenleF0` — sıra no, bölüm renkleri, isteğe bağlı fiyat (`CBox1`), montaj pano dağıtımı (`CBox6`), pano ref (`CBox5`)
5. `panolisteleF0` — `İCMAL` sayfası
6. `fytsilF1` — fiyat sütunları gizleme (`CBox1=False`)
7. `bicim1` — boş satır / toplam gizleme (`CBox2`, `CBox3`)

### UserFormS1 — `teklifyap2` (malzeme listesi → şablon)

- `S.No` başlıklı sayfa gerekli
- `mlzveri1` — `ListBoxFB` / `LabelFAD` firma bilgileri + `Sayfa3`
- Detay sayfasına malzeme satırları kopyala; `CBoxkdv` KDV satırları; `CBox212` iskonto sütunları

### UserFormS1 — Seçenek kutuları (özet)

| Checkbox | Etki |
|----------|------|
| `CBox1` | Birim fiyat + tutar sütunları |
| `CBox2` | Bölüm sonrası boş satır |
| `CBox3` | Bölüm toplamı gizleme |
| `CBox4` | Açıklama satır kaydırma |
| `CBox5` | Bölüm adına pano ref. ekleme |
| `CBox6` | Montaj tutarını pano satırlarına dağıt |
| `CBoxkdv` | %18 KDV + genel toplam |
| `CBox212` | İskonto sütunları (malzeme modu) |
| `CBox21` | `Veriler` sayfasını görünür/gizli yap (`Visible`) |

### UserFormS1 — `Veriler` düzenleme paneli (`TBAA1`…`TBAA9`)

Şablon çıktısındaki gizli `Veriler` sayfası meta alanlarını form üzerinden düzenler.

| Kontrol | İşlev |
|---------|--------|
| `Initialize` | `Veriler` B1:B9 → `TBAA1`…`TBAA9`; C8/D8 → `TBAAB8`/`TBAAC8`; sayfa görünürse `CBox21=True` |
| `CBOK1` | Form değerlerini `Veriler` B1:B9, C8, D8 hücrelerine yazar |
| `CBIP1` | Formu kapat |
| `CBox21_Click` | `Veriler` sayfası `Visible` aç/kapa |
| `Label251` | `TBAA5` = şu anki tarih/saat |
| `Label252` | `TBAA4` = `"TD-"` + `DDMMYY-hhmmss` |
| `Label253` | `DOSYAADI_R1` — **Module7**; `TBAA4` ile `.xlsx` yeniden adlandırma |

`Veriler` satır eşlemesi (`teklifveri1` ile uyumlu): proje, adres, işveren, teklif no, tarih, hazırlayan, ilgili, telefon(lar)/fax, e-posta.

### UserFormS1 — `dilbul`

Şablon dili `İNGİLİZCE` ise sayfa adları ve metinler İngilizceye çevrilir (`COVER`, `SECTION TOTAL`, vb.).

### UserFormS1 — Diğer

- `ssno` — form modu (`Initialize`: 2; `ListBoxDS`: 2 veya 3)
- `UserFormS1.Caption` — `teklifyap1/2` sırasında ilerleme metni
- `dosyabicim` — şablon sütun başlıklarını MsgBox ile gösterir
- `ProgressBarP21` — `düzenleF0`, `panolisteleF0`, `bicim1`

### UserFormS1 — Üst bağlantılar

- `UFFirma` — firma bilgileri (`mlzveri1` için `ListBoxFB`)
- Kaynak workbook: `DT1`; üretilen: `ds1`
- `UFTH` / teklif `Sayfa3` para birimi formatı

---

## Module1 (Standart modül — Ribbon callbacks)

Custom UI şeridi (`IRibbonUI`) — malzeme listesi galerileri, teklif şablonu, sekme görünürlüğü, MRP entegrasyonu.

| App | Section | Key | İşlem | Kullanıldığı yer | Açıklama |
|-----|---------|-----|-------|------------------|----------|
| `ilhan` | `Settings` | `malzemedizini` | **Get** | `rxgal_getItemCount*`, `OnActionCallback*`, `rxbtn_Click2/4`, `Anasalt` | Malzeme kök dizini → `fm1` |

**Registry Save:** Yok (bu modülde).

### Module1 — Sabit / dinamik yollar

| Yol | Kullanım |
|-----|----------|
| `{malzemedizini}\Malzeme Listeleri\1` | Galeri 0 — `rxgal_*`, `OnActionCallbackscn` → `macro_01` → `UF2` |
| `{malzemedizini}\Malzeme Listeleri\2` | Galeri 1 — `OnActionCallback1` |
| `{malzemedizini}\Malzeme Listeleri\3` | Galeri 2 — `OnActionCallback2` |
| `{malzemedizini}\Malzeme Listeleri\4` | Galeri 3 — `OnActionCallback3`; `rxbtn_Click2` Explorer |
| `{malzemedizini}\Otomatik Seçim` | Galeri 4 — `OnActionCallback4` → `macro_02` → `UFKW`; `rxbtn_Click4` Explorer |
| `C:\Belgelerim\Cemex\Liste Kapakları\{dosya}.jpg` | `rxgal_getItemImage` |
| `C:\Belgelerim\Cemex\Resimler\{3harf}\logo.jpg` | Kapak yoksa fallback |
| `C:\Belgelerim\Cemex\Resimler\noimage.jpg` | Görsel yoksa |
| `C:\Belgelerim\CEMEX\Yeni Teklif Şablonları\Yeni Teklif V1.2.xltx` | `MY1` — yeni teklif |

### Module1 — Genel değişkenler

| Değişken | Açıklama |
|----------|----------|
| `dt` | Aktif teklif workbook adı |
| `mlz` / `b` | Açık malzeme listesi `.xlsb` |
| `sl1` | Otomatik seçim listesi dosyası |
| `fm` / `fm1` | Malzeme klasörü / kök dizin |
| `sUF` | `UFKW` mod bayrağı (`macro_02`) |
| `t` | `MyToggleMacro` — malzeme seçim modu |
| `Rib` | Ribbon UI nesnesi |
| `teklifTabVisible` | Özel `teklif` sekmesi görünürlüğü |

### Module1 — Ribbon prosedürleri

| Prosedür | Açıklama |
|----------|----------|
| `ribbonLoaded` | **zInternet** `RunRemoteCode("AutoStartOnExcelOpen")`; `Rib` ata |
| `macro_01` | `.xlsb` aç (gizli), `dt` aktif, `UF2.Show` |
| `macro_02` | Otomatik seçim `.xlsb` aç → `UFKW.Show` |
| `Macro76` | `DL1.Show` — malzeme dizini |
| `MY1` | Yeni teklif şablonu aç; `dt` güncelle; `Rib.Invalidate` |
| `ilhan` / `msgteklif` | `Sayfa3!I55555` teklif doğrulama |
| `Anasalt` | Ana malzeme listesine dön (`asalt`) |
| `GetTabVisibility` / `ShowCustomTab` / `HideCustomTab` | `teklif` ribbon sekmesi (**zInternet** `TestInternetConnection` sonrası) |
| `teklifDosyaOlustur` | **zDosyaİslemleri** `teklif_klasor_olustur` |
| `mrpiSend` | `MrpApi_SendWorkbookForServerBuild` |
| `kapat` | Tüm `.xlam` workbook’ları kapat |

### Module1 — Galeri callback’leri

Her klasör için: `getItemCount`, `getItemLabel`, `getItemScreentip`, `getItemImage` (0), `OnActionCallback` — `Dir("*.xlsb")` ile indeks seçimi.

### Module1 — GetSetting özeti

```vba
fm1 = GetSetting("ilhan", "Settings", "malzemedizini")
```

### Module1 — Açılan formlar

`DL1`, `UF2`, `UFKW` — malzeme seçim zinciri; `UF2` önceki liste kapanır (`macro_01`).

---

## Module2 (Standart modül — Ribbon callbacks II)

Teklif ribbon’unun ana işlevleri — pano girişi, formlar, bölüm başlıkları, para birimi, stok kontrolü.

| App | Section | Key | İşlem | Kullanıldığı yer | Açıklama |
|-----|---------|-----|-------|------------------|----------|
| `ilhan` | `Settings` | `pfc` | **Get** / **Save** | `Macropanofiyat` / `Macro311`, `Macro312` | Pano giriş modu: `0` = listeden, `1` = çarpandan |
| `ilhan` | `Settings` | `malzemedizini` | **Get** | `panogir0` | Montajlı pano `.xlsb` yolu |
| `ilhan` | `Settings` | `panodizini` | **Get** | `panogir0` | Seçili montajlı pano dosya adı (`UFOPAN02` ile kayıt) |
| `ilhan` | `Settings` | `dtxs` | **Get** / **Save** | `Macro16` | Malzeme değişim sütun sırası (ör. `C-E-F`) |
| `ilhan` | `Settings` | `deposabitdosya` | **Get** | `Stokkontrol1` | Stok dosya adı |
| `ilhan` | `Settings` | `depodizini` | **Get** | `Stokkontrol1` | Stok klasörü |

### Module2 — Sabit / dinamik yollar

| Yol | Kullanım |
|-----|----------|
| `{malzemedizini}\Malzeme Listeleri\4\{panodizini}` | `panogir0` → `UFOPAN00` (montajlı pano listesi) |
| `C:\Belgelerim\CEMEX\Kayıtlar\Kayıtlar.xlsb` | `Macro17` → `UserFormT4` |

### Module2 — Pano giriş zinciri

| Prosedür | Akış |
|----------|------|
| `Macropanofiyat` | `pfc` oku → `panogir0` veya `panogir1` |
| `Macro311` | `SaveSetting pfc=1` → `panogir1` → `UFOPAN00` (`pfc=1`) |
| `Macro312` | `SaveSetting pfc=0` → `UFOPAN02.Show` |
| `panogir0` | `panodizini` `.xlsb` aç (gizli) → `UFOPAN00` (`pfc=0`) |
| `panogir1` | `UFOPAN00.Show` (`pfc=1`) |

### Module2 — Ribbon makroları (gruplar)

| Grup | Makrolar | Form / çağrı |
|------|----------|----------------|
| Teklif / dosya | `Macro32`, `Macro5`, `Macro2` | `UFTH` (`ssno` 1/2), `dosyaac` |
| Pano | `Macropanofiyat`, `Macro311`, `Macro312`, `panoboyut` | `UFOPAN00`, `UFOPAN02`, `UFOPAN11` |
| Malzeme | `Macro13`, `Macro14`, `Macro16` | `UFmy`, `UFmd`, `UFMZ` |
| Döviz / analiz | `teklifkur`, `Macro23`–`25`, `Macro15`, `Macro500` | `UFKur`, `MakroTL/EURO/DOLAR`, `Sayfa3` gizle/göster |
| Biçim / başlık | `baslık1`–`3`, `toplam1`–`2`, `MacroDAD*` | `UserFormAD`, `UFDAD`, `UFDAD0`, `UFDAD1` |
| Şablon / güncelleme | `Macro28`, `Macro74`–`742` | `UserFormS1`, `UFDD`, `UFDD1`, `UFDD2` |
| PM satırları | `Macro43`–`45`, `Macro1621`, `Macro1641` | `iscilik_gir1`, `sarf_gir1`, `amb_gir1`, `pfkod` |
| Diğer | `Macro17`–`19`, `Macro190`–`191`, `Macro4`, `Macro18` | `UserFormT4`, `Hatabul`, `AraToplamlar`, `UFadT`, `Stokkontrol1` |
| Çıktı | `Macro501`–`503` | `PrintPDf`, `teklifsablonkaydet`, `UserFormS2` |

### Module2 — Yardımcılar ve değişkenler

| Öğe | Açıklama |
|-----|----------|
| `WorkbookOpen()` | Workbook açık mı kontrolü |
| `ilhan` | `Sayfa3!I55555` teklif doğrulama (Module1’den) |
| `pfc`, `dtx`, `tp1`, `tsi1`, `sa`, `ssno` | Form mod bayrakları |
| `tdosya` | Public (kullanım bağlamı harici) |

### Module2 — GetSetting / SaveSetting özeti

```vba
pfc = GetSetting("ilhan", "Settings", "pfc")
SaveSetting "ilhan", "Settings", "pfc", 0   ' veya 1

fm1 = GetSetting("ilhan", "Settings", "malzemedizini")
mlz = GetSetting("ilhan", "Settings", "panodizini")

dtxs = GetSetting("ilhan", "Settings", "dtxs")
SaveSetting "ilhan", "Settings", "dtxs", dtxs

d1m = GetSetting("ilhan", "Settings", "deposabitdosya")
yolm = GetSetting("ilhan", "Settings", "depodizini")
```

---

## Module3 (Standart modül — yardımcı prosedürler)

Teklif sayfası işlemleri, para birimi, toplamlar, PDF/şablon kayıt, klemens/kablo ve montaj fiyat güncelleme.

| App | Section | Key | İşlem | Kullanıldığı yer | Açıklama |
|-----|---------|-----|-------|------------------|----------|
| `ilhan` | `Settings` | `malzemedizini` | **Get** | `Macro171` | `Kablo ve Klemensler.xlsb` yolu |

**Registry Save:** Yok (bu modülde).

### Module3 — Sabit / dinamik yollar

| Yol | Kullanım |
|-----|----------|
| `{malzemedizini}\Otomatik Seçim\Kablo ve Klemensler.xlsb` | `Macro171` → `UFOSARF` (`sk1`) |
| `C:\Belgelerim\CEMEX\Parametreler\Montaj Fiyatları.xls` | `Macro75` — montaj ad/dk güncelleme |
| `C:\Belgelerim\Cemex\Yeni Teklif Şablonları\Yeni Teklif V1.2.xltx` | `sablonkaydet1` |
| `C:\Users\İlhan Şirin\OneDrive\Belgeler\Özel Office Şablonları\` | `sablonkaydet1` (geliştirici yolu) |
| `Veriler!B30` / `B31` | `pdfkaydet`, `teklifsablonkaydet` — PDF/şablon kayıt klasörü |

### Module3 — Ana prosedürler

| Prosedür | Açıklama | Çağıran |
|----------|----------|---------|
| `dosyaac` | `GetOpenFilename` ile teklif aç; `wd`, `dsa` | `UFTH` `Macro2` |
| `AraToplamlar` | Bölüm X toplamları + genel toplam | Module2 `Macro18`, formlar |
| `Hatabul` | Aynı sipariş kodunda fiyat/iskonto/formül farkı | Module2 `Macro4` |
| `MakroTL` / `MakroDOLAR` / `MakroEURO` | W:X format, `Tpbr`, `Tpb` | Module2, `UFTH.uygula_bilgiler` |
| `katekle` / `katsil` / `*hepsi` | Bölüm adet çarpanı (E sütunu `*E5`) | Module2 `Macro21`–`22` |
| `formulyap` / `yazırenk` | CML sayfa A sütunu / font | Module2 `Macro72`–`73` |
| `Macro75` | Montaj fiyatlarını harici xls’ten H sütununa yaz | Ribbon |
| `Macro170` | `UFObara1.Show` | Module2 |
| `Macro171` | `UFOSARF.Show` + klemens workbook | Module2 |
| `PrintPDf` / `pdfkaydet` | `Veriler` sayfası varsa PDF export | Module2 `Macro501` |
| `teklifsablonkaydet` / `sablonkaydet1` | Şablon `.xlsx` / `.xltx` kayıt | Module2 `Macro502` |
| `teklifformat` | `yeniteklif` | Ribbon |

### Module3 — Genel değişkenler

| Değişken | Açıklama |
|----------|----------|
| `wd` | `dosyaac` başarı bayrağı (`1` = dosya seçildi) |
| `sk1` | Klemens/kablo fiyat workbook adı (`Kablo ve Klemensler.xlsb`) |

### Module3 — Koşullar

- `Sayfa3!I55555` — `MakroTL`, kat işlemleri
- `Cells(1,"X") = "Toplam Fiyat"` — `AraToplamlar`
- `Veriler` veya `T1` CodeName — PDF/şablon kayıt

### Module3 — GetSetting özeti

```vba
fm1 = GetSetting("ilhan", "Settings", "malzemedizini")
' Macro171 → fm1 & "\Otomatik Seçim\Kablo ve Klemensler.xlsb"
```

---

## Module4 (Standart modül — UFOPAN11 görsel yardımcıları)

`UFOPAN11` pano yerleşim formundaki dinamik `Image` kontrollerini toplar ve örtü sacı renklendirir.

**Registry:** Yok (`GetSetting` / `SaveSetting` kullanılmıyor).

### Module4 — Sabit dosya yolu

| Yol | Kullanım |
|-----|----------|
| `C:\Belgelerim\Cemex\Şablonlar\` | `Alt_Klasör_İsimlerixxx` — alt klasör adlarını MsgBox (test/ölü kod) |

### Module4 — Değişkenler ve sınıf

| Öğe | Açıklama |
|-----|----------|
| `r()` | `EVN` sınıfı dizisi — her öğede `.Resimler` = `Image` kontrolü |
| `say` | `UFOPAN11` MultiPage1 sayfa 0’daki `Image` sayısı |

### Module4 — Prosedürler

| Prosedür | Açıklama |
|----------|----------|
| `isimler1` | `UFOPAN11.MultiPage1.Pages(0)` içindeki tüm `Image` kontrollerini `r(1..say)` dizisine yükler |
| `renkeo` | `EGOM1`…`EGOM{say}` ve `EGOD1`…`EGOD{say}` arka plan rengi `&H8000000F` |
| `Alt_Klasör_İsimlerixxx` | `Şablonlar` alt klasör listesi (debug; `xxx` soneki — muhtemelen kullanılmıyor) |

### Module4 — Üst bağlantı

- `UFOPAN11` — `bitti21` → `formayarlama` → `isimler1` zincirinde çağrılır
- `EVN` sınıfı — `Resimler` özelliği ile pano modül görsellerine referans

---

## Module5 (Standart modül — teklif format dönüşümü / işçilik)

Eski teklif formatını yeni şablona taşır, pano/PM kodlarını günceller, bölüm bazlı işçilik satırı ekler.

| App | Section | Key | İşlem | Kullanıldığı yer | Açıklama |
|-----|---------|-----|-------|------------------|----------|
| `ilhan` | `Settings` | `misi` | **Get** | `iscilik_gir2` | İşçilik markası (D sütunu) |
| `ilhan` | `Settings` | `misia` | **Get** | `iscilik_gir2` | İşçilik açıklaması (C sütunu) |

Kayıt: `UFOPAN00.CommandButton10`. **Workbook adı:** `CkarO` (kar formülü).

### Module5 — Sabit dosya yolu

| Yol | Kullanım |
|-----|----------|
| `C:\Belgelerim\CEMEX\Yeni Teklif Şablonları\Yeni Teklif V1.2.xltx` | `yeniteklif` — yeni şablon açılır, veri kopyalanır |

### Module5 — `yeniteklif`

1. Mevcut `Sayfa1` A:X → yeni şablona kopyala
2. Eski workbook adları (`Sayfa3` named range) yeni dosyaya aktar (kısa adlar)
3. `dt` = yeni workbook; `pfkod` çağrılır
4. Çağıran: `Module3.teklifformat` (ribbon `teklifformat`)

### Module5 — `pfkod` (eski → yeni referans kodları)

`PP-` satırlarında A sütunu metin değiştirme (C sütununda anahtar kelimeye göre):

| Eski önek | Koşul (A:C) | Yeni |
|-----------|-------------|------|
| `PP-PD` | FORM2/3/4 | `PP-F2/F3/F4`; yoksa `PP-DD` |
| `PP-PS` | HARİCİ / PASLANMAZ / ALTI | `PP-SH/SX/SA`; yoksa `PP-SD` |
| `PP-PH/PX/PK/PT` | — | `PP-DH/DX/KK/TK` |
| `PP-MM/MS/MB/MH/MN` | — | `PM-MP/MS/MB/MA/MN` |

Çağıran: `yeniteklif`, `Module2.Macro45`.

### Module5 — `iscilik_gir2`

Her `BÖLÜM ADI/NO:` … `BÖLÜM TOPLAMI:` aralığına `PM-MP-auto` satırı ekler (yoksa):

- C = `misia`, D = `misi`, E = 1, F = bölüm montaj toplamı (`Sum(Q…)`)
- Tam teklif formül seti (J–X, `CkarO`, `bfyt`, para birimi)
- `AraToplamlar` sonunda

### Module5 — GetSetting özeti

```vba
Range("C" & k) = GetSetting("ilhan", "Settings", "misia")
Range("D" & k) = GetSetting("ilhan", "Settings", "misi")
```

---

## Module6 (Standart modül — PM satırları: işçilik / sarf / ambalaj / bara)

Teklif `Sayfa1` üzerinde `PM-MP/MS/MA/MB` satırlarını pano verilerinden hesaplayarak ekler, günceller veya siler.

| App | Section | Key | İşlem | Kullanıldığı yer | Açıklama |
|-----|---------|-----|-------|------------------|----------|
| `ilhan` | `Settings` | `misi`, `misia` | **Get** | `verilergir` | İşçilik marka/açıklama |
| `ilhan` | `Settings` | `msas`, `msasa` | **Get** | `verilergir` | Sarf marka/açıklama |
| `ilhan` | `Settings` | `mama`, `mamaa` | **Get** | `verilergir` | Ambalaj marka/açıklama |
| `ilhan` | `Settings` | `mbab`, `mbaba` | **Get** | `verilergir` | Bara marka/açıklama |
| `ilhan` | `Settings` | `skdi`, `skds`, `skdb`, `skda` | **Get** | `verilergir` | B/C sipariş kodu önekleri |
| `ilhan` | `Settings` | `bara` | **Get** | `verilergir` (`PM-MB`) | Bara birim fiyatı |
| `ilhan` | `Settings` | `amb` | **Get** | `amb_gir1` | `Catv` yoksa varsayılan ambalaj çarpanı |

Kayıt: `UFOPAN00.CommandButton10` (PM alanları); `bara`/`amb` ayrıca `UFObara1` / `Catv` adı.

**Workbook adları:** `CkarO`, `Chtv`, `Cmtv`, `Cstv`, `Catv`.

### Module6 — Silme prosedürleri

| Prosedür | Hedef | Çağıran |
|----------|-------|---------|
| `isciliksarfambsil` | Tüm `PM-M*-auto` | `Module2.Macro44` |
| `isciliksil` | `PM-MP*` / `PM-MP-auto` | `UFOPAN00.CBmontajsil` |
| `sarfsil` | `PM-MS*` / `PM-MS-auto` | `UFOPAN00.CBsarfsil` |
| `ambsil` | `PM-MA*` / `PM-MA-auto` | `UFOPAN00.CBambsil` |

`UFOPAN00.CKPKOD` ve `tp1` — silinecek kod deseni (`PM-MP` vs `PM-MP-auto`).

### Module6 — Tek satır ekleme

| Prosedür | A sütunu | Çağıran |
|----------|----------|---------|
| `isciliksatır_gir1` | `PM-MP` | `Module2.Macro1621` |
| `sarfsatır_gir1` | `PM-MS` | `Module2.Macro1641` |
| `bakir_gir` | `PM-MB` | `Module2.Macro169` |

### Module6 — Bölüm bazlı otomatik giriş

| Prosedür | Satır | Hesap |
|----------|-------|-------|
| `iscilik_gir1` | `PM-MP-auto` | Pano `PP-*` + I sütunu harf → `Chtv`/`Cmtv` çarpanı × alan × adet |
| `sarf_gir1` | `PM-MS-auto` | Aynı mantık, `Cstv` çarpanları |
| `amb_gir1` | `PM-MA-auto` | Pano alanı × `Catv` (veya `amb` / InputBox çarpan) |

Ortak: `BÖLÜM ADI/NO:` … `BÖLÜM TOPLAMI:` döngüsü; kat çarpanı (`E` sütunu `# Adet`); `verilergir`; `AraToplamlar`.

`pnkod`: `PP-*-auto*` veya `CKPKOD` açıkken `PP-*` (`tp1=0` UFOPAN00 modu).

### Module6 — `verilergir` (ortak satır doldurma)

A sütununa göre B–D registry’den, F=`tfy`, L=kar formülü (`Oisci`/`Osarf`/`Oamb`/`Obara`), J–X teklif formülleri, biçimleme.

`PM-MB` + `CodeName` TM → `malzemedosya1` (malzeme listesi fiyat karşılaştırma sütunları).

### Module6 — Mod bayrakları

| Değişken | Kaynak | Etki |
|----------|--------|------|
| `tp1` | `Module2` (`Macro43`/`431`) | `1` = sessiz toplu giriş; progress bar kapalı |
| `tsi1` | `Macro431` | İşçilik+sarf birlikte mesajı |
| `tfy`, `s`, `adr` | Modül içi | Bölüm fiyatı ve kat adresi |

### Module6 — Üst bağlantılar

- `UFOPAN00` — `CBmontaj`, `CBsarf`, `CBamb`, silme butonları, `ProgressBarP21`
- `Module2` — ribbon PM makroları
- `Module3.AraToplamlar`, `Module5.iscilik_gir2` (farklı işçilik mantığı)

### Module6 — GetSetting özeti (`verilergir`)

```vba
' Örnek PM-MP-auto
Range("B" & k) = GetSetting("ilhan", "Settings", "skdi")
Range("C" & k) = GetSetting("ilhan", "Settings", "misia")
Range("D" & k) = GetSetting("ilhan", "Settings", "misi")
' PM-MB → GetSetting("bara") for F column
' amb_gir1 → GetSetting("amb") when Catv name empty
```

---

## Module7 (Standart modül — UFTH TreeView bağlam menüsü ve dosya adı)

`UFTH` klasör ağacı için geçici CommandBar (`flexgrid_rc`) ve dosya/klasör CRUD; `UserFormS1` şablon çıktısında dosya yeniden adlandırma.

**Registry:** Bu modülde `GetSetting` / `SaveSetting` yok.

### Module7 — Sabit yollar

| Yol | Kullanım |
|-----|----------|
| `C:\Belgelerim\Cemex\Yeni Teklif Şablonları\Yeni Teklif V1.2.xltx` | `menu2` — TreeView’den yeni teklif oluşturma |

*(Aynı şablon: **Module5** `yeniteklif`, **UserFormS1** `sablonkaydet1`; büyük/küçük harf `Cemex` / `CEMEX` farklı yazımlar.)*

### Module7 — CommandBar

| Prosedür | İşlev |
|----------|--------|
| `CreateCmdBar` | `flexgrid_rc` popup oluşturur (`Temporary:=False`); 4 düğme |
| `DestroyCmdBar` | Popup siler |

| Düğme | Caption | `OnAction` |
|-------|---------|------------|
| 1 | Yeni Dizin | `menu1` |
| 2 | Yeni Teklif | `menu2` |
| 3 | Yeniden Adlandır | `menu3` |
| 4 | Sil | `menu4` |

### Module7 — TreeView işlemleri (`UFTH`)

| Prosedür | İşlev |
|----------|--------|
| `Menu1` | Seçili düğüm altında klasör oluşturur (`Scripting.FileSystemObject.CreateFolder`); TreeView’e `Image:=2` ekler; `Labelyol`, `Lb` günceller |
| `menu2` | Şablonu açar → `Sayfa3!C8` = teklif tarihi → `SaveAs` `{Labelyol}\{ad}.xlsx` → TreeView `Image:=3` → workbook kapatılır; `TBds` güncellenir |
| `menu3` | Dosya veya klasör yeniden adlandırır; açık dosyada `ChangeFileAccess` + `movefile`; klasörde alt düğüm `key` güncellemesi |
| `menu4` | Onay sonrası `DeleteFile` veya `DeleteFolder`; TreeView düğümü kaldırır; Toolbar buton 5 devre dışı |
| `menu4xx` | Eski silme mantığı (kullanılmıyor / hatalı `DeleteFolder`+`DeleteFile` birlikte) |

**Ortak:** `UFTH.Labelyol`, `UFTH.Lb`, `UFTH.TreeView1`, `UFTH.TBds`, `UFTH.Toolbar1`; `WorkbookOpen()` (başka modülde tanımlı).

`menu2` / `menu3`: Seçim dosya ise (`TBds` / `kA` `.xls*`) üst klasör düğümüne iner.

### Module7 — `DOSYAADI_R1`

`UserFormS1` `Label253` tetikler. `TBAA4` doluysa aktif workbook’u yeniden adlandırır:

1. `Veriler!B4` = yeni ad (uzantısız)
2. `movefile` → `{path}\{TBAA4}.xlsx`
3. `Veriler!B30` = yol, `B31` = dosya adı (uzantısız)
4. `UserFormS2` kapat/aç

### Module7 — Üst bağlantılar

- **UFTH** — TreeView sağ tık menüsü (`CreateCmdBar` / `DestroyCmdBar` form yaşam döngüsünde)
- **UserFormS1** — `Label253` → `DOSYAADI_R1`
- **Module5** — aynı teklif şablonu yolu

---

## Module9 (Standart modül — ListBox tekerlek kaydırma + UF2 pano metin kopyala/yapıştır)

UserForm `ListBox` / benzeri kontrollerde fare tekerleği ile kaydırma (Win32 düşük seviye mouse hook) ve **UF2** pano açıklama alanı pano kopyala/yapıştır yardımcıları.

**Registry:** `GetSetting` / `SaveSetting` yok.  
**Sabit dosya yolu:** yok.

### Module9 — Win32 API (mouse hook)

| Öğe | Açıklama |
|-----|----------|
| `POINTAPI`, `MOUSEHOOKSTRUCT` | Hook callback yapıları |
| `SetWindowsHookEx` / `UnhookWindowsHookEx` | `WH_MOUSE_LL` (14) düşük seviye mouse hook |
| `WindowFromPoint`, `GetCursorPos` | İmleç altındaki HWND |
| `GetWindowLong` (`GWL_HINSTANCE`) | Hook modül örneği |

`FindWindow` tanımlı ancak bu modülde **kullanılmıyor**.

### Module9 — ListBox kaydırma

| Prosedür | İşlev |
|----------|--------|
| `HookListBoxScroll(frm, ctl)` | `ctl` odaklanır; imleç altı HWND `mListBoxHwnd` olur; `MouseProc` hook kurulur |
| `UnhookListBoxScroll` | Hook kaldırılır; `mCtl` temizlenir |
| `MouseProc` (Private) | `WM_MOUSEWHEEL` (`&H20A`) — `mCtl.TopIndex` ±1; imleç başka penceredeyse unhook |

**Kullanım:** UserForm `ListBox_MouseMove` / odak olaylarında `HookListBoxScroll Me, ListBox1` (çağıran formlar kodda dağıtık; VBA ListBox varsayılan olarak tekerleği desteklemez).

Modül düzeyi durum: `mLngMouseHook`, `mListBoxHwnd`, `mbHook`, `mCtl`.

### Module9 — UF2 pano metin

| Prosedür | İşlev |
|----------|--------|
| `Textbox_Copy2` (Private) | `UF2.TextBoxPPT` tamamını seçip panoya kopyalar (boşsa çık) |
| `Textbox_Paste` | `UF2.TextBox22.Paste` → `UF2.TBox22` (yapıştırma sonrası güncelleme) |

`TextBoxPPT` / `TextBox22` — UF2 pano/ürün açıklama alanları; ribbon veya form butonlarından `Textbox_Paste` çağrılabilir.

### Module9 — Üst bağlantılar

- **UF2** — `TextBoxPPT`, `TextBox22`, `TBox22`
- UserForm’lar — `HookListBoxScroll` / `UnhookListBoxScroll` (QueryClose veya `Deactivate` ile unhook önerilir)

---

## zActiveWb (Standart modül — aktif workbook yolu)

Aktif Excel çalışma kitabının tam dosya yolunu registry’ye yazar. **`ilhan` / `Settings` değil** — ayrı uygulama adı `sercan`.

| App | Section | Key | İşlem | Kullanıldığı yer | Açıklama |
|-----|---------|-----|-------|------------------|----------|
| `sercan` | `fileOpenWorkBooks` | `nowOpenPropsFile` | **Save** | `GetActiveWorkbookPath` | Aktif kitabın `FullName` (kaydedilmiş dosya tam yolu) |

**Get:** Bu modülde yok; tüketici **zInternet** `PostDataToServer` ve diğer modüller.

### zActiveWb — `GetActiveWorkbookPath`

1. `Application.ActiveWorkbook` — yoksa MsgBox, çık
2. `activeWB.Path = ""` — henüz kaydedilmemişse MsgBox, çık
3. `wbPath = activeWB.FullName` → `SaveSetting`

Sabit dosya yolu yok.

### zActiveWb — SaveSetting özeti

```vba
SaveSetting "sercan", "fileOpenWorkBooks", "nowOpenPropsFile", wbPath
```

### zActiveWb — Notlar

- Muhtemel kullanım: uzaktan modül / özellik paneli için “şu an hangi teklif dosyası açık” bilgisi
- **zInternet** `PostDataToServer` — `GetSetting` ile okur
- `ilhan` anahtarlarından bağımsız HKCU kaydı (`HKEY_CURRENT_USER\Software\VB and VBA Program Settings\sercan\...`)

---

## zDosyaİslemleri (Standart modül — iş programından teklif klasörü)

**İş Programı** Excel dosyasından masaüstünde proje klasör yapısı oluşturur; teklif şablonunu kopyalayıp `Sayfa3` meta alanlarını doldurur.

**Registry:** `GetSetting` / `SaveSetting` yok.

### zDosyaİslemleri — Sabit yollar

| Yol | Kullanım |
|-----|----------|
| `C:\Belgelerim\Cemex\Yeni Teklif Şablonları\sablon.xlsm` | `SABLON_YOLU` — `FileCopy` kaynağı (`teklif_klasor_olustur`) |
| `{Masaüstü}\{F sütunu klasör adı}\` | Ana proje klasörü (`WScript.Shell` `SpecialFolders("Desktop")`) |

Yorum satırında alternatif: `sablon.xlsx`. Ana teklif akışındaki `Yeni Teklif V1.2.xltx` / `.xlsx` şablonundan **farklı** dosya.

### zDosyaİslemleri — Ön koşul

Aktif workbook adı `"TEKLİF VE İŞ PROGRAMI"` içermeli (`GEREKLI_DOSYA_ADI_KISMI`); aksi halde makro çalışmaz.

### zDosyaİslemleri — Kaynak sayfa (aktif sheet)

| Sütun | Alan | Kullanım |
|-------|------|----------|
| A | Teklif ID | `EPR-ddmmyy-0001` formatında otomatik sıra; son dolu satırın altına yazılır |
| C | Firma | `Sayfa3!C5` (işveren) |
| D | Teklif ilgilisi | `Sayfa3!C11` |
| E | Proje adı | `Sayfa3!C3` |
| F | Klasör adı | Masaüstü ana klasör + teklif dosya adı (`{ad}.xlsm`) |

ID üretimi: önceki satır `...-000N` ise `N+1`; format hatalıysa sıra 1’den başlar.

### zDosyaİslemleri — Oluşturulan klasör yapısı

```
{Masaüstü}\{F}\
  1-İş Emri
  2-Projeler
  3-Malzeme Listesi
  4-Teklif          ← sablon.xlsm → {F}.xlsm
  5-Resimler
  6-Test Raporu
  7-Analiz
```

Ana klasör zaten varsa işlem durur.

### zDosyaİslemleri — Teklif dosyası doldurma

`Workbooks.Open(hedefTamYol)` → `Sayfa3`:

| Hücre | Değer |
|-------|--------|
| C3 | `projeAdi` (E sütunu) |
| C4 | `"PROJE KISA ADI"` (sabit) |
| C5 | `firmaAdi` (C sütunu) |
| C7 | `yeniID` |
| C8 | `Date` |
| C10 | `"SERCAN GÜNGÖR"` (sabit hazırlayan) |
| C11 | `teklifIlgilisi` (D sütunu) |

Dosya açık bırakılır (`Close` yorum satırında).

### zDosyaİslemleri — Üst bağlantılar

- **Module1** — `teklifDosyaOlustur` → `teklif_klasor_olustur`
- **UFTH** / **Module7** / **Module5** — farklı şablon yolları (`V1.2.xltx`); bu modül `sablon.xlsm` kullanır

---

## zInternet (Standart modül — lisans API, uzaktan kod, eklenti güncelleme)

Next.js lisans/sunucu API’si ile iletişim: MAC tabanlı lisans sorgusu, teklif verisi POST, uzaktan VBA modül çalıştırma, `.xlam` indirme/güncelleme.

| App | Section | Key | İşlem | Kullanıldığı yer | Açıklama |
|-----|---------|-----|-------|------------------|----------|
| `sercan` | `fileOpenWorkBooks` | `nowOpenPropsFile` | **Get** | `PostDataToServer` | Aktif teklif tam yolu (okunur; gönderimde `ActiveWorkbook` kullanılır) |

**Save:** Bu modülde yok (`**zActiveWb**` yazar).  
**`ilhan` / `Settings`:** yok.

### zInternet — API sabitleri ve uç noktalar

| Sabit / URL | VBA prosedür | HTTP | Sunucu route (proje) |
|-------------|--------------|------|----------------------|
| `GET_LICENSE_URL` = `http://localhost:3000/api/` | `GetLicenseStatus` | GET | `/api/license/{mac}` |
| `POST_LICENSE_URL` = `http://localhost:3000/api/license/` | `RegisterLicense`, `PostDataToServer` | POST | `/api/license` |
| `GET_LICENSE_URL & "/module"` | `RunRemoteCode` | POST | `/api/module` |
| `GET_LICENSE_URL & "download/teklif"` | `LisansKontrolVeGuncelleme` | POST | `/api/download/teklif` |

`RegisterLicense` sabit `FIRMA_ADI = "ABC"` gönderir; `dosyaAdi` = `ThisWorkbook.Name`.

### zInternet — Lisans akışı

| Prosedür | İşlev |
|----------|--------|
| `IsInternetConnected` | `wininet.dll` `InternetGetConnectedState` |
| `GetFirstMACAddress` / `GetFirstIPAddress` | WMI `Win32_NetworkAdapterConfiguration` (`IPEnabled=True`) |
| `GetLicenseStatus` | GET → JSON yanıt metni |
| `RegisterLicense` | POST yeni kayıt; `201` başarı |
| `TestLicenseCheck` | GET parse (`JsonConverter.ParseJson`); `success=false` → POST; `data.license` → **`zLicense.SaveLicenseToRegistry`** |
| `TestInternetConnection` | İnternet varsa `TestLicenseCheck` → **`zLicense.GetLicenseFromRegistry`** → `Module1.ShowCustomTab` / `HideCustomTab` |

Bağımlılık: **`JsonConverter`** modülü. Lisans registry: **`zLicense`** (`scngnr` / `license`).

**Not:** `If zLicense.GetLicenseFromRegistry Then` — fonksiyon `String` döndürür; VBA’da yalnızca boş `""` False sayılır. Sunucu `license: "false"` döndürürse sekme yine görünür olabilir.

### zInternet — Teklif verisi gönderimi

**`PostDataToServer`** — `ActiveWorkbook` `Sayfa3` + `Sayfa1` satırlarından JSON:

| Alan | Kaynak |
|------|--------|
| `macAdresi` | `GetFirstMACAddress` |
| `dosyaAdi`, `firmaAdi`, `projeAdi`, … | `wb.Name`, `Sayfa3` C3–C5, D16–D17, I22, I25, M31 |
| `veritabaniTeklif` | `CreateVeritabaniJson(Sayfa1)` — B–G sütunları |

Yardımcılar: `CleanJson`, `CreateVeritabaniJson`.

### zInternet — Uzaktan kod çalıştırma

| Prosedür | İşlev |
|----------|--------|
| `RunRemoteCode(methodName)` | POST `/api/module` → `ExtractCodeFromJSON` → `ExecuteDynamicFunction` |
| `ExecuteDynamicFunction` | Gizli geçici workbook + `TempMod`; `Application.Run` → `DynamicFunc(targetWb, param)` |
| `ExtractCodeFromJSON` | Yanıttan `"code"` alanını ayıklar |

**Module1** `ribbonLoaded` → `RunRemoteCode("AutoStartOnExcelOpen")`.

VBA proje erişimi gerekir: *Güven Merkezi → VBA projesi nesne modeline erişime güven*.

`RunRemoteCode` hedefi: `ThisWorkbook` (eklenti). `zInternet-additions.bas` sürümünde `GetHostWorkbook(ActiveWorkbook)` kullanılır.

### zInternet — Eklenti güncelleme

**`LisansKontrolVeGuncelleme`:**

1. POST `download/teklif` + `{ macAdresi }`
2. Binary yanıt → `%TEMP%\kitap23.xlam`
3. **`SafeUpdateAddin`** → `Application.UserLibraryPath & "kitap23.xlam"`; AddIn adı `teklif`

| Sabit | Değer |
|-------|--------|
| `YEREL_DOSYA_ADI` | `kitap23.xlam` |
| `EKLENTI_ISMI` | `teklif` (uzantısız) |

### zInternet — Sabit / dinamik yollar

| Yol | Kullanım |
|-----|----------|
| `http://localhost:3000/api/…` | Tüm API çağrıları (ortam değişkeni değil, kod sabiti) |
| `Application.UserLibraryPath` | Eklenti hedef klasörü |
| `Environ("TEMP")` | İndirilen `.xlam` geçici dosyası |

### zInternet — GetSetting özeti

```vba
registryFilePath = GetSetting("sercan", "fileOpenWorkBooks", "nowOpenPropsFile", "")
' PostDataToServer içinde okunur; asıl veri ActiveWorkbook'tan alınır
```

### zInternet — Üst bağlantılar

- **Module1** — ribbon yükleme, sekme görünürlüğü, `RunRemoteCode`
- **zActiveWb** — `nowOpenPropsFile` yazma
- **zLicense** — registry’de lisans bayrağı
- **Next.js API** — `src/app/api/license`, `module`, `download/teklif`

---

## zLicense (Standart modül — lisans registry CRUD + VBA koruma)

Sunucudan gelen lisans değerini Windows registry’ye yazar/okur; şifresiz VBA projelerinde kod temizleme koruması.

| App | Section | Key | İşlem | Kullanıldığı yer | Açıklama |
|-----|---------|-----|-------|------------------|----------|
| `scngnr` | `Settings` | `license` | **Get** / **Save** / **Delete** | CRUD fonksiyonları | Sunucu `data.license` değeri (`"true"` / `"false"` metin) |
| `ilhan` | `Settings` | `license` | **Save** | `deleteFirstLicenseSystem` | `Setting:=False` — eski/paralel namespace |

Sabitler: `APP_NAME=scngnr`, `SECTION_NAME=Settings`, `LICENSE_KEY_NAME=license`, `DEFAULT_LICENSE_VALUE=""`.

Sabit dosya yolu yok.

### zLicense — CRUD

| Prosedür | İşlem |
|----------|--------|
| `SaveLicenseToRegistry(licenseValue)` | `SaveSetting` — oluştur/güncelle |
| `GetLicenseFromRegistry()` | `GetSetting`; yoksa `""` |
| `DeleteLicenseFromRegistry()` | `DeleteSetting` |
| `DoesLicenseExistInRegistry()` | `GetLicenseFromRegistry <> ""` |

### zLicense — Test subs

| Sub | Açıklama |
|-----|----------|
| `Test_Registry_CRUD_Flow` | Immediate penceresinde okuma testi (çoğu adım yorumlu) |
| `TestGetLicenseFromRegistry` | MsgBox ile mevcut değer |

### zLicense — VBA koruma / temizlik

| Prosedür | İşlev |
|----------|--------|
| `checkVbaPasswordProtect` | `ThisWorkbook.VBProject.Protection = 1` ise korunmuş; değilse `deleteAllVbaAndFormCode` |
| `deleteAllVbaAndFormCode` | Şifresiz projede tüm modül/class/form sil; document modüllerinde satırları temizle |
| `deleteFirstLicenseSystem` | `ilhan` / `Settings` / `license` = `False` yazar |

VBA proje erişimi gerekir (*VBA projesi nesne modeline erişime güven*). `deleteAllVbaAndFormCode` `ThisWorkbook` üzerinde çalışır; `ThisWorkbook` kodunu silerken makro durur.

### zLicense — GetSetting / SaveSetting özeti

```vba
' Ana lisans (scngnr)
SaveSetting "scngnr", "Settings", "license", licenseValue
GetSetting("scngnr", "Settings", "license", "")
DeleteSetting "scngnr", "Settings", "license"

' Eski köprü (ilhan)
SaveSetting "ilhan", "Settings", "license", False
```

### zLicense — Üst bağlantılar

- **zInternet** — `TestLicenseCheck` → `SaveLicenseToRegistry`; `TestInternetConnection` → `GetLicenseFromRegistry`
- **Module1** — `ShowCustomTab` / `HideCustomTab` (lisans kontrolü sonrası)
- **Next.js API** — `data.license` alanı (`licenses.json`)

---

## Registry canlı export (`data/registry-settings-export.json`)

**ExportRegistrySettings** modülü ile `teklif.xlam` üzerinden alınmış örnek (2026-06-21). Okunabilir kopya: `data/registry-settings-export.pretty.json`.

### Özet — dolu / boş anahtarlar

| Durum | `ilhan` anahtarlar |
|-------|-------------------|
| **Dolu** | `malzemedizini`, `pfc`, `teklifdizini`, `tbnoek`, `TBveren`, `bara`, `baralar`, `amb`, `panodizini`, `panomarka`, PM marka/açıklama/kodlar (`misi`…`skda`), tüm UFOPAN00 çarpanları (`PTA*`, `*cp`, `cpi*`, `cps*`, `drcp`, `drmi`, `mdip`, `Cemex`) |
| **Boş** | `deposabitdosya`, `depodizini`, `dtxs`, `sonteklif` |
| **Dikkat** | `tbnoek` çok satırlı (`\n` ile tekrarlı); `ilhan.license`=`False` ile `scngnr.license`=`false` farklı namespace |

### `sercan` / `scngnr` (aynı export)

| App | Key | Değer (kısaltılmış) |
|-----|-----|---------------------|
| `sercan` | `nowOpenPropsFile` | `...\4-Teklif\GELEN\ARKAS İLAVE PANOLAR TEKLİF.xlsx` |
| `scngnr` | `license` | `false` |

### Tüm `ilhan` key'leri (güncel — 71 adet: 27 temel + 44 UFOPAN00 çarpan)

| Key | Form(lar) | Get | Save |
|-----|-----------|-----|------|
| `malzemedizini` | DL1, UFDD, UFKW, UFmy, UFMZ, UFOPAN02, Module1, Module2, Module3 | ✓ | DL1 |
| `deposabitdosya` | UF2, **Module2** | ✓ | — |
| `depodizini` | UF2, **Module2** | ✓ | — |
| `dtxs` | UF2, **Module2** | ✓ | **Module2** |
| `pfc` | **Module2** | ✓ | ✓ |
| `teklifdizini` | UFTH | ✓ | ✓ |
| `tbnoek` | UFTH | ✓ | ✓ |
| `sonteklif` | UFTH | ✓ | ✓ |
| `TBveren` | UFTH | ✓ | ✓ |
| `bara` | UFObara1, **Module6** | ✓ | ✓ |
| `baralar` | UFObara1 | *(ölü)* | ✓ |
| `amb` | **Module6** | ✓ | *(Catv / UFObara1?)* |
| `panodizini` | UFOPAN02, Module2 | ✓ | ✓ |
| `panomarka` | UFOPAN00 | ✓ | ✓ |
| `misi` | UFOPAN00, Module5, **Module6** | ✓ | ✓ |
| `msas` | UFOPAN00, **Module6** | ✓ | ✓ |
| `mbab` | UFOPAN00, **Module6** | ✓ | ✓ |
| `mama` | UFOPAN00, **Module6** | ✓ | ✓ |
| `misia` | UFOPAN00, Module5, **Module6** | ✓ | ✓ |
| `msasa` | UFOPAN00, **Module6** | ✓ | ✓ |
| `mbaba` | UFOPAN00, **Module6** | ✓ | ✓ |
| `mamaa` | UFOPAN00, **Module6** | ✓ | ✓ |
| `skdi` | UFOPAN00, **Module6** | ✓ | ✓ |
| `skds` | UFOPAN00, **Module6** | ✓ | ✓ |
| `skdb` | UFObara1, UFOPAN00, **Module6** | ✓ | ✓ |
| `skda` | UFOPAN00, **Module6** | ✓ | ✓ |
| `license` | **zLicense** (`deleteFirstLicenseSystem`) | — | ✓ |
| `Cemex`, `mdip`, `drcp`, `drmi` | **UFOPAN00** (export) | ✓ | ✓ |
| `PTA1`…`PTA12` | **UFOPAN00** pano tip kısaltmaları | ✓ | ✓ |
| `ddcp`, `dhcp`, `dxcp`, `sdcp`, `shcp`, `sxcp`, `sacp`, `kkcp`, `tkcp`, `ddf2`…`ddf4` | **UFOPAN00** pano tip çarpanları | ✓ | ✓ |
| `cpia`…`cpih`, `cpsa`…`cpsh` | **UFOPAN00** montaj harf katsayıları | ✓ | ✓ |

---

### Tüm `scngnr` key'leri (güncel — 1 adet)

| App | Section | Key | Modül | Get | Save | Delete | Açıklama |
|-----|---------|-----|-------|-----|------|--------|----------|
| `scngnr` | `Settings` | `license` | **zLicense** | ✓ | ✓ | ✓ | Sunucu lisans metni (`"true"` / `"false"`) |

---

### Tüm `sercan` key'leri (güncel — 1 adet)

| App | Section | Key | Modül | Get | Save | Açıklama |
|-----|---------|-----|-------|-----|------|----------|
| `sercan` | `fileOpenWorkBooks` | `nowOpenPropsFile` | **zActiveWb**, **zInternet** | ✓ | ✓ | Aktif workbook `FullName` |

