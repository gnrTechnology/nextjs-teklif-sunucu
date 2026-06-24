# Teklif Sunucu — UI Geliştirme Planı

> Son güncelleme: 2025-06 · Neon DB tek kaynak

## Durum özeti

| Faz | Konu | Durum |
|-----|------|-------|
| 1 | Tasarım sistemi (Badge, StatCard, DataTable, …) | ✅ |
| 2 | Nav grupları, topbar, mobil drawer | ✅ |
| 3 | Dashboard alarm KPI + önizleme tabloları | ✅ |
| 4 | Tablolar, modül filtreleri, cihaz kartları | ✅ |
| 5 | Toast, tema, kurulum rehberi, loading | ✅ |
| 6 | Admin auth, error sayfası | ✅ |
| 7 | Alt menü grupları + N1–N13 temel altyapı | ✅ |

---

## Tamamlanan (son tur)

### Operasyon
- **Ekran yakalama** — `CaptureScreenshot` modülü JPEG base64 olarak `/api/module-output/` POST eder
- **Modül Çıktıları → Ekran Görüntüleri** sekmesi — galeri, tam ekran zoom, MAC filtresi
- **Cihaz snapshot geçmişi (N1)** — `device_snapshot_history` tablosu; drawer'da diff UI
- **Coğrafi IP (N2)** — Dashboard'da `publicIp` → ülke/şehir özeti (`ip-api.com`)
- **Webhook (N3)** — `WEBHOOK_URL` env; komut `error` durumunda POST
- **Komut şablonları (N4)** — `command_templates` + Komutlar sayfası chip'leri
- **Hızlı komut (N12)** — Dashboard'da MAC + modül seç → kuyruğa ekle

### Veri & performans
- **Server-side modül arama (N6)** — `GET /api/modules/?search=&category=&offset=&limit=`
- **Analitik grafikleri (N7)** — yatay bar grafikleri (modül çalışma, kategori, cihaz durumu)

### Navigasyon
- **Alt gruplar** — Operasyon: Genel Bakış / Canlı İzleme / Cihaz & Komut; Yapılandırma ve Geliştirici alt bölümleri; daraltılabilir sidebar (localStorage)

### Modül referansı
| Amaç | Modül |
|------|-------|
| Açık **pencereler** (başlık listesi) | `EnumVisibleWindows` (DLL D20) |
| Çalışan **process** listesi (CPU/RAM) | `GetRunningProcesses` |
| Aktif pencere başlığı | `GetForegroundWindowTitle` |
| Ekran görüntüsü | `CaptureScreenshot` |

---

## Öneriler — uygulama durumu

### P1 — Operasyon
| # | Öneri | Durum |
|---|--------|-------|
| N1 | Cihaz snapshot geçmişi + diff UI | ✅ |
| N2 | Dashboard harita / coğrafi IP | ✅ (badge özeti; tam harita: gelecek) |
| N3 | Webhook: offline > 1 saat, komut error | ⚠️ Komut error ✅; offline cron: gelecek |
| N4 | Komut şablonları | ✅ |

### P2 — Veri & performans
| # | Öneri | Durum |
|---|--------|-------|
| N5 | Modül listesi tam virtual scroll | ⚠️ Load-more mevcut; sabit satır modu: gelecek |
| N6 | Server-side modül arama/filtre | ✅ API |
| N7 | Analitik zaman serisi grafikleri | ✅ CSS bar; recharts: opsiyonel |

### P3 — Kurumsal
| # | Öneri | Durum |
|---|--------|-------|
| N8 | Çoklu admin rolü | 📋 `ADMIN_ROLE` env taslağı — middleware genişletme bekliyor |
| N9 | Firma bazlı tenant görünümü | 📋 Global MAC filtresi var; firma tenant: gelecek |
| N10 | PDF/Excel export | 📋 JSON export cihazlarda ✅; toplu rapor: gelecek |

### P4 — Excel ↔ Web
| # | Öneri | Durum |
|---|--------|-------|
| N11 | Ribbon "Cihaz Bilgisi Gönder" | 📋 `xlam-inventory.md` B1 — xlam merge gerekli |
| N12 | Dashboard tek tık komut | ✅ |
| N13 | Modül çıktısı web önizleme | ✅ (screenshot galeri + JSON viewer) |

---

## Sayfa haritası (alt gruplar)

```
Operasyon
  Genel Bakış     → /  Dashboard, /analitik
  Canlı İzleme    → /heartbeats, /loglar
  Cihaz & Komut   → /cihazlar, /komutlar, /modul-ciktilari

Yapılandırma
  Lisans & Erişim → /lisanslar
  Otomasyon       → /firma-modulleri, /klasor-izleme

Geliştirici
  Modül Yönetimi  → /moduller, /oneriler
  Kaynaklar       → /api-referans, /kurulum
```

---

## Bileşen kataloğu

| Bileşen | Yol |
|---------|-----|
| PageHeader + Refresher | `components/ui/PageHeader.tsx` |
| EmptyState | `components/ui/EmptyState.tsx` |
| DataTable | `components/ui/DataTable.tsx` |
| DetailDrawer | `components/ui/DetailDrawer.tsx` |
| VirtualList | `components/ui/VirtualList.tsx` |
| MacFilter | `lib/mac-filter.tsx` |
| LiveRefresh | `components/LiveRefresh.tsx` |
| Webhook | `lib/webhook.ts` (`WEBHOOK_URL`) |
| GeoIP | `lib/geoip.ts` |
| Snapshot geçmişi API | `/api/device-info/history/` |
| Komut şablonları API | `/api/command-templates/` |

---

## Ortam değişkenleri

| Değişken | Açıklama |
|----------|----------|
| `DATABASE_URL` | Neon PostgreSQL |
| `ADMIN_PASSWORD` | Admin giriş |
| `WEBHOOK_URL` | Opsiyonel; komut hata bildirimi (Slack/Discord hook) |

---

## Ekran görüntüsü akışı

1. Komutlar veya Dashboard → `CaptureScreenshot` gönder
2. Excel modülü birincil ekranı yakalar, JPEG base64 üretir, `POST /api/module-output/`
3. Web → **Modül Çıktıları → Ekran Görüntüleri** sekmesi
