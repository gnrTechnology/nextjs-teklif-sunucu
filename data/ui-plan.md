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

---

## Tamamlanan (bu tur)

- **Global MAC filtresi** — topbar altı şerit; heartbeats/komutlar/cihazlar ile paylaşılır
- **LiveRefresh** — operasyon sayfalarında SSE (`/api/events`) ile `router.refresh()`
- **Cihazlar** — grid/liste görünümü, detay drawer, EmptyState, PageHeader canlı
- **Modüller** — feasibility sekmeleri (Tam / Kısmi / Altyapı), "daha fazla yükle"
- **Komutlar** — MAC'e göre gruplama, stuck komut vurgusu
- **loading.tsx** — dashboard, cihazlar, komutlar, modüller

---

## Yeni öneriler (gelecek)

### P1 — Operasyon
| # | Öneri | Etki |
|---|--------|------|
| N1 | Cihaz snapshot **geçmişi** (DB: `device_snapshot_history`) + diff UI | Donanım değişikliği takibi |
| N2 | Dashboard'da **harita / coğrafi IP** (public IP → ülke) | Çok şubeli firmalar |
| N3 | **Webhook** bildirimi: offline > 1 saat, komut error | Proaktif uyarı |
| N4 | Komut şablonları (sık kullanılan modül+param kaydet) | Operatör hızı |

### P2 — Veri & performans
| # | Öneri | Etki |
|---|--------|------|
| N5 | Modül listesinde **tam virtual scroll** (sabit satır yüksekliği modu) | 700+ modül |
| N6 | **Server-side** modül arama/filtre (API pagination) | Büyük DB |
| N7 | Analitik: zaman serisi grafikleri (recharts / chart.js) | Trend analizi |

### P3 — Kurumsal
| # | Öneri | Etki |
|---|--------|------|
| N8 | Çoklu admin rolü (salt okunur / operatör / geliştirici) | Güvenlik |
| N9 | Firma bazlı **tenant** görünümü (yalnızca kendi MAC'leri) | SaaS hazırlığı |
| N10 | PDF/Excel export — lisans, cihaz envanteri raporu | Raporlama |

### P4 — Excel ↔ Web
| # | Öneri | Etki |
|---|--------|------|
| N11 | Ribbon "Cihaz Bilgisi Gönder" → `CollectDeviceInfoServer` | xlam-inventory B1 |
| N12 | Dashboard'dan **tek tık komut** (MAC seçili → modül gönder) | Operasyon |
| N13 | Modül çıktısı önizleme (Excel yerine web tablo) | Hızlı kontrol |

---

## Sayfa haritası

```
Operasyon          Yapılandırma        Geliştirici
/  Dashboard       /lisanslar          /moduller
/heartbeats        /firma-modulleri    /oneriler
/komutlar          /klasor-izleme      /api-referans
/modul-ciktilari                       /kurulum
/cihazlar
/loglar
/analitik
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
