export const dynamic = "force-dynamic";

import { listFirmAutoModules } from "@/lib/firm-auto-modules";
import { listLicenses } from "@/lib/db";
import { listModules } from "@/lib/modules";
import styles from "./page.module.css";

export default function Home() {
  const licenses = listLicenses();
  const modules = listModules();
  const firmAutoModules = listFirmAutoModules();

  return (
    <main className={styles.main}>
      <h1>Teklif Sunucu</h1>
      <p className={styles.lead}>
        Excel VBA istemcisi ile uyumlu lisans ve veri sunucusu.
      </p>

      <section className={styles.section}>
        <h2>API Uç Noktaları</h2>
        <ul>
          <li>
            <span className={styles.code}>GET /api/license/&#123;mac&#125;</span> — Lisans sorgula
          </li>
          <li>
            <span className={styles.code}>POST /api/license/</span> — Yeni kayıt / teklif verisi gönder
          </li>
          <li>
            <span className={styles.code}>POST /api/module</span> — Uzak VBA modülü al
          </li>
          <li>
            <span className={styles.code}>GET /api/auto-start/&#123;mac&#125;</span> — Firma bazlı otomatik modüller
          </li>
          <li>
            <span className={styles.code}>POST /api/download/teklif</span> — Eklenti indir (.xlam)
          </li>
        </ul>
      </section>

      <section className={styles.section}>
        <h2>Kayıtlı Lisanslar ({licenses.length})</h2>
        {licenses.length === 0 ? (
          <p>Henüz kayıt yok. Excel ilk bağlandığında otomatik oluşturulur.</p>
        ) : (
          <table className={styles.table}>
            <thead>
              <tr>
                <th>MAC</th>
                <th>Firma</th>
                <th>Lisans</th>
                <th>Dosya</th>
                <th>Güncelleme</th>
              </tr>
            </thead>
            <tbody>
              {licenses.map((item) => (
                <tr key={item.macAdresi}>
                  <td>{item.macAdresi}</td>
                  <td>{item.firmaAdi ?? "—"}</td>
                  <td>{item.license}</td>
                  <td>{item.dosyaAdi ?? "—"}</td>
                  <td>{new Date(item.updatedAt).toLocaleString("tr-TR")}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </section>

      <section className={styles.section}>
        <h2>Uzak Modüller ({modules.length})</h2>
        <p>
          Modüller <span className={styles.code}>data/modules.json</span> dosyasından okunur.
          Excel&apos;de test için: <span className={styles.code}>RunRemoteCode &quot;SelamTest&quot;</span>
        </p>
        {modules.length === 0 ? (
          <p>Henüz modül tanımlı değil.</p>
        ) : (
          <table className={styles.table}>
            <thead>
              <tr>
                <th>methodName</th>
                <th>Açıklama</th>
              </tr>
            </thead>
            <tbody>
              {modules.map((item) => (
                <tr key={item.methodName}>
                  <td>{item.methodName}</td>
                  <td>{item.description ?? "—"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </section>

      <section className={styles.section}>
        <h2>Firma Otomatik Modüller ({firmAutoModules.length})</h2>
        <p>
          Yapılandırma: <span className={styles.code}>data/firm-auto-modules.json</span>
        </p>
        {firmAutoModules.length === 0 ? (
          <p>Tanım yok.</p>
        ) : (
          <table className={styles.table}>
            <thead>
              <tr>
                <th>Firma</th>
                <th>Açıklama</th>
                <th>Excel Açılış</th>
                <th>Modüller</th>
              </tr>
            </thead>
            <tbody>
              {firmAutoModules.map((item) => (
                <tr key={item.firmaAdi}>
                  <td>
                    {item.firmaAdi === "*" ? "Tüm firmalar" : item.firmaAdi}
                  </td>
                  <td>{item.description ?? "—"}</td>
                  <td>{item.onExcelOpen.enabled ? "Açık" : "Kapalı"}</td>
                  <td>
                    {item.onExcelOpen.modules
                      .sort((a, b) => a.order - b.order)
                      .map((m) => m.methodName)
                      .join(" → ")}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </section>
    </main>
  );
}
