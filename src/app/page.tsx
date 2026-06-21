export const dynamic = "force-dynamic";

import { listFirmAutoModules } from "@/lib/firm-auto-modules";
import { listLicenses } from "@/lib/db";
import { listModules } from "@/lib/modules";
import Refresher from "./components/Refresher";
import styles from "./page.module.css";

function LicenseBadge({ value }: { value: string }) {
  const v = value.toLowerCase();
  const active = v === "true" || v === "1" || v === "active" || v === "evet";
  const cls = active ? styles.badgeGreen : styles.badgeRed;
  const label = active ? "Aktif" : "Pasif";
  return (
    <span className={`${styles.badge} ${cls}`}>
      <span className={styles.badgeDot} />
      {label}
    </span>
  );
}

export default async function Home() {
  const [licenses, modules, firmAutoModules] = await Promise.all([
    listLicenses(),
    Promise.resolve(listModules()),
    Promise.resolve(listFirmAutoModules()),
  ]);

  const activeLicenses = licenses.filter((l) => {
    const v = l.license.toLowerCase();
    return v === "true" || v === "1" || v === "active" || v === "evet";
  });

  const lastUpdated =
    licenses.length > 0
      ? new Date(licenses[0].updatedAt).toLocaleString("tr-TR", {
          day: "2-digit",
          month: "2-digit",
          hour: "2-digit",
          minute: "2-digit",
        })
      : null;

  return (
    <div className={styles.shell}>
      {/* ── Header ── */}
      <header className={styles.header}>
        <div className={styles.headerLeft}>
          <div className={styles.headerIcon}>🖥️</div>
          <div>
            <div className={styles.headerTitle}>Teklif Sunucu</div>
            <div className={styles.headerSub}>Excel VBA · Lisans &amp; Modül Yönetimi</div>
          </div>
        </div>
        <Refresher />
      </header>

      {/* ── Main ── */}
      <main className={styles.main}>

        {/* ── Stats ── */}
        <div className={styles.stats}>
          <div className={styles.statCard}>
            <div className={styles.statLabel}>Toplam Lisans</div>
            <div className={styles.statValue}>{licenses.length}</div>
            <div className={styles.statHint}>Kayıtlı cihaz</div>
          </div>
          <div className={styles.statCard}>
            <div className={styles.statLabel}>Aktif</div>
            <div className={`${styles.statValue} ${styles.green}`}>{activeLicenses.length}</div>
            <div className={styles.statHint}>Onaylı lisans</div>
          </div>
          <div className={styles.statCard}>
            <div className={styles.statLabel}>Uzak Modül</div>
            <div className={`${styles.statValue} ${styles.accent}`}>{modules.length}</div>
            <div className={styles.statHint}>VBA modülü</div>
          </div>
          <div className={styles.statCard}>
            <div className={styles.statLabel}>Son Güncelleme</div>
            <div className={styles.statValue} style={{ fontSize: "16px", marginTop: "4px" }}>
              {lastUpdated ?? "—"}
            </div>
            <div className={styles.statHint}>En son kayıt</div>
          </div>
        </div>

        {/* ── Lisanslar ── */}
        <div className={styles.card}>
          <div className={styles.cardHeader}>
            <span className={styles.cardTitle}>
              🔑 Kayıtlı Lisanslar
              <span className={styles.cardCount}>{licenses.length}</span>
            </span>
          </div>
          {licenses.length === 0 ? (
            <div className={styles.empty}>
              <div className={styles.emptyIcon}>📭</div>
              <div>Henüz kayıt yok. Excel ilk bağlandığında otomatik oluşturulur.</div>
            </div>
          ) : (
            <div className={styles.tableWrap}>
              <table className={styles.table}>
                <thead>
                  <tr>
                    <th>MAC Adresi</th>
                    <th>Firma</th>
                    <th>Kullanıcı</th>
                    <th>Dosya</th>
                    <th>IP</th>
                    <th>Lisans</th>
                    <th>Güncelleme</th>
                  </tr>
                </thead>
                <tbody>
                  {licenses.map((item) => (
                    <tr key={item.macAdresi}>
                      <td><span className={styles.mono}>{item.macAdresi}</span></td>
                      <td>{item.firmaAdi ?? <span style={{ color: "var(--text-dim)" }}>—</span>}</td>
                      <td>{item.userAdi ?? <span style={{ color: "var(--text-dim)" }}>—</span>}</td>
                      <td><span className={styles.mono}>{item.dosyaAdi ?? "—"}</span></td>
                      <td><span className={styles.mono}>{item.ipAdresi ?? "—"}</span></td>
                      <td><LicenseBadge value={item.license} /></td>
                      <td style={{ color: "var(--text-muted)", fontSize: "12px" }}>
                        {new Date(item.updatedAt).toLocaleString("tr-TR")}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>

        {/* ── İkinci satır: API + Modüller ── */}
        <div className={styles.grid2}>

          {/* API Uç Noktaları */}
          <div className={styles.card}>
            <div className={styles.cardHeader}>
              <span className={styles.cardTitle}>🔌 API Uç Noktaları</span>
            </div>
            <div className={styles.cardBody}>
              <ul className={styles.apiList}>
                {[
                  { method: "GET",  path: "/api/license/{mac}", desc: "Lisans sorgula" },
                  { method: "POST", path: "/api/license/",       desc: "Kayıt / teklif gönder" },
                  { method: "POST", path: "/api/module",         desc: "Uzak VBA modülü al" },
                  { method: "GET",  path: "/api/auto-start/{mac}", desc: "Firma modülleri" },
                  { method: "POST", path: "/api/download/teklif", desc: "Eklenti indir (.xlam)" },
                ].map((ep) => (
                  <li key={ep.path} className={styles.apiItem}>
                    <span className={`${styles.apiMethod} ${ep.method === "GET" ? styles.methodGet : styles.methodPost}`}>
                      {ep.method}
                    </span>
                    <span className={styles.apiPath}>{ep.path}</span>
                    <span className={styles.apiDesc}>{ep.desc}</span>
                  </li>
                ))}
              </ul>
            </div>
          </div>

          {/* Uzak Modüller */}
          <div className={styles.card}>
            <div className={styles.cardHeader}>
              <span className={styles.cardTitle}>
                📦 Uzak Modüller
                <span className={styles.cardCount}>{modules.length}</span>
              </span>
            </div>
            {modules.length === 0 ? (
              <div className={styles.empty}>
                <div className={styles.emptyIcon}>📭</div>
                <div>Tanımlı modül yok.</div>
              </div>
            ) : (
              <div className={styles.tableWrap}>
                <table className={styles.table}>
                  <thead>
                    <tr>
                      <th>Method</th>
                      <th>Açıklama</th>
                    </tr>
                  </thead>
                  <tbody>
                    {modules.map((m) => (
                      <tr key={m.methodName}>
                        <td><span className={styles.mono}>{m.methodName}</span></td>
                        <td className={styles.descCell}>{m.description ?? "—"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>

        </div>

        {/* ── Firma Otomatik Modüller ── */}
        <div className={styles.card}>
          <div className={styles.cardHeader}>
            <span className={styles.cardTitle}>
              ⚡ Firma Otomatik Modüller
              <span className={styles.cardCount}>{firmAutoModules.length}</span>
            </span>
          </div>
          {firmAutoModules.length === 0 ? (
            <div className={styles.empty}>
              <div className={styles.emptyIcon}>📭</div>
              <div>Tanım yok.</div>
            </div>
          ) : (
            <div className={styles.tableWrap}>
              <table className={styles.table}>
                <thead>
                  <tr>
                    <th>Firma</th>
                    <th>Açıklama</th>
                    <th>Excel Açılış</th>
                    <th>Modüller (sıra)</th>
                  </tr>
                </thead>
                <tbody>
                  {firmAutoModules.map((item) => (
                    <tr key={item.firmaAdi}>
                      <td>
                        <span className={styles.mono}>
                          {item.firmaAdi === "*" ? "Tüm firmalar" : item.firmaAdi}
                        </span>
                      </td>
                      <td className={styles.descCell}>{item.description ?? "—"}</td>
                      <td>
                        {item.onExcelOpen.enabled ? (
                          <span className={`${styles.badge} ${styles.badgeGreen}`}>
                            <span className={styles.badgeDot} /> Açık
                          </span>
                        ) : (
                          <span className={`${styles.badge} ${styles.badgeRed}`}>
                            <span className={styles.badgeDot} /> Kapalı
                          </span>
                        )}
                      </td>
                      <td className={styles.descCell}>
                        {item.onExcelOpen.modules
                          .sort((a, b) => a.order - b.order)
                          .map((m) => m.methodName)
                          .join(" → ") || "—"}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>

      </main>
    </div>
  );
}
