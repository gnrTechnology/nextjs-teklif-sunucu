export const dynamic = "force-dynamic";

import { listLicenses } from "@/lib/db";
import Refresher from "../components/Refresher";
import ToggleLicenseButton from "../components/ToggleLicenseButton";

function LicenseBadge({ value }: { value: string }) {
  const active = ["true", "1", "active", "evet"].includes(value.toLowerCase());
  return (
    <span className={`badge ${active ? "badge-green" : "badge-red"}`}>
      <span className="badge-dot" />{active ? "Aktif" : "Pasif"}
    </span>
  );
}

export default async function LisanslarPage() {
  const licenses = await listLicenses();
  const active = licenses.filter((l) => ["true", "1", "active", "evet"].includes(l.license.toLowerCase()));

  return (
    <div className="page-wrap">
      <div className="page-header">
        <div>
          <div className="page-title">🔑 Lisanslar</div>
          <div className="page-sub">{licenses.length} kayıtlı cihaz · {active.length} aktif</div>
        </div>
        <Refresher />
      </div>

      {/* Stats */}
      <div className="stats-grid" style={{ gridTemplateColumns: "repeat(3, 1fr)" }}>
        <div className="stat-card">
          <div className="stat-label">Toplam</div>
          <div className="stat-value">{licenses.length}</div>
          <div className="stat-hint">Kayıtlı cihaz</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Aktif</div>
          <div className="stat-value green">{active.length}</div>
          <div className="stat-hint">Onaylı</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Pasif</div>
          <div className="stat-value yellow">{licenses.length - active.length}</div>
          <div className="stat-hint">Beklemede / Devre dışı</div>
        </div>
      </div>

      <div className="card">
        <div className="card-header">
          <span className="card-title">Kayıtlı Lisanslar <span className="card-count">{licenses.length}</span></span>
        </div>
        {licenses.length === 0 ? (
          <div className="empty-state">
            <div className="empty-state-icon">📭</div>
            <div>Henüz lisans kaydı yok. Excel ilk bağlandığında otomatik oluşturulur.</div>
          </div>
        ) : (
          <div className="table-wrap">
            <table className="data-table">
              <thead>
                <tr>
                  <th>MAC Adresi</th>
                  <th>Firma</th>
                  <th>Kullanıcı</th>
                  <th>Dosya</th>
                  <th>IP Adresi</th>
                  <th>Lisans</th>
                  <th>Oluşturma</th>
                  <th>Güncelleme</th>
                  <th>İşlem</th>
                </tr>
              </thead>
              <tbody>
                {licenses.map((item) => (
                  <tr key={item.macAdresi}>
                    <td><span className="mono">{item.macAdresi}</span></td>
                    <td>{item.firmaAdi ?? <span style={{ color: "var(--text-dim)" }}>—</span>}</td>
                    <td>{item.userAdi ?? <span style={{ color: "var(--text-dim)" }}>—</span>}</td>
                    <td>
                      <span className={`mono ${item.dosyaAdi && item.dosyaAdi !== "teklif.xlam" ? "mono-warn" : ""}`}>
                        {item.dosyaAdi ?? "—"}
                      </span>
                      {item.dosyaAdi && item.dosyaAdi !== "teklif.xlam" && (
                        <span title="Dosya adı teklif.xlam değil!" style={{ marginLeft: 4, color: "var(--yellow)" }}>⚠</span>
                      )}
                    </td>
                    <td><span className="mono">{item.ipAdresi ?? "—"}</span></td>
                    <td><LicenseBadge value={item.license} /></td>
                    <td style={{ color: "var(--text-muted)", fontSize: 12, whiteSpace: "nowrap" }}>
                      {new Date(item.createdAt).toLocaleString("tr-TR")}
                    </td>
                    <td style={{ color: "var(--text-muted)", fontSize: 12, whiteSpace: "nowrap" }}>
                      {new Date(item.updatedAt).toLocaleString("tr-TR")}
                    </td>
                    <td>
                      <ToggleLicenseButton mac={item.macAdresi} current={item.license} />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* TODO: dosyaAdi güvenlik kontrolü */}
      <div className="card" style={{ borderColor: "var(--yellow)", background: "var(--yellow-dim)" }}>
        <div className="card-header" style={{ borderColor: "var(--yellow)" }}>
          <span className="card-title" style={{ color: "var(--yellow)" }}>
            ⚠ TODO: Güvenlik Kontrolü — dosyaAdi İhlal Tespiti
          </span>
        </div>
        <div style={{ padding: "14px 18px", fontSize: 13, color: "var(--text-muted)", lineHeight: 1.8 }}>
          <strong style={{ color: "var(--yellow)" }}>Planlanmış özellik:</strong> Her lisans başvurusunda{" "}
          <span className="mono">dosyaAdi === &quot;teklif.xlam&quot;</span> kontrolü yapılacak.
          Farklıysa (dosya kopyalanmış / yeniden adlandırılmış):
          <ol style={{ marginTop: "0.5rem", paddingLeft: "1.5rem", display: "flex", flexDirection: "column", gap: "4px" }}>
            <li>Lisans <strong>pasife</strong> alınır</li>
            <li>Log&apos;a <span className="mono">event_type=violation</span> yazılır</li>
            <li>Yanıtta <span className="mono">{"{ action: 'delete', targets: ['copy','addin'] }"}</span> döndürülür</li>
            <li>VBA bu yanıtı alınca kopya dosyayı ve <span className="mono">teklif.xlam</span>&apos;ı bilgisayardan siler</li>
          </ol>
          <div style={{ marginTop: "0.5rem" }}>
            Bakınız: <span className="mono">src/app/api/license/route.ts</span> — TODO yorumu
          </div>
        </div>
      </div>
    </div>
  );
}
