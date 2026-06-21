export const dynamic = "force-dynamic";

import { listLicenses } from "@/lib/db";
import Refresher from "../components/Refresher";
import LisanslarTable from "../components/LisanslarTable";

export default async function LisanslarPage() {
  const licenses = await listLicenses();
  const active = licenses.filter((l) =>
    ["true", "1", "active", "evet"].includes(l.license.toLowerCase()),
  );
  const violations = licenses.filter(
    (l) => l.dosyaAdi && l.dosyaAdi.toLowerCase() !== "teklif.xlam",
  );

  return (
    <div className="page-wrap">
      <div className="page-header">
        <div>
          <div className="page-title">🔑 Lisanslar</div>
          <div className="page-sub">
            {licenses.length} kayıtlı cihaz · {active.length} aktif
            {violations.length > 0 && (
              <span style={{ marginLeft: 8, color: "var(--red)", fontWeight: 600 }}>
                · ⚠ {violations.length} ihlal
              </span>
            )}
          </div>
        </div>
        <Refresher />
      </div>

      <div className="stats-grid" style={{ gridTemplateColumns: "repeat(4, 1fr)" }}>
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
        <div className="stat-card">
          <div className="stat-label">İhlal</div>
          <div className="stat-value red">{violations.length}</div>
          <div className="stat-hint">Kopya tespit edildi</div>
        </div>
      </div>

      <div className="card">
        <div className="card-header">
          <span className="card-title">
            Kayıtlı Lisanslar <span className="card-count">{licenses.length}</span>
          </span>
          <span style={{ fontSize: 12, color: "var(--text-muted)" }}>
            Checkbox ile seçip toplu işlem yapabilirsiniz
          </span>
        </div>
        {licenses.length === 0 ? (
          <div className="empty-state">
            <div className="empty-state-icon">📭</div>
            <div>Henüz lisans kaydı yok. Excel ilk bağlandığında otomatik oluşturulur.</div>
          </div>
        ) : (
          <LisanslarTable licenses={licenses} />
        )}
      </div>
    </div>
  );
}
