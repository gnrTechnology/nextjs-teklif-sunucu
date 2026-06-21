export const dynamic = "force-dynamic";

import { listLicenses, listLogs } from "@/lib/db";
import { listFirmAutoModules } from "@/lib/firm-auto-modules";
import { listModules } from "@/lib/modules";
import Refresher from "./components/Refresher";

function LicenseBadge({ value }: { value: string }) {
  const active = ["true", "1", "active", "evet"].includes(value.toLowerCase());
  return (
    <span className={`badge ${active ? "badge-green" : "badge-red"}`}>
      <span className="badge-dot" />{active ? "Aktif" : "Pasif"}
    </span>
  );
}

function EventBadge({ type }: { type: string }) {
  const map: Record<string, string> = {
    register: "badge badge-blue",
    update: "badge badge-yellow",
    activate: "badge badge-green",
    deactivate: "badge badge-red",
    violation: "badge badge-red",
  };
  const labels: Record<string, string> = {
    register: "Yeni Kayıt",
    update: "Güncelleme",
    activate: "Aktifleştirildi",
    deactivate: "Pasifleştirildi",
    violation: "⚠ İhlal",
  };
  return (
    <span className={map[type] ?? "badge badge-blue"}>
      {labels[type] ?? type}
    </span>
  );
}

export default async function Dashboard() {
  const [licenses, modules, firmAutoModules, logs] = await Promise.all([
    listLicenses(),
    Promise.resolve(listModules()),
    Promise.resolve(listFirmAutoModules()),
    listLogs(10),
  ]);

  const activeLicenses = licenses.filter((l) =>
    ["true", "1", "active", "evet"].includes(l.license.toLowerCase()),
  );

  const lastUpdated = licenses.length > 0
    ? new Date(licenses[0].updatedAt).toLocaleString("tr-TR", { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" })
    : null;

  return (
    <div className="page-wrap">
      <div className="page-header">
        <div>
          <div className="page-title">📊 Dashboard</div>
          <div className="page-sub">Genel bakış ve son aktivite</div>
        </div>
        <Refresher />
      </div>

      {/* Stats */}
      <div className="stats-grid">
        <div className="stat-card">
          <div className="stat-label">Toplam Lisans</div>
          <div className="stat-value">{licenses.length}</div>
          <div className="stat-hint">Kayıtlı cihaz</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Aktif</div>
          <div className="stat-value green">{activeLicenses.length}</div>
          <div className="stat-hint">Onaylı lisans</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Uzak Modül</div>
          <div className="stat-value accent">{modules.length}</div>
          <div className="stat-hint">VBA modülü</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Son Güncelleme</div>
          <div className="stat-value" style={{ fontSize: "16px", marginTop: "4px" }}>{lastUpdated ?? "—"}</div>
          <div className="stat-hint">En son kayıt</div>
        </div>
      </div>

      {/* Son Lisanslar */}
      <div className="card">
        <div className="card-header">
          <span className="card-title">🔑 Kayıtlı Lisanslar <span className="card-count">{licenses.length}</span></span>
          <a href="/lisanslar" style={{ fontSize: 12, color: "var(--accent)" }}>Tümünü gör →</a>
        </div>
        {licenses.length === 0 ? (
          <div className="empty-state"><div className="empty-state-icon">📭</div><div>Henüz kayıt yok.</div></div>
        ) : (
          <div className="table-wrap">
            <table className="data-table">
              <thead><tr><th>MAC</th><th>Firma</th><th>Kullanıcı</th><th>Dosya</th><th>Lisans</th><th>Güncelleme</th></tr></thead>
              <tbody>
                {licenses.slice(0, 5).map((item) => (
                  <tr key={item.macAdresi}>
                    <td><span className="mono">{item.macAdresi}</span></td>
                    <td>{item.firmaAdi ?? <span style={{ color: "var(--text-dim)" }}>—</span>}</td>
                    <td>{item.userAdi ?? <span style={{ color: "var(--text-dim)" }}>—</span>}</td>
                    <td><span className="mono">{item.dosyaAdi ?? "—"}</span></td>
                    <td><LicenseBadge value={item.license} /></td>
                    <td style={{ color: "var(--text-muted)", fontSize: 12 }}>{new Date(item.updatedAt).toLocaleString("tr-TR")}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* Son Loglar */}
      <div className="card">
        <div className="card-header">
          <span className="card-title">📋 Son Aktivite</span>
          <a href="/loglar" style={{ fontSize: 12, color: "var(--accent)" }}>Tüm loglar →</a>
        </div>
        {logs.length === 0 ? (
          <div className="empty-state"><div className="empty-state-icon">📭</div><div>Log yok.</div></div>
        ) : (
          <div className="table-wrap">
            <table className="data-table">
              <thead><tr><th>Zaman</th><th>Olay</th><th>MAC</th><th>Firma</th><th>Kullanıcı</th><th>Detay</th></tr></thead>
              <tbody>
                {logs.map((log) => (
                  <tr key={log.id}>
                    <td style={{ fontSize: 12, color: "var(--text-muted)", whiteSpace: "nowrap" }}>{new Date(log.createdAt).toLocaleString("tr-TR")}</td>
                    <td><EventBadge type={log.eventType} /></td>
                    <td><span className="mono">{log.macAdresi ?? "—"}</span></td>
                    <td>{log.firmaAdi ?? "—"}</td>
                    <td>{log.userAdi ?? "—"}</td>
                    <td style={{ color: "var(--text-muted)", fontSize: 12 }}>{log.details ?? "—"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* Firma Modülleri özet */}
      <div className="grid-2">
        <div className="card">
          <div className="card-header">
            <span className="card-title">⚡ Oto. Modüller <span className="card-count">{firmAutoModules.length}</span></span>
            <a href="/firma-modulleri" style={{ fontSize: 12, color: "var(--accent)" }}>Düzenle →</a>
          </div>
          <div className="table-wrap">
            <table className="data-table">
              <thead><tr><th>Firma</th><th>Açılış</th><th>Modüller</th></tr></thead>
              <tbody>
                {firmAutoModules.map((item) => (
                  <tr key={item.firmaAdi}>
                    <td><span className="mono">{item.firmaAdi === "*" ? "Tümü" : item.firmaAdi}</span></td>
                    <td>
                      <span className={`badge ${item.onExcelOpen.enabled ? "badge-green" : "badge-red"}`}>
                        <span className="badge-dot" />{item.onExcelOpen.enabled ? "Açık" : "Kapalı"}
                      </span>
                    </td>
                    <td style={{ fontSize: 12, color: "var(--text-muted)" }}>
                      {item.onExcelOpen.modules.sort((a, b) => a.order - b.order).map((m) => m.methodName).join(" → ") || "—"}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        <div className="card">
          <div className="card-header">
            <span className="card-title">📦 Uzak Modüller <span className="card-count">{modules.length}</span></span>
            <a href="/moduller" style={{ fontSize: 12, color: "var(--accent)" }}>Tümü →</a>
          </div>
          <div className="table-wrap">
            <table className="data-table">
              <thead><tr><th>Method</th><th>Açıklama</th></tr></thead>
              <tbody>
                {modules.map((m) => (
                  <tr key={m.methodName}>
                    <td><span className="mono">{m.methodName}</span></td>
                    <td style={{ color: "var(--text-muted)", fontSize: 12 }}>{m.description ?? "—"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
}
