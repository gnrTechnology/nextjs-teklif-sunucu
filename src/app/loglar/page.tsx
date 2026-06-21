export const dynamic = "force-dynamic";

import { listLogs } from "@/lib/db";
import Refresher from "../components/Refresher";

function EventBadge({ type }: { type: string }) {
  const cfg: Record<string, { cls: string; label: string }> = {
    register:   { cls: "badge badge-blue",   label: "Yeni Kayıt"       },
    update:     { cls: "badge badge-yellow", label: "Güncelleme"       },
    activate:   { cls: "badge badge-green",  label: "Aktifleştirildi"  },
    deactivate: { cls: "badge badge-red",    label: "Pasifleştirildi"  },
    violation:  { cls: "badge badge-red",    label: "⚠ İhlal"         },
  };
  const { cls, label } = cfg[type] ?? { cls: "badge badge-blue", label: type };
  return <span className={cls}>{label}</span>;
}

export default async function LoglarPage() {
  const logs = await listLogs(500);

  const counts = logs.reduce(
    (acc, l) => {
      acc[l.eventType] = (acc[l.eventType] ?? 0) + 1;
      return acc;
    },
    {} as Record<string, number>,
  );

  return (
    <div className="page-wrap">
      <div className="page-header">
        <div>
          <div className="page-title">📋 Denetim Logları</div>
          <div className="page-sub">Her lisans başvurusu ve değişikliği kayıt altında</div>
        </div>
        <Refresher />
      </div>

      {/* Stats */}
      <div className="stats-grid">
        <div className="stat-card">
          <div className="stat-label">Toplam Log</div>
          <div className="stat-value">{logs.length}</div>
          <div className="stat-hint">Son 500 kayıt</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Yeni Kayıt</div>
          <div className="stat-value accent">{counts.register ?? 0}</div>
          <div className="stat-hint">register</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Güncelleme</div>
          <div className="stat-value yellow">{counts.update ?? 0}</div>
          <div className="stat-hint">update</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">İhlal</div>
          <div className="stat-value" style={{ color: "var(--red)" }}>{counts.violation ?? 0}</div>
          <div className="stat-hint">violation</div>
        </div>
      </div>

      <div className="card">
        <div className="card-header">
          <span className="card-title">Tüm Loglar <span className="card-count">{logs.length}</span></span>
        </div>
        {logs.length === 0 ? (
          <div className="empty-state">
            <div className="empty-state-icon">📭</div>
            <div>Henüz log kaydı yok. İlk lisans başvurusu geldiğinde burada görünür.</div>
          </div>
        ) : (
          <div className="table-wrap">
            <table className="data-table">
              <thead>
                <tr>
                  <th>#</th>
                  <th>Tarih / Saat</th>
                  <th>Olay</th>
                  <th>MAC Adresi</th>
                  <th>Firma</th>
                  <th>Kullanıcı</th>
                  <th>Dosya</th>
                  <th>IP</th>
                  <th>Detay</th>
                </tr>
              </thead>
              <tbody>
                {logs.map((log) => (
                  <tr key={log.id} style={log.eventType === "violation" ? { background: "var(--red-dim)" } : undefined}>
                    <td style={{ color: "var(--text-dim)", fontSize: 11 }}>{log.id}</td>
                    <td style={{ fontSize: 12, color: "var(--text-muted)", whiteSpace: "nowrap" }}>
                      {new Date(log.createdAt).toLocaleString("tr-TR")}
                    </td>
                    <td><EventBadge type={log.eventType} /></td>
                    <td><span className="mono">{log.macAdresi ?? "—"}</span></td>
                    <td>{log.firmaAdi ?? "—"}</td>
                    <td>{log.userAdi ?? "—"}</td>
                    <td>
                      <span className="mono">{log.dosyaAdi ?? "—"}</span>
                      {log.dosyaAdi && log.dosyaAdi !== "teklif.xlam" && (
                        <span title="Dosya adı teklif.xlam değil!" style={{ marginLeft: 4, color: "var(--yellow)" }}>⚠</span>
                      )}
                    </td>
                    <td><span className="mono" style={{ fontSize: 11 }}>{log.ipAdresi ?? "—"}</span></td>
                    <td style={{ color: "var(--text-muted)", fontSize: 12 }}>{log.details ?? "—"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}
