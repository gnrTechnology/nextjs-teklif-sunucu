import { listDbModules, listHeartbeats, listClientCommands, listModuleOutputs, listLicenses } from "@/lib/db";
import { formatTR } from "@/lib/date-utils";

export default async function AnalitikPage() {
  const [modules, heartbeats, commands, outputs, licenses] = await Promise.all([
    listDbModules(),
    listHeartbeats(),
    listClientCommands({ limit: 500 }),
    listModuleOutputs({ limit: 500 }),
    listLicenses(),
  ]);

  /* ── Modül istatistikleri ── */
  const topModules = [...modules]
    .filter((m) => (m.runCount ?? 0) > 0)
    .sort((a, b) => (b.runCount ?? 0) - (a.runCount ?? 0))
    .slice(0, 20);

  const totalRuns = modules.reduce((s, m) => s + (m.runCount ?? 0), 0);
  const activeModules = modules.filter((m) => m.active !== false).length;

  /* ── Cihaz istatistikleri ── */
  const now = Date.now();
  const onlineCount  = heartbeats.filter((h) => now - new Date(h.last_seen).getTime() < 5 * 60 * 1000).length;
  const idleCount    = heartbeats.filter((h) => { const ms = now - new Date(h.last_seen).getTime(); return ms >= 5*60000 && ms < 60*60000; }).length;
  const offlineCount = heartbeats.length - onlineCount - idleCount;

  /* ── Komut istatistikleri ── */
  const cmdDone    = commands.filter((c) => c.status === "done").length;
  const cmdError   = commands.filter((c) => c.status === "error").length;
  const cmdPending = commands.filter((c) => c.status === "pending" || c.status === "running").length;

  /* ── Lisans istatistikleri ── */
  const activeLicenses = licenses.filter((l) => ["true","1","active","evet"].includes(l.license.toLowerCase())).length;

  /* ── Kategori dağılımı ── */
  const catCounts: Record<string, number> = {};
  modules.forEach((m) => {
    const cat = m.category ?? "genel";
    catCounts[cat] = (catCounts[cat] ?? 0) + 1;
  });
  const catEntries = Object.entries(catCounts).sort((a, b) => b[1] - a[1]);

  const CAT_COLOR: Record<string, string> = {
    lisans: "#10b981", sistem: "#f59e0b", dosya: "#8b5cf6", internet: "#3b82f6",
    excel: "#10b981", powershell: "#f59e0b", registry: "#f97316", guvenlik: "#ef4444",
    bildirim: "#3b82f6", genel: "#6b7280", zamanlanmis: "#a78bfa", uzman: "#ec4899",
    donanim: "#06b6d4",
  };

  /* ── Module outputs özet ── */
  const outputByModule: Record<string, number> = {};
  outputs.forEach((o) => {
    outputByModule[o.moduleName] = (outputByModule[o.moduleName] ?? 0) + 1;
  });

  return (
    <div className="page-wrap">
      <div className="page-header">
        <div>
          <div className="page-title">Analitik & İstatistikler</div>
          <div className="page-sub">Modül kullanımı, cihaz durumu ve sistem özeti</div>
        </div>
      </div>

      {/* Özet Kartlar */}
      <div className="stats-grid" style={{ marginBottom: 20 }}>
        <div className="stat-card">
          <div className="stat-label">Toplam Modül Çalışması</div>
          <div className="stat-value accent">{totalRuns}</div>
          <div className="stat-hint">{activeModules} aktif modül</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Online Cihaz</div>
          <div className="stat-value green">{onlineCount}</div>
          <div className="stat-hint">{idleCount} boşta · {offlineCount} çevrimdışı</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Komut Kuyruğu</div>
          <div className="stat-value">{cmdDone + cmdError}</div>
          <div className="stat-hint">{cmdPending} bekliyor · {cmdError} hata</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Lisans / Cihaz</div>
          <div className="stat-value">{licenses.length}</div>
          <div className="stat-hint">{activeLicenses} aktif</div>
        </div>
      </div>

      <div className="grid-2">
        {/* Top Modüller */}
        <div className="card">
          <div className="card-header">
            <span className="card-title">🏆 En Çok Çalışan Modüller</span>
            <span style={{ fontSize: 12, color: "var(--text-dim)" }}>Toplam {totalRuns} çalışma</span>
          </div>
          {topModules.length === 0 ? (
            <div className="empty-state">
              <div className="empty-state-icon">📊</div>
              <div>Henüz modül çalışma verisi yok.</div>
            </div>
          ) : (
            <div style={{ padding: "8px 0" }}>
              {topModules.map((m, i) => {
                const pct = totalRuns > 0 ? ((m.runCount ?? 0) / totalRuns) * 100 : 0;
                const color = CAT_COLOR[m.category ?? "genel"] ?? "#6b7280";
                return (
                  <div key={m.methodName} style={{
                    padding: "10px 18px",
                    borderBottom: "1px solid var(--border)",
                  }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 5 }}>
                      <span style={{
                        width: 22, height: 22, borderRadius: "50%",
                        background: "var(--accent-dim)", color: "var(--accent)",
                        display: "flex", alignItems: "center", justifyContent: "center",
                        fontSize: 10, fontWeight: 700, flexShrink: 0,
                      }}>{i + 1}</span>
                      <span className="mono" style={{ fontSize: 13, flex: 1, fontWeight: 600 }}>
                        {m.methodName}
                      </span>
                      <span style={{
                        fontSize: 10, padding: "1px 7px", borderRadius: 10,
                        background: color + "20", color,
                      }}>{m.category ?? "genel"}</span>
                      <span style={{ fontSize: 13, fontWeight: 700, color: "var(--accent)" }}>
                        {m.runCount}×
                      </span>
                    </div>
                    <div style={{
                      height: 4, background: "var(--border)", borderRadius: 2, overflow: "hidden",
                    }}>
                      <div style={{
                        height: "100%", width: `${pct}%`,
                        background: color, borderRadius: 2, transition: "width 0.5s",
                      }} />
                    </div>
                    {m.lastRunAt && (
                      <div style={{ fontSize: 10, color: "var(--text-dim)", marginTop: 3 }}>
                        Son: {formatTR(m.lastRunAt)}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}
        </div>

        {/* Kategori dağılımı + Cihaz durumu */}
        <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
          <div className="card">
            <div className="card-header">
              <span className="card-title">📦 Kategori Dağılımı</span>
              <span style={{ fontSize: 12, color: "var(--text-dim)" }}>{modules.length} modül</span>
            </div>
            <div style={{ padding: "10px 18px", display: "flex", flexDirection: "column", gap: 8 }}>
              {catEntries.map(([cat, cnt]) => {
                const color = CAT_COLOR[cat] ?? "#6b7280";
                const pct = (cnt / modules.length) * 100;
                return (
                  <div key={cat}>
                    <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, marginBottom: 3 }}>
                      <span style={{ color }}>{cat}</span>
                      <span style={{ color: "var(--text-dim)" }}>{cnt}</span>
                    </div>
                    <div style={{ height: 6, background: "var(--border)", borderRadius: 3, overflow: "hidden" }}>
                      <div style={{ height: "100%", width: `${pct}%`, background: color, borderRadius: 3 }} />
                    </div>
                  </div>
                );
              })}
            </div>
          </div>

          {/* Cihaz durumu */}
          <div className="card">
            <div className="card-header">
              <span className="card-title">📡 Cihaz Durumu</span>
            </div>
            <div style={{ padding: "14px 18px" }}>
              {[
                { label: "Online (< 5dk)",     count: onlineCount,  color: "#10b981" },
                { label: "Boşta (5dk–1sa)",    count: idleCount,    color: "#f59e0b" },
                { label: "Çevrimdışı (> 1sa)", count: offlineCount, color: "#6b7280" },
              ].map(({ label, count, color }) => (
                <div key={label} style={{
                  display: "flex", alignItems: "center", gap: 10,
                  padding: "10px 0", borderBottom: "1px solid var(--border)",
                }}>
                  <span style={{
                    width: 10, height: 10, borderRadius: "50%", background: color, flexShrink: 0,
                  }} />
                  <span style={{ flex: 1, fontSize: 13 }}>{label}</span>
                  <span style={{ fontWeight: 700, fontSize: 18, color }}>{count}</span>
                </div>
              ))}
              <div style={{ marginTop: 10, fontSize: 12, color: "var(--text-dim)" }}>
                Toplam {heartbeats.length} kayıtlı cihaz
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Module Outputs özet */}
      {Object.keys(outputByModule).length > 0 && (
        <div className="card" style={{ marginTop: 16 }}>
          <div className="card-header">
            <span className="card-title">📤 Gönderilen Modül Çıktıları</span>
            <span style={{ fontSize: 12, color: "var(--text-dim)" }}>{outputs.length} kayıt</span>
          </div>
          <div className="table-wrap">
            <table className="data-table">
              <thead><tr><th>Modül</th><th>Kayıt Sayısı</th><th>Son Gönderim</th></tr></thead>
              <tbody>
                {Object.entries(outputByModule).sort((a, b) => b[1] - a[1]).map(([mod, cnt]) => {
                  const last = outputs.find((o) => o.moduleName === mod);
                  return (
                    <tr key={mod}>
                      <td><span className="mono">{mod}</span></td>
                      <td>{cnt}</td>
                      <td style={{ fontSize: 12, color: "var(--text-muted)" }}>
                        {last ? formatTR(last.createdAt) : "—"}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}
