import { listFirmAutoModules } from "@/lib/firm-auto-modules";

export default function FirmaModulleriPage() {
  const items = listFirmAutoModules();

  return (
    <div className="page-wrap">
      <div className="page-header">
        <div>
          <div className="page-title">⚡ Firma Otomatik Modülleri</div>
          <div className="page-sub">Excel açılışında firma bazlı çalışacak modüller — data/firm-auto-modules.json</div>
        </div>
      </div>

      {items.map((item) => (
        <div key={item.firmaAdi} className="card">
          <div className="card-header">
            <span className="card-title">
              <span className="mono">
                {item.firmaAdi === "*" ? "🌐 Tüm Firmalar" : `🏢 ${item.firmaAdi}`}
              </span>
              {item.enabled !== false ? (
                <span className="badge badge-green"><span className="badge-dot" />Aktif</span>
              ) : (
                <span className="badge badge-red"><span className="badge-dot" />Devre Dışı</span>
              )}
            </span>
            {item.description && (
              <span style={{ fontSize: 12, color: "var(--text-muted)" }}>{item.description}</span>
            )}
          </div>

          <div style={{ padding: "14px 18px" }}>
            <div style={{ marginBottom: 8, fontSize: 11, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.06em", color: "var(--text-muted)" }}>
              Excel Açılış Modülleri
            </div>

            {item.onExcelOpen.enabled === false ? (
              <span className="badge badge-red"><span className="badge-dot" />Devre dışı</span>
            ) : item.onExcelOpen.modules.length === 0 ? (
              <span style={{ color: "var(--text-dim)", fontSize: 13 }}>Bu firma için tanımlı modül yok.</span>
            ) : (
              <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                {item.onExcelOpen.modules
                  .sort((a, b) => a.order - b.order)
                  .map((mod) => (
                    <div
                      key={mod.methodName}
                      style={{
                        display: "flex",
                        alignItems: "center",
                        gap: 12,
                        background: "var(--bg)",
                        border: "1px solid var(--border)",
                        borderRadius: "var(--radius-sm)",
                        padding: "10px 14px",
                      }}
                    >
                      <span
                        style={{
                          width: 24,
                          height: 24,
                          borderRadius: "50%",
                          background: "var(--accent-dim)",
                          color: "var(--accent)",
                          display: "flex",
                          alignItems: "center",
                          justifyContent: "center",
                          fontSize: 11,
                          fontWeight: 700,
                          flexShrink: 0,
                        }}
                      >
                        {mod.order}
                      </span>
                      <span className="mono" style={{ fontSize: 13, fontWeight: 600 }}>
                        {mod.methodName}
                      </span>
                      {mod.delaySeconds && mod.delaySeconds > 0 && (
                        <span className="badge badge-yellow">
                          ⏱ {mod.delaySeconds}s gecikme
                        </span>
                      )}
                    </div>
                  ))}
              </div>
            )}
          </div>
        </div>
      ))}

      <div className="card" style={{ borderColor: "var(--border-strong)" }}>
        <div className="card-header">
          <span className="card-title" style={{ color: "var(--text-muted)" }}>
            📝 Yapılandırma Dosyası
          </span>
        </div>
        <div style={{ padding: "12px 18px", fontSize: 12, color: "var(--text-muted)" }}>
          Modüller <span className="mono">data/firm-auto-modules.json</span> dosyasından okunur.
          Değişiklik için dosyayı düzenleyip Vercel&apos;e push edin.
        </div>
      </div>
    </div>
  );
}
