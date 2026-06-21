import { listModules } from "@/lib/modules";

export default function ModullerPage() {
  const modules = listModules();

  return (
    <div className="page-wrap">
      <div className="page-header">
        <div>
          <div className="page-title">📦 Uzak VBA Modülleri</div>
          <div className="page-sub">Excel açılışında RunRemoteCode ile indirilen modüller</div>
        </div>
      </div>

      <div className="card">
        <div className="card-header">
          <span className="card-title">Modüller <span className="card-count">{modules.length}</span></span>
        </div>
        {modules.length === 0 ? (
          <div className="empty-state">
            <div className="empty-state-icon">📭</div>
            <div>data/modules.json boş.</div>
          </div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 0 }}>
            {modules.map((m) => (
              <details
                key={m.methodName}
                style={{
                  borderBottom: "1px solid var(--border)",
                  padding: "0",
                }}
              >
                <summary
                  style={{
                    padding: "14px 18px",
                    cursor: "pointer",
                    display: "flex",
                    alignItems: "center",
                    gap: "12px",
                    listStyle: "none",
                    userSelect: "none",
                  }}
                >
                  <span
                    style={{
                      fontSize: 11,
                      color: "var(--text-dim)",
                      marginRight: 4,
                    }}
                  >
                    ▶
                  </span>
                  <span className="mono" style={{ fontSize: 13, fontWeight: 600 }}>
                    {m.methodName}
                  </span>
                  <span
                    style={{
                      flex: 1,
                      fontSize: 12,
                      color: "var(--text-muted)",
                    }}
                  >
                    {m.description ?? "—"}
                  </span>
                  {m.active !== false ? (
                    <span className="badge badge-green">
                      <span className="badge-dot" />Aktif
                    </span>
                  ) : (
                    <span className="badge badge-red">
                      <span className="badge-dot" />Pasif
                    </span>
                  )}
                </summary>
                <div
                  style={{
                    background: "var(--bg)",
                    borderTop: "1px solid var(--border)",
                    padding: "16px 18px",
                  }}
                >
                  <div
                    style={{
                      fontSize: 11,
                      fontWeight: 600,
                      textTransform: "uppercase",
                      letterSpacing: "0.06em",
                      color: "var(--text-muted)",
                      marginBottom: "8px",
                    }}
                  >
                    VBA Kaynak Kodu
                  </div>
                  <pre
                    style={{
                      fontFamily: "var(--font-geist-mono, monospace)",
                      fontSize: 12,
                      lineHeight: 1.7,
                      color: "var(--code-text)",
                      background: "var(--code-bg)",
                      padding: "14px 16px",
                      borderRadius: "var(--radius-sm)",
                      overflowX: "auto",
                      maxHeight: "400px",
                      overflowY: "auto",
                      whiteSpace: "pre-wrap",
                      wordBreak: "break-word",
                    }}
                  >
                    {m.code}
                  </pre>
                </div>
              </details>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
