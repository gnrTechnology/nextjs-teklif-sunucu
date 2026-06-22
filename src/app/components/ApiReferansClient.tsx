"use client";

import { useMemo, useState } from "react";
import { API_ENDPOINTS, API_GROUPS, getApiCatalogStats } from "@/lib/api-catalog";
import Link from "next/link";

function MethodBadge({ method }: { method: string }) {
  const primary = method.split("|")[0].trim();
  const cls =
    primary === "GET" ? "method-get" :
    primary === "POST" ? "method-post" :
    primary === "PATCH" ? "method-patch" :
    primary === "DELETE" ? "method-delete" : "method-get";
  return (
    <span className={`badge ${cls}`} style={{ fontFamily: "monospace", fontWeight: 700, fontSize: 11 }}>
      {method}
    </span>
  );
}

export default function ApiReferansClient() {
  const [group, setGroup] = useState<string>("all");
  const [search, setSearch] = useState("");
  const stats = getApiCatalogStats();

  const filtered = useMemo(() => {
    return API_ENDPOINTS.filter((ep) => {
      if (group !== "all" && ep.group !== group) return false;
      if (!search) return true;
      const q = search.toLowerCase();
      return (
        ep.path.toLowerCase().includes(q) ||
        ep.title.toLowerCase().includes(q) ||
        ep.desc.toLowerCase().includes(q) ||
        ep.group.toLowerCase().includes(q)
      );
    });
  }, [group, search]);

  const baseUrl =
    typeof window !== "undefined"
      ? `${window.location.origin}/api`
      : "https://nextjs-teklif-sunucu.vercel.app/api";

  return (
    <div className="page-wrap">
      <div className="page-header">
        <div>
          <div className="page-title">API Referans</div>
          <div className="page-sub">
            <span className="mono">{baseUrl}</span>
            {" · "}
            {stats.total} endpoint · {stats.withUi} UI bağlantılı
          </div>
        </div>
        <Link href="/oneriler" className="btn btn-ghost">Modül Önerileri →</Link>
      </div>

      <div className="card" style={{ marginBottom: 16, padding: "14px 18px" }}>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 10, alignItems: "center" }}>
          <input
            className="form-input"
            placeholder="Endpoint ara…"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            style={{ flex: "1 1 200px", maxWidth: 360 }}
          />
          <select className="form-input" value={group} onChange={(e) => setGroup(e.target.value)} style={{ width: 180 }}>
            <option value="all">Tüm gruplar</option>
            {API_GROUPS.map((g) => (
              <option key={g} value={g}>{g}</option>
            ))}
          </select>
          <span style={{ fontSize: 12, color: "var(--text-dim)" }}>{filtered.length} endpoint</span>
        </div>
      </div>

      {filtered.map((ep) => (
        <div key={`${ep.method}-${ep.path}`} className="card" style={{ marginBottom: "1rem" }}>
          <div className="card-header">
            <span className="card-title" style={{ gap: 12, flexWrap: "wrap" }}>
              <MethodBadge method={ep.method} />
              <span className="mono" style={{ fontSize: 14 }}>{ep.path}</span>
              <span style={{ color: "var(--text-muted)", fontWeight: 400, fontSize: 13 }}>{ep.title}</span>
              <span className="badge badge-blue" style={{ fontSize: 10 }}>{ep.group}</span>
              {ep.uiLink && (
                <Link href={ep.uiLink} style={{ fontSize: 11, color: "var(--accent)" }}>UI →</Link>
              )}
            </span>
          </div>

          <div style={{ padding: "14px 18px", display: "flex", flexDirection: "column", gap: 14 }}>
            <p style={{ fontSize: 13, color: "var(--text-muted)" }}>{ep.desc}</p>
            {ep.tags && (
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                {ep.tags.map((t) => (
                  <span key={t} className="badge badge-yellow" style={{ fontSize: 10 }}>{t}</span>
                ))}
              </div>
            )}

            <div style={{ display: "grid", gridTemplateColumns: ep.request ? "1fr 1fr" : "1fr", gap: 14 }}>
              {ep.request && (
                <div>
                  <div style={LABEL}>Request Body</div>
                  <pre style={PRE}>{ep.request}</pre>
                </div>
              )}
              {ep.response && (
                <div>
                  <div style={LABEL}>Örnek Yanıt</div>
                  <pre style={PRE}>{ep.response}</pre>
                </div>
              )}
            </div>

            <div>
              <div style={LABEL}>HTTP Kodları</div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 12 }}>
                {ep.responses.map((r) => (
                  <span key={r.code} style={{ fontSize: 12, color: "var(--text-muted)" }}>
                    <span className="mono">{r.code}</span> {r.desc}
                  </span>
                ))}
              </div>
            </div>
          </div>
        </div>
      ))}
    </div>
  );
}

const LABEL: React.CSSProperties = {
  fontSize: 11,
  fontWeight: 600,
  textTransform: "uppercase",
  letterSpacing: "0.06em",
  color: "var(--text-muted)",
  marginBottom: 6,
};

const PRE: React.CSSProperties = {
  fontFamily: "var(--font-geist-mono, monospace)",
  fontSize: 12,
  lineHeight: 1.7,
  color: "var(--code-text)",
  background: "var(--code-bg)",
  padding: "12px 14px",
  borderRadius: "var(--radius-sm)",
  overflowX: "auto",
  whiteSpace: "pre-wrap",
};
