"use client";

import { useMemo, useState } from "react";
import type { ProposalItem, ProposalsSummary } from "@/lib/proposals";

type Filter = "all" | "done" | "planned" | "inDb" | "missing";

function StatusBadge({ item, inDb }: { item: ProposalItem; inDb: boolean }) {
  if (item.status === "done" && inDb) {
    return <span className="badge badge-green"><span className="badge-dot" />DB&apos;de</span>;
  }
  if (item.status === "done" && !inDb) {
    return <span className="badge badge-yellow"><span className="badge-dot" />Eksik</span>;
  }
  if (item.status === "blocked") {
    return <span className="badge badge-red"><span className="badge-dot" />Kısıtlı</span>;
  }
  return <span className="badge badge-blue"><span className="badge-dot" />Planlı</span>;
}

export default function OnerilerClient({ summary }: { summary: ProposalsSummary }) {
  const [filter, setFilter] = useState<Filter>("all");
  const [search, setSearch] = useState("");
  const [source, setSource] = useState<"all" | "vba" | "dll">("all");
  const [seeding, setSeeding] = useState(false);
  const [seedMsg, setSeedMsg] = useState("");

  const dbSet = useMemo(
    () => new Set(summary.implementedInDb.map((n) => n.toLowerCase())),
    [summary.implementedInDb],
  );

  const filtered = useMemo(() => {
    return summary.items.filter((item) => {
      if (source !== "all" && item.source !== source) return false;
      const inDb = dbSet.has(item.methodName.toLowerCase());
      if (filter === "done" && item.status !== "done") return false;
      if (filter === "planned" && item.status !== "planned") return false;
      if (filter === "inDb" && !inDb) return false;
      if (filter === "missing" && !(item.status === "done" && !inDb)) return false;
      if (search) {
        const q = search.toLowerCase();
        return (
          item.methodName.toLowerCase().includes(q) ||
          item.description.toLowerCase().includes(q) ||
          item.section.toLowerCase().includes(q)
        );
      }
      return true;
    });
  }, [summary.items, filter, search, source, dbSet]);

  async function runSeed(endpoint: string, label: string) {
    setSeeding(true);
    setSeedMsg("");
    try {
      const r = await fetch(endpoint);
      const j = await r.json();
      setSeedMsg(j.success ? `✓ ${label}: ${j.message ?? j.seeded + " kayıt"}` : `✗ ${j.error ?? "Hata"}`);
    } catch (e) {
      setSeedMsg(`✗ ${String(e)}`);
    }
    setSeeding(false);
  }

  const { stats } = summary;

  return (
    <div className="page-wrap">
      <div className="page-header">
        <div>
          <div className="page-title">Modül Önerileri</div>
          <div className="page-sub">
            module-proposals.md + dll önerileri — uygulanan ve planlanan modüller
          </div>
        </div>
      </div>

      <div className="stats-grid" style={{ marginBottom: 20 }}>
        <div className="stat-card">
          <div className="stat-label">Toplam Öneri</div>
          <div className="stat-value">{stats.total}</div>
          <div className="stat-hint">VBA + DLL listesi</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Uygulanan (✅)</div>
          <div className="stat-value green">{stats.done}</div>
          <div className="stat-hint">{stats.missingFromDb} DB&apos;de eksik</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">DB Modül</div>
          <div className="stat-value accent">{stats.inDb}</div>
          <div className="stat-hint">{stats.extraInDb} öneri dışı</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Planlanan</div>
          <div className="stat-value" style={{ color: "var(--yellow)" }}>{stats.planned}</div>
          <div className="stat-hint">{stats.blocked} kısıtlı</div>
        </div>
      </div>

      <div className="card" style={{ marginBottom: 16, padding: "14px 18px" }}>
        <div style={{ fontSize: 12, fontWeight: 600, color: "var(--text-muted)", marginBottom: 10 }}>
          Yönetim — Seed &amp; Senkron
        </div>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 8, alignItems: "center" }}>
          <button className="btn btn-primary" disabled={seeding} onClick={() => runSeed("/api/seed-modules/", "Modül seed")}>
            modules.json → DB
          </button>
          <button className="btn btn-ghost" disabled={seeding} onClick={() => runSeed("/api/firm-modules/seed/", "Firma seed")}>
            firm-auto-modules → DB
          </button>
          <a href="/api-referans" className="btn btn-ghost">API Referans →</a>
          <a href="/moduller" className="btn btn-ghost">Uzak Modüller →</a>
        </div>
        {seedMsg && (
          <div style={{ marginTop: 10, fontSize: 12, color: "var(--text-muted)" }}>{seedMsg}</div>
        )}
      </div>

      <div className="card" style={{ marginBottom: 16, padding: "14px 18px" }}>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 10, alignItems: "center" }}>
          <input
            className="form-input"
            placeholder="Modül veya açıklama ara…"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            style={{ flex: "1 1 200px", maxWidth: 320 }}
          />
          <select className="form-input" value={filter} onChange={(e) => setFilter(e.target.value as Filter)} style={{ width: 140 }}>
            <option value="all">Tümü</option>
            <option value="done">Uygulanan</option>
            <option value="planned">Planlanan</option>
            <option value="inDb">DB&apos;de var</option>
            <option value="missing">✅ ama DB eksik</option>
          </select>
          <select className="form-input" value={source} onChange={(e) => setSource(e.target.value as "all" | "vba" | "dll")} style={{ width: 120 }}>
            <option value="all">Tüm kaynak</option>
            <option value="vba">VBA</option>
            <option value="dll">DLL</option>
          </select>
          <span style={{ fontSize: 12, color: "var(--text-dim)" }}>{filtered.length} sonuç</span>
        </div>
      </div>

      <div className="card" style={{ overflow: "hidden" }}>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr style={{ borderBottom: "1px solid var(--border)" }}>
              {["#", "MethodName", "Açıklama", "Bölüm", "Kaynak", "Durum"].map((h) => (
                <th key={h} style={TH}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.slice(0, 200).map((item) => {
              const inDb = dbSet.has(item.methodName.toLowerCase());
              return (
                <tr key={`${item.source}-${item.methodName}`} style={{ borderBottom: "1px solid var(--border)" }}>
                  <td style={TD}><span style={{ color: "var(--text-dim)", fontSize: 11 }}>{item.number}</span></td>
                  <td style={TD}><span className="mono" style={{ fontSize: 12 }}>{item.methodName}</span></td>
                  <td style={{ ...TD, fontSize: 12, color: "var(--text-muted)", maxWidth: 280 }}>{item.description}</td>
                  <td style={{ ...TD, fontSize: 11, color: "var(--text-dim)" }}>{item.section}</td>
                  <td style={TD}><span className="badge badge-blue">{item.source.toUpperCase()}</span></td>
                  <td style={TD}><StatusBadge item={item} inDb={inDb} /></td>
                </tr>
              );
            })}
          </tbody>
        </table>
        {filtered.length > 200 && (
          <div style={{ padding: 12, fontSize: 12, color: "var(--text-dim)", textAlign: "center" }}>
            İlk 200 kayıt gösteriliyor — arama ile daraltın
          </div>
        )}
      </div>
    </div>
  );
}

const TH: React.CSSProperties = {
  padding: "10px 14px",
  textAlign: "left",
  fontSize: 11,
  fontWeight: 600,
  textTransform: "uppercase",
  letterSpacing: "0.05em",
  color: "var(--text-muted)",
};

const TD: React.CSSProperties = {
  padding: "10px 14px",
  verticalAlign: "middle",
};
