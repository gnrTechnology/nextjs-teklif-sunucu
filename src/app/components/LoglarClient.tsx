"use client";

import { useCallback, useEffect, useMemo, useState } from "react";
import type { ActivityCategory, UnifiedActivityItem } from "@/lib/types";
import { ACTIVITY_CATEGORY_LABELS } from "@/lib/activity";
import { formatTR } from "@/lib/date-utils";
import Refresher from "./Refresher";

const CATEGORIES: { id: ActivityCategory; label: string }[] = [
  { id: "all", label: "Tümü" },
  { id: "lisans", label: "Lisans" },
  { id: "ihlal", label: "İhlal" },
  { id: "guncelleme", label: "Güncelleme" },
  { id: "dashboard", label: "Dashboard" },
  { id: "modul", label: "Modül Çıktısı" },
  { id: "heartbeat", label: "Heartbeat" },
  { id: "komut", label: "Komut" },
  { id: "klasor", label: "Klasör" },
  { id: "cihaz", label: "Cihaz" },
];

function CategoryBadge({ category }: { category: Exclude<ActivityCategory, "all"> }) {
  const cfg = ACTIVITY_CATEGORY_LABELS[category];
  return (
    <span className="badge badge-blue" style={{ fontSize: 11 }}>
      {cfg.icon} {cfg.label}
    </span>
  );
}

export default function LoglarClient({ initial }: { initial: UnifiedActivityItem[] }) {
  const [items, setItems] = useState(initial);
  const [category, setCategory] = useState<ActivityCategory>("all");
  const [search, setSearch] = useState("");
  const [loading, setLoading] = useState(false);

  const refresh = useCallback(async () => {
    setLoading(true);
    try {
      const q = category !== "all" ? `?category=${category}&limit=400` : "?limit=400";
      const r = await fetch(`/api/activity${q}`);
      const j = await r.json();
      if (j.success) setItems(j.data);
    } finally {
      setLoading(false);
    }
  }, [category]);

  useEffect(() => {
    refresh();
  }, [category, refresh]);

  const filtered = useMemo(() => {
    if (!search.trim()) return items;
    const q = search.toLowerCase();
    return items.filter(
      (i) =>
        i.title.toLowerCase().includes(q) ||
        (i.detail ?? "").toLowerCase().includes(q) ||
        (i.mac ?? "").toLowerCase().includes(q) ||
        (i.source ?? "").toLowerCase().includes(q),
    );
  }, [items, search]);

  const counts = useMemo(() => {
    const c: Record<string, number> = {};
    for (const item of items) c[item.category] = (c[item.category] ?? 0) + 1;
    return c;
  }, [items]);

  return (
    <div className="page-wrap">
      <div className="page-header">
        <div>
          <div className="page-title">Denetim &amp; Aktivite Logları</div>
          <div className="page-sub">
            Lisans, ihlal, dashboard, modül çıktıları, heartbeat, komutlar ve klasör izleme — tek ekran
          </div>
        </div>
        <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
          {loading && <span style={{ fontSize: 12, color: "var(--text-dim)" }}>Yükleniyor…</span>}
          <Refresher />
          <button className="btn btn-ghost" onClick={refresh}>↻ Yenile</button>
        </div>
      </div>

      <div
        style={{
          display: "flex",
          flexWrap: "wrap",
          gap: 8,
          marginBottom: 16,
        }}
      >
        {CATEGORIES.map((c) => (
          <button
            key={c.id}
            className={`btn ${category === c.id ? "btn-primary" : "btn-ghost"}`}
            style={{ fontSize: 12, padding: "6px 12px" }}
            onClick={() => setCategory(c.id)}
          >
            {c.label}
            {c.id !== "all" && counts[c.id] != null && (
              <span style={{ marginLeft: 6, opacity: 0.7 }}>({counts[c.id]})</span>
            )}
            {c.id === "all" && <span style={{ marginLeft: 6, opacity: 0.7 }}>({items.length})</span>}
          </button>
        ))}
      </div>

      <div className="card" style={{ marginBottom: 16, padding: "12px 16px" }}>
        <input
          className="form-input"
          placeholder="Başlık, MAC, modül veya detay ara…"
          value={search}
          onChange={(e) => setSearch(e.target.value)}
        />
      </div>

      <div className="card">
        <div className="card-header">
          <span className="card-title">
            Aktivite Akışı <span className="card-count">{filtered.length}</span>
          </span>
        </div>
        {filtered.length === 0 ? (
          <div className="empty-state">
            <div className="empty-state-icon">📭</div>
            <div>Bu kategoride kayıt yok.</div>
          </div>
        ) : (
          <div className="table-wrap">
            <table className="data-table">
              <thead>
                <tr>
                  <th>Zaman</th>
                  <th>Kategori</th>
                  <th>Başlık</th>
                  <th>MAC / PC</th>
                  <th>Kaynak</th>
                  <th>Detay</th>
                </tr>
              </thead>
              <tbody>
                {filtered.map((item) => (
                  <tr
                    key={item.id}
                    style={
                      item.category === "ihlal"
                        ? { background: "var(--red-dim)" }
                        : undefined
                    }
                  >
                    <td style={{ fontSize: 12, color: "var(--text-muted)", whiteSpace: "nowrap" }}>
                      {formatTR(item.createdAt)}
                    </td>
                    <td><CategoryBadge category={item.category} /></td>
                    <td style={{ fontSize: 13, fontWeight: 500 }}>{item.title}</td>
                    <td style={{ fontSize: 12 }}>
                      {item.mac && <div className="mono">{item.mac}</div>}
                      {item.hostname && (
                        <div style={{ color: "var(--text-dim)", fontSize: 11 }}>{item.hostname}</div>
                      )}
                      {!item.mac && !item.hostname && "—"}
                    </td>
                    <td style={{ fontSize: 12, color: "var(--text-muted)" }}>
                      {item.source ?? "—"}
                    </td>
                    <td
                      style={{
                        fontSize: 12,
                        color: "var(--text-muted)",
                        maxWidth: 320,
                        overflow: "hidden",
                        textOverflow: "ellipsis",
                      }}
                      title={item.detail ?? ""}
                    >
                      {item.detail ?? "—"}
                    </td>
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
