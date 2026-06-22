"use client";

import { useState, useTransition } from "react";
import { useRouter } from "next/navigation";
import { formatTR } from "@/lib/date-utils";
import type { ModuleRecord } from "@/lib/types";

const CATEGORIES = [
  "genel", "lisans", "sistem", "dosya",
  "internet", "excel", "powershell", "registry", "guvenlik", "bildirim",
];

const EMPTY_FORM = {
  methodName: "",
  description: "",
  category: "genel",
  code: "",
  active: true,
};

type FormState = typeof EMPTY_FORM;

function CategoryBadge({ cat }: { cat?: string }) {
  const colors: Record<string, string> = {
    lisans: "badge-green", sistem: "badge-yellow",
    dosya: "#8b5cf6", internet: "badge-accent",
    excel: "badge-green", powershell: "badge-yellow",
    registry: "#f97316", guvenlik: "badge-red",
    bildirim: "badge-accent", genel: "",
  };
  const cls = colors[cat ?? ""] ?? "";
  return (
    <span className={`badge ${cls}`} style={{ fontSize: 10 }}>
      {cat ?? "genel"}
    </span>
  );
}

export default function ModullerClient({ initial }: { initial: ModuleRecord[] }) {
  const router = useRouter();
  const [modules, setModules] = useState<ModuleRecord[]>(initial);
  const [isPending, startTransition] = useTransition();

  const [panelOpen, setPanelOpen] = useState(false);
  const [editing, setEditing] = useState<ModuleRecord | null>(null);
  const [form, setForm] = useState<FormState>(EMPTY_FORM);
  const [saving, setSaving] = useState(false);
  const [deleting, setDeleting] = useState<number | null>(null);
  const [expanded, setExpanded] = useState<Set<string>>(new Set());
  const [filterCat, setFilterCat] = useState<string>("tümü");

  function refresh() {
    startTransition(() => router.refresh());
    fetch("/api/modules")
      .then((r) => r.json())
      .then((j) => { if (j.success) setModules(j.data); });
  }

  function openNew() {
    setEditing(null);
    setForm(EMPTY_FORM);
    setPanelOpen(true);
  }

  function openEdit(m: ModuleRecord) {
    setEditing(m);
    setForm({
      methodName: m.methodName,
      description: m.description ?? "",
      category: m.category ?? "genel",
      code: m.code,
      active: m.active ?? true,
    });
    setPanelOpen(true);
  }

  function closePanel() {
    setPanelOpen(false);
    setEditing(null);
    setForm(EMPTY_FORM);
  }

  async function handleSave() {
    if (!form.methodName.trim()) { alert("methodName zorunludur."); return; }
    if (!form.code.trim()) { alert("Kod zorunludur."); return; }
    setSaving(true);
    try {
      if (editing?.id) {
        await fetch(`/api/modules/${editing.id}`, {
          method: "PATCH",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(form),
        });
      } else {
        await fetch("/api/modules", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(form),
        });
      }
      closePanel();
      refresh();
    } finally {
      setSaving(false);
    }
  }

  async function handleDelete(m: ModuleRecord) {
    if (!m.id) return;
    if (!confirm(`"${m.methodName}" modülü silinecek. Emin misiniz?`)) return;
    setDeleting(m.id);
    try {
      await fetch(`/api/modules/${m.id}`, { method: "DELETE" });
      refresh();
    } finally {
      setDeleting(null);
    }
  }

  async function handleToggle(m: ModuleRecord) {
    if (!m.id) return;
    await fetch(`/api/modules/${m.id}`, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ active: !m.active }),
    });
    refresh();
  }

  function toggleExpand(name: string) {
    setExpanded((prev) => {
      const next = new Set(prev);
      next.has(name) ? next.delete(name) : next.add(name);
      return next;
    });
  }

  const cats = ["tümü", ...Array.from(new Set(modules.map((m) => m.category ?? "genel")))];
  const filtered = filterCat === "tümü" ? modules : modules.filter((m) => (m.category ?? "genel") === filterCat);

  return (
    <>
      {/* ── Panel overlay ── */}
      {panelOpen && (
        <div
          style={{
            position: "fixed", inset: 0, zIndex: 1000,
            background: "rgba(0,0,0,0.55)", backdropFilter: "blur(3px)",
            display: "flex", alignItems: "flex-start", justifyContent: "flex-end",
          }}
          onClick={(e) => { if (e.target === e.currentTarget) closePanel(); }}
        >
          <div style={{
            width: "min(680px, 100vw)", height: "100dvh",
            background: "var(--bg-card)", borderLeft: "1px solid var(--border)",
            display: "flex", flexDirection: "column", overflow: "hidden",
          }}>
            {/* Panel header */}
            <div style={{
              padding: "18px 22px", borderBottom: "1px solid var(--border)",
              display: "flex", alignItems: "center", justifyContent: "space-between",
            }}>
              <div>
                <div style={{ fontWeight: 700, fontSize: 15 }}>
                  {editing ? "Modül Düzenle" : "Yeni Modül"}
                </div>
                {editing && (
                  <div style={{ fontSize: 12, color: "var(--text-muted)", marginTop: 2 }}>
                    Son güncelleme: {formatTR(editing.updatedAt)}
                  </div>
                )}
              </div>
              <button className="btn btn-ghost" onClick={closePanel} style={{ fontSize: 18, padding: "4px 10px" }}>✕</button>
            </div>

            {/* Panel body */}
            <div style={{ flex: 1, overflowY: "auto", padding: "22px" }}>
              <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>

                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                  <label style={labelStyle}>
                    <span>methodName *</span>
                    <input
                      className="form-input mono"
                      value={form.methodName}
                      disabled={!!editing}
                      onChange={(e) => setForm({ ...form, methodName: e.target.value })}
                      placeholder="GetComputerName"
                      style={{ fontFamily: "var(--font-geist-mono, monospace)", opacity: editing ? 0.6 : 1 }}
                    />
                  </label>
                  <label style={labelStyle}>
                    <span>Kategori</span>
                    <select
                      className="form-input"
                      value={form.category}
                      onChange={(e) => setForm({ ...form, category: e.target.value })}
                    >
                      {CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}
                    </select>
                  </label>
                </div>

                <label style={labelStyle}>
                  <span>Açıklama</span>
                  <input
                    className="form-input"
                    value={form.description}
                    onChange={(e) => setForm({ ...form, description: e.target.value })}
                    placeholder="Modül işlevini kısaca açıklayın"
                  />
                </label>

                <label style={{ ...labelStyle, display: "flex", flexDirection: "row", alignItems: "center", gap: 12, cursor: "pointer" }}>
                  <div style={{
                    width: 40, height: 22, borderRadius: 11,
                    background: form.active ? "var(--green)" : "var(--border-strong)",
                    position: "relative", transition: "background 0.2s", cursor: "pointer",
                    flexShrink: 0,
                  }} onClick={() => setForm({ ...form, active: !form.active })}>
                    <div style={{
                      position: "absolute", top: 3, left: form.active ? 20 : 3,
                      width: 16, height: 16, borderRadius: "50%",
                      background: "#fff", transition: "left 0.2s",
                    }} />
                  </div>
                  <span style={{ fontSize: 13 }}>{form.active ? "Aktif" : "Pasif"}</span>
                </label>

                <label style={labelStyle}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <span>VBA Kodu *</span>
                    <span style={{ fontSize: 11, color: "var(--text-muted)" }}>{form.code.length} karakter</span>
                  </div>
                  <textarea
                    className="form-input mono"
                    value={form.code}
                    onChange={(e) => setForm({ ...form, code: e.target.value })}
                    placeholder={"Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object\n    ' ...\n    Set DynamicFunc = Nothing\nEnd Function"}
                    style={{
                      minHeight: 340, resize: "vertical",
                      fontFamily: "var(--font-geist-mono, monospace)",
                      fontSize: 12, lineHeight: 1.7,
                    }}
                  />
                </label>
              </div>
            </div>

            {/* Panel footer */}
            <div style={{
              padding: "14px 22px", borderTop: "1px solid var(--border)",
              display: "flex", gap: 10, justifyContent: "flex-end",
            }}>
              <button className="btn btn-ghost" onClick={closePanel}>İptal</button>
              <button className="btn btn-primary" onClick={handleSave} disabled={saving}>
                {saving ? "Kaydediliyor…" : editing ? "Güncelle" : "Oluştur"}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ── Ana sayfa içeriği ── */}
      <div className="page-wrap">
        <div className="page-header">
          <div>
            <div className="page-title">Modüller</div>
            <div className="page-sub">RunRemoteCode ile VBA'ya gönderilen modüller — Neon DB</div>
          </div>
          <button className="btn btn-primary" onClick={openNew}>+ Yeni Modül</button>
        </div>

        {/* Kategori filtresi */}
        {cats.length > 2 && (
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 16 }}>
            {cats.map((c) => (
              <button
                key={c}
                onClick={() => setFilterCat(c)}
                style={{
                  padding: "4px 12px", borderRadius: 20, fontSize: 12, cursor: "pointer", border: "1px solid var(--border)",
                  background: filterCat === c ? "var(--accent)" : "var(--bg-card)",
                  color: filterCat === c ? "#fff" : "var(--text-muted)",
                  transition: "all 0.15s",
                }}
              >
                {c} {c !== "tümü" && <span style={{ opacity: 0.6 }}>({modules.filter((m) => (m.category ?? "genel") === c).length})</span>}
              </button>
            ))}
          </div>
        )}

        <div className="card" style={{ padding: 0 }}>
          <div className="card-header">
            <span className="card-title">
              Modüller <span className="card-count">{filtered.length}</span>
            </span>
          </div>

          {filtered.length === 0 ? (
            <div className="empty-state">
              <div className="empty-state-icon">📭</div>
              <div>Modül bulunamadı.</div>
            </div>
          ) : (
            <div>
              {filtered.map((m) => {
                const isOpen = expanded.has(m.methodName);
                return (
                  <div key={m.methodName} style={{ borderBottom: "1px solid var(--border)" }}>
                    {/* Row */}
                    <div style={{
                      padding: "13px 18px", display: "flex",
                      alignItems: "center", gap: 12,
                    }}>
                      {/* Expand toggle */}
                      <button
                        onClick={() => toggleExpand(m.methodName)}
                        style={{
                          background: "none", border: "none", cursor: "pointer",
                          color: "var(--text-dim)", fontSize: 11, padding: 4,
                          transform: isOpen ? "rotate(90deg)" : "none",
                          transition: "transform 0.15s",
                        }}
                      >▶</button>

                      {/* Name + meta */}
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                          <span className="mono" style={{ fontSize: 13, fontWeight: 600 }}>{m.methodName}</span>
                          <CategoryBadge cat={m.category} />
                        </div>
                        {m.description && (
                          <div style={{ fontSize: 12, color: "var(--text-muted)", marginTop: 2 }}>
                            {m.description}
                          </div>
                        )}
                      </div>

                      {/* Updated */}
                      <div style={{ fontSize: 11, color: "var(--text-dim)", whiteSpace: "nowrap", display: "none" }}
                        className="hide-mobile">
                        {formatTR(m.updatedAt)}
                      </div>

                      {/* Active toggle */}
                      <div
                        onClick={() => handleToggle(m)}
                        style={{
                          width: 34, height: 19, borderRadius: 10,
                          background: m.active !== false ? "var(--green)" : "var(--border-strong)",
                          position: "relative", cursor: "pointer", flexShrink: 0,
                        }}
                        title={m.active !== false ? "Pasif yap" : "Aktif yap"}
                      >
                        <div style={{
                          position: "absolute", top: 2.5,
                          left: m.active !== false ? 17 : 2.5,
                          width: 14, height: 14, borderRadius: "50%",
                          background: "#fff", transition: "left 0.15s",
                        }} />
                      </div>

                      {/* Actions */}
                      <div style={{ display: "flex", gap: 6 }}>
                        <button
                          className="btn btn-ghost"
                          style={{ padding: "4px 10px", fontSize: 12 }}
                          onClick={() => openEdit(m)}
                        >Düzenle</button>
                        <button
                          className="btn btn-danger"
                          style={{ padding: "4px 10px", fontSize: 12 }}
                          onClick={() => handleDelete(m)}
                          disabled={deleting === m.id}
                        >{deleting === m.id ? "…" : "Sil"}</button>
                      </div>
                    </div>

                    {/* Code accordion */}
                    {isOpen && (
                      <div style={{
                        background: "var(--bg)", borderTop: "1px solid var(--border)",
                        padding: "16px 18px",
                      }}>
                        <div style={{
                          fontSize: 11, fontWeight: 600, textTransform: "uppercase",
                          letterSpacing: "0.06em", color: "var(--text-muted)", marginBottom: 8,
                        }}>
                          VBA Kaynak Kodu — {m.code.split("\n").length} satır
                        </div>
                        <pre style={{
                          fontFamily: "var(--font-geist-mono, monospace)",
                          fontSize: 12, lineHeight: 1.7,
                          color: "var(--code-text)", background: "var(--code-bg)",
                          padding: "14px 16px", borderRadius: "var(--radius-sm)",
                          overflowX: "auto", maxHeight: 400, overflowY: "auto",
                          whiteSpace: "pre-wrap", wordBreak: "break-word",
                        }}>
                          {m.code}
                        </pre>
                        <div style={{ marginTop: 8, fontSize: 11, color: "var(--text-dim)" }}>
                          Oluşturuldu: {formatTR(m.createdAt)} · Güncellendi: {formatTR(m.updatedAt)}
                        </div>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}
        </div>
      </div>
    </>
  );
}

const labelStyle: React.CSSProperties = {
  display: "flex", flexDirection: "column", gap: 6,
  fontSize: 13, fontWeight: 500, color: "var(--text-muted)",
};
