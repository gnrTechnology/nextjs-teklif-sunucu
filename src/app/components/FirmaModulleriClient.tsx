"use client";

import { useState, useTransition, useRef, useEffect } from "react";
import { useRouter } from "next/navigation";
import type { FirmAutoModuleRecord, FirmAutoStartModule } from "@/lib/types";

/* ── Kategori renkleri (ModullerClient ile uyumlu) ───── */
const CAT_COLOR: Record<string, string> = {
  lisans: "#10b981", sistem: "#f59e0b", dosya: "#8b5cf6", internet: "#3b82f6",
  excel: "#10b981", powershell: "#f59e0b", registry: "#f97316", guvenlik: "#ef4444",
  bildirim: "#3b82f6", genel: "#6b7280", zamanlanmis: "#a78bfa", uzman: "#ec4899",
  donanim: "#06b6d4",
};

/* ── Aranabilir modül seçici ─────────────────────────── */
function SearchableModuleSelect({
  value,
  onChange,
  modules,
  moduleCategories,
}: {
  value: string;
  onChange: (v: string) => void;
  modules: string[];
  moduleCategories?: Record<string, string>;
}) {
  const [search, setSearch] = useState("");
  const [open, setOpen] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);
  const listRef  = useRef<HTMLDivElement>(null);

  const displayText = !open && value ? value : search;
  const filtered = modules.filter((m) =>
    m.toLowerCase().includes(search.toLowerCase())
  );

  useEffect(() => {
    if (!open) setSearch("");
  }, [open]);

  function select(m: string) {
    onChange(m);
    setOpen(false);
    setSearch("");
  }

  return (
    <div style={{ position: "relative" }}>
      <div style={{ position: "relative" }}>
        <input
          ref={inputRef}
          className="form-input mono"
          value={displayText}
          placeholder="Modül adı yazarak ara…"
          onFocus={() => setOpen(true)}
          onBlur={(e) => {
            if (!listRef.current?.contains(e.relatedTarget as Node)) {
              setOpen(false);
            }
          }}
          onChange={(e) => { setSearch(e.target.value); onChange(""); }}
          autoComplete="off"
        />
        {value && (
          <button
            onMouseDown={() => { onChange(""); setSearch(""); inputRef.current?.focus(); }}
            style={{
              position: "absolute", right: 8, top: "50%", transform: "translateY(-50%)",
              background: "none", border: "none", cursor: "pointer",
              color: "var(--text-dim)", fontSize: 14, padding: "2px 4px",
            }}
          >×</button>
        )}
      </div>

      {open && (
        <div
          ref={listRef}
          tabIndex={-1}
          style={{
            position: "absolute", top: "calc(100% + 4px)", left: 0, right: 0, zIndex: 500,
            background: "var(--bg-card)", border: "1px solid var(--border)",
            borderRadius: "var(--radius-sm)", maxHeight: 280, overflowY: "auto",
            boxShadow: "0 8px 32px rgba(0,0,0,0.35)",
          }}
        >
          {filtered.length === 0 ? (
            <div style={{ padding: "12px 14px", fontSize: 12, color: "var(--text-dim)" }}>
              Sonuç bulunamadı
            </div>
          ) : (
            filtered.map((m) => {
              const cat = moduleCategories?.[m];
              const color = CAT_COLOR[cat ?? ""] ?? "#6b7280";
              return (
                <div
                  key={m}
                  onMouseDown={() => select(m)}
                  style={{
                    padding: "9px 14px", cursor: "pointer", display: "flex",
                    alignItems: "center", gap: 10,
                    background: m === value ? "var(--accent-dim)" : "transparent",
                    borderBottom: "1px solid var(--border)",
                    transition: "background 0.1s",
                  }}
                  onMouseEnter={(e) => { (e.currentTarget as HTMLDivElement).style.background = "var(--bg)"; }}
                  onMouseLeave={(e) => {
                    (e.currentTarget as HTMLDivElement).style.background = m === value ? "var(--accent-dim)" : "transparent";
                  }}
                >
                  <span style={{
                    width: 8, height: 8, borderRadius: "50%",
                    background: color, flexShrink: 0,
                  }} />
                  <span style={{ fontFamily: "var(--font-geist-mono, monospace)", fontSize: 13, flex: 1 }}>
                    {m}
                  </span>
                  {cat && (
                    <span style={{ fontSize: 10, color, opacity: 0.8 }}>{cat}</span>
                  )}
                </div>
              );
            })
          )}
          <div
            onMouseDown={() => { onChange("__custom__"); setOpen(false); }}
            style={{
              padding: "9px 14px", cursor: "pointer", fontSize: 12,
              color: "var(--text-muted)", borderTop: "1px solid var(--border)",
              display: "flex", alignItems: "center", gap: 8,
            }}
          >
            ✏️ Özel modül adı gir
          </div>
        </div>
      )}
    </div>
  );
}

/* ── Küçük yardımcılar ─────────────────────────────── */
function Badge({ ok, labels = ["Aktif", "Pasif"] }: { ok: boolean; labels?: [string, string] }) {
  return (
    <span className={`badge ${ok ? "badge-green" : "badge-red"}`}>
      <span className="badge-dot" />{ok ? labels[0] : labels[1]}
    </span>
  );
}

async function apiFetch(url: string, method: string, body?: unknown) {
  const r = await fetch(url, {
    method,
    headers: { "Content-Type": "application/json" },
    body: body ? JSON.stringify(body) : undefined,
  });
  return r.json();
}

/* ── Yeni firma formu ──────────────────────────────── */
function AddFirmaPanel({
  onClose,
  onSaved,
}: {
  onClose: () => void;
  onSaved: () => void;
}) {
  const [firmaAdi, setFirmaAdi] = useState("");
  const [description, setDescription] = useState("");
  const [saving, setSaving] = useState(false);

  async function save() {
    if (!firmaAdi.trim()) { alert("Firma adı zorunludur."); return; }
    setSaving(true);
    const r = await apiFetch("/api/firm-modules", "POST", { firmaAdi: firmaAdi.trim(), description });
    setSaving(false);
    if (r.success) { onSaved(); onClose(); }
    else alert(r.error ?? "Hata oluştu.");
  }

  return (
    <Overlay onClose={onClose} title="Yeni Firma Ekle">
      <div style={{ display: "flex", flexDirection: "column", gap: 16, padding: "22px" }}>
        <label style={LS}>
          <span>Firma Adı *</span>
          <input className="form-input mono" value={firmaAdi} onChange={(e) => setFirmaAdi(e.target.value)}
            placeholder="EPRON  (veya * = tüm firmalar)" />
        </label>
        <label style={LS}>
          <span>Açıklama</span>
          <input className="form-input" value={description} onChange={(e) => setDescription(e.target.value)}
            placeholder="Bu firma için kısa açıklama" />
        </label>
      </div>
      <PanelFooter onClose={onClose} onSave={save} saving={saving} saveLabel="Ekle" />
    </Overlay>
  );
}

/* ── Modül ekle/düzenle formu ──────────────────────── */
function ModuleEditPanel({
  firmaAdi,
  editing,
  availableModules,
  moduleCategories,
  maxOrder,
  onClose,
  onSaved,
}: {
  firmaAdi: string;
  editing: FirmAutoStartModule | null;
  availableModules: string[];
  moduleCategories: Record<string, string>;
  maxOrder: number;
  onClose: () => void;
  onSaved: () => void;
}) {
  const [methodName, setMethodName] = useState(editing?.methodName ?? "");
  const [customName, setCustomName]  = useState("");
  const [order, setOrder]            = useState(editing?.order ?? maxOrder + 1);
  const [delay, setDelay]            = useState(editing?.delaySeconds ?? 0);
  const [runOnce, setRunOnce]        = useState(editing?.runOnce ?? false);
  const [saving, setSaving]          = useState(false);

  const finalName = methodName === "__custom__" ? customName.trim() : methodName.trim();

  async function save() {
    if (!finalName) { alert("Modül adı zorunludur."); return; }
    setSaving(true);
    const body = editing
      ? { updateModule: { methodName: finalName, order, delaySeconds: delay, runOnce } }
      : { addModule:    { methodName: finalName, order, delaySeconds: delay, runOnce } };
    const r = await apiFetch(`/api/firm-modules/${encodeURIComponent(firmaAdi)}`, "PATCH", body);
    setSaving(false);
    if (r.success) { onSaved(); onClose(); }
    else alert(r.error ?? "Hata oluştu.");
  }

  return (
    <Overlay onClose={onClose} title={editing ? "Modül Düzenle" : "Modül Ekle"}>
      <div style={{ display: "flex", flexDirection: "column", gap: 16, padding: "22px" }}>
        <label style={LS}>
          <span>Modül Adı *</span>
          {editing ? (
            <input className="form-input mono" value={methodName} disabled style={{ opacity: 0.6 }} />
          ) : (
            <SearchableModuleSelect
              value={methodName}
              onChange={setMethodName}
              modules={availableModules}
              moduleCategories={moduleCategories}
            />
          )}
          {(!editing && methodName === "__custom__") && (
            <input
              className="form-input mono"
              placeholder="MethodName girin"
              style={{ marginTop: 6 }}
              value={customName}
              onChange={(e) => setCustomName(e.target.value)}
            />
          )}
          {finalName && finalName !== "__custom__" && (
            <div style={{
              marginTop: 6, padding: "6px 10px", borderRadius: "var(--radius-sm)",
              background: "var(--accent-dim)", fontSize: 12,
              color: "var(--accent)", fontFamily: "var(--font-geist-mono, monospace)",
            }}>
              ✓ {finalName}
              {moduleCategories[finalName] && (
                <span style={{ marginLeft: 8, opacity: 0.7 }}>— {moduleCategories[finalName]}</span>
              )}
            </div>
          )}
        </label>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
          <label style={LS}>
            <span>Sıra (order)</span>
            <input className="form-input" type="number" min={1} value={order}
              onChange={(e) => setOrder(Number(e.target.value))} />
          </label>
          <label style={LS}>
            <span>Gecikme (saniye)</span>
            <input className="form-input" type="number" min={0} value={delay}
              onChange={(e) => setDelay(Number(e.target.value))} />
          </label>
        </div>
        <label style={{ ...LS, flexDirection: "row", alignItems: "center", gap: 10, cursor: "pointer" }}>
          <input
            type="checkbox"
            checked={runOnce}
            onChange={(e) => setRunOnce(e.target.checked)}
            style={{ width: 16, height: 16, accentColor: "var(--accent)" }}
          />
          <span>
            Tek seferlik çalıştır (runOnce)
            <span style={{ display: "block", fontSize: 11, fontWeight: 400, color: "var(--text-dim)", marginTop: 2 }}>
              Heartbeat, komut kuyruğu gibi kurulum modülleri için — istemci bir kez çalıştırır, sonra atlar
            </span>
          </span>
        </label>
      </div>
      <PanelFooter onClose={onClose} onSave={save} saving={saving}
        saveLabel={editing ? "Güncelle" : "Ekle"} />
    </Overlay>
  );
}

/* ── Ana bileşen ───────────────────────────────────── */
export default function FirmaModulleriClient({
  initial,
  allModuleNames,
  moduleCategories = {},
}: {
  initial: FirmAutoModuleRecord[];
  allModuleNames: string[];
  moduleCategories?: Record<string, string>;
}) {
  const router = useRouter();
  const [items, setItems] = useState<FirmAutoModuleRecord[]>(initial);
  const [, startTransition] = useTransition();

  const [addFirmaOpen, setAddFirmaOpen] = useState(false);
  const [modulePanel, setModulePanel] = useState<{
    firmaAdi: string;
    editing: FirmAutoStartModule | null;
  } | null>(null);
  const [busy, setBusy] = useState<string | null>(null);

  function refresh() {
    startTransition(() => router.refresh());
    fetch("/api/firm-modules").then((r) => r.json()).then((j) => {
      if (j.success) setItems(j.data);
    });
  }

  async function patchFirma(firmaAdi: string, body: Record<string, unknown>) {
    setBusy(firmaAdi);
    await apiFetch(`/api/firm-modules/${encodeURIComponent(firmaAdi)}`, "PATCH", body);
    setBusy(null);
    refresh();
  }

  async function deleteFirma(firma: FirmAutoModuleRecord) {
    if (!confirm(`"${firma.firmaAdi}" firma kaydı silinecek. Emin misiniz?`)) return;
    setBusy(firma.firmaAdi);
    await apiFetch(`/api/firm-modules/${encodeURIComponent(firma.firmaAdi)}`, "DELETE");
    setBusy(null);
    refresh();
  }

  async function removeModule(firmaAdi: string, methodName: string) {
    if (!confirm(`"${methodName}" modülü çıkarılacak?`)) return;
    await patchFirma(firmaAdi, { removeModule: methodName });
  }

  function moveOrder(firm: FirmAutoModuleRecord, mod: FirmAutoStartModule, dir: -1 | 1) {
    const mods = [...firm.onExcelOpen.modules].sort((a, b) => a.order - b.order);
    const i = mods.findIndex((m) => m.methodName === mod.methodName);
    const j = i + dir;
    if (j < 0 || j >= mods.length) return;
    const reorder = mods.map((m, idx) => {
      if (idx === i) return { methodName: m.methodName, order: mods[j].order };
      if (idx === j) return { methodName: m.methodName, order: mods[i].order };
      return { methodName: m.methodName, order: m.order };
    });
    patchFirma(firm.firmaAdi, { reorderModules: reorder });
  }

  const curPanel = modulePanel
    ? items.find((f) => f.firmaAdi === modulePanel.firmaAdi)
    : null;

  return (
    <>
      {/* Firma ekle paneli */}
      {addFirmaOpen && (
        <AddFirmaPanel onClose={() => setAddFirmaOpen(false)} onSaved={refresh} />
      )}

      {/* Modül ekle/düzenle paneli */}
      {modulePanel && curPanel && (
        <ModuleEditPanel
          firmaAdi={modulePanel.firmaAdi}
          editing={modulePanel.editing}
          availableModules={allModuleNames}
          moduleCategories={moduleCategories}
          maxOrder={Math.max(0, ...curPanel.onExcelOpen.modules.map((m) => m.order))}
          onClose={() => setModulePanel(null)}
          onSaved={refresh}
        />
      )}

      <div className="page-wrap">
        <div className="page-header">
          <div>
            <div className="page-title">Firma Otomatik Modülleri</div>
            <div className="page-sub">Excel açılışında firma bazlı çalışan modüller</div>
          </div>
          <button className="btn btn-primary" onClick={() => setAddFirmaOpen(true)}>
            + Yeni Firma
          </button>
        </div>

        {items.length === 0 && (
          <div className="card">
            <div className="empty-state">
              <div className="empty-state-icon">📭</div>
              <div>Henüz firma tanımlanmamış.</div>
            </div>
          </div>
        )}

        {items.map((item) => {
          const isGlobal = item.firmaAdi === "*";
          const mods = [...(item.onExcelOpen.modules ?? [])].sort((a, b) => a.order - b.order);
          const isBusy = busy === item.firmaAdi;

          return (
            <div key={item.firmaAdi} className="card" style={{ marginBottom: 16 }}>
              {/* Firma başlık */}
              <div className="card-header">
                <span className="card-title" style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                  <span style={{ fontSize: 16 }}>{isGlobal ? "🌐" : "🏢"}</span>
                  <span className="mono" style={{ fontWeight: 700 }}>
                    {isGlobal ? "Tüm Firmalar (*)" : item.firmaAdi}
                  </span>
                  <Badge ok={item.enabled !== false} />
                  {item.description && (
                    <span style={{ fontSize: 12, color: "var(--text-muted)", fontWeight: 400 }}>
                      {item.description}
                    </span>
                  )}
                </span>

                {/* Firma düzenleme butonları */}
                <div style={{ display: "flex", gap: 8, flexShrink: 0 }}>
                  <button
                    className={`btn ${item.enabled !== false ? "btn-ghost" : "btn-accent"}`}
                    style={{ fontSize: 12, padding: "4px 10px" }}
                    disabled={isBusy}
                    onClick={() => patchFirma(item.firmaAdi, { enabled: !(item.enabled !== false) })}
                    title={item.enabled !== false ? "Devre dışı bırak" : "Aktif et"}
                  >
                    {item.enabled !== false ? "Devre Dışı Bırak" : "Aktif Et"}
                  </button>
                  {!isGlobal && (
                    <button
                      className="btn btn-danger"
                      style={{ fontSize: 12, padding: "4px 10px" }}
                      disabled={isBusy}
                      onClick={() => deleteFirma(item)}
                    >
                      {isBusy ? "…" : "Sil"}
                    </button>
                  )}
                </div>
              </div>

              {/* Excel Açılış Modülleri */}
              <div style={{ padding: "14px 18px" }}>
                <div style={{
                  display: "flex", alignItems: "center", justifyContent: "space-between",
                  marginBottom: 12,
                }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                    <span style={{ fontSize: 11, fontWeight: 600, textTransform: "uppercase",
                      letterSpacing: "0.06em", color: "var(--text-muted)" }}>
                      Excel Açılış Modülleri
                    </span>
                    <Badge ok={item.onExcelOpen.enabled !== false} labels={["Açık", "Kapalı"]} />
                  </div>
                  <div style={{ display: "flex", gap: 8 }}>
                    <button
                      className="btn btn-ghost"
                      style={{ fontSize: 11, padding: "3px 10px" }}
                      onClick={() => patchFirma(item.firmaAdi, {
                        onExcelOpenEnabled: !(item.onExcelOpen.enabled !== false),
                      })}
                    >
                      {item.onExcelOpen.enabled !== false ? "Kapat" : "Aç"}
                    </button>
                    <button
                      className="btn btn-primary"
                      style={{ fontSize: 11, padding: "3px 10px" }}
                      onClick={() => setModulePanel({ firmaAdi: item.firmaAdi, editing: null })}
                    >
                      + Modül Ekle
                    </button>
                  </div>
                </div>

                {mods.length === 0 ? (
                  <div style={{ color: "var(--text-dim)", fontSize: 13, padding: "8px 0" }}>
                    Henüz modül eklenmemiş.
                  </div>
                ) : (
                  <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                    {mods.map((mod, i) => (
                      <div key={mod.methodName} style={{
                        display: "flex", alignItems: "center", gap: 10,
                        background: "var(--bg)", border: "1px solid var(--border)",
                        borderRadius: "var(--radius-sm)", padding: "9px 14px",
                      }}>
                        {/* Sıra numarası */}
                        <span style={{
                          width: 24, height: 24, borderRadius: "50%",
                          background: "var(--accent-dim)", color: "var(--accent)",
                          display: "flex", alignItems: "center", justifyContent: "center",
                          fontSize: 11, fontWeight: 700, flexShrink: 0,
                        }}>{mod.order}</span>

                        {/* Modül adı */}
                        <span className="mono" style={{ fontSize: 13, fontWeight: 600, flex: 1 }}>
                          {mod.methodName}
                        </span>

                        {/* Gecikme */}
                        {(mod.delaySeconds ?? 0) > 0 && (
                          <span className="badge badge-yellow">⏱ {mod.delaySeconds}s</span>
                        )}
                        {mod.runOnce && (
                          <span className="badge badge-green" title="Excel açılışında yalnızca bir kez çalışır">1×</span>
                        )}

                        {/* Yukarı / Aşağı */}
                        <div style={{ display: "flex", gap: 2 }}>
                          <button
                            onClick={() => moveOrder(item, mod, -1)}
                            disabled={i === 0}
                            style={arrowBtn}
                            title="Yukarı taşı"
                          >▲</button>
                          <button
                            onClick={() => moveOrder(item, mod, 1)}
                            disabled={i === mods.length - 1}
                            style={arrowBtn}
                            title="Aşağı taşı"
                          >▼</button>
                        </div>

                        {/* Düzenle */}
                        <button
                          className="btn btn-ghost"
                          style={{ fontSize: 11, padding: "3px 9px" }}
                          onClick={() => setModulePanel({ firmaAdi: item.firmaAdi, editing: mod })}
                        >Düzenle</button>

                        {/* Çıkar */}
                        <button
                          className="btn btn-danger"
                          style={{ fontSize: 11, padding: "3px 9px" }}
                          onClick={() => removeModule(item.firmaAdi, mod.methodName)}
                        >Çıkar</button>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
          );
        })}
      </div>
    </>
  );
}

/* ── Paylaşılan mini bileşenler ────────────────────── */
function Overlay({
  onClose,
  title,
  children,
}: {
  onClose: () => void;
  title: string;
  children: React.ReactNode;
}) {
  return (
    <div
      style={{
        position: "fixed", inset: 0, zIndex: 1000,
        background: "rgba(0,0,0,0.55)", backdropFilter: "blur(3px)",
        display: "flex", alignItems: "flex-start", justifyContent: "flex-end",
      }}
      onClick={(e) => { if (e.target === e.currentTarget) onClose(); }}
    >
      <div style={{
        width: "min(520px, 100vw)", minHeight: "100dvh",
        background: "var(--bg-card)", borderLeft: "1px solid var(--border)",
        display: "flex", flexDirection: "column",
      }}>
        <div style={{
          padding: "18px 22px", borderBottom: "1px solid var(--border)",
          display: "flex", alignItems: "center", justifyContent: "space-between",
        }}>
          <div style={{ fontWeight: 700, fontSize: 15 }}>{title}</div>
          <button className="btn btn-ghost" onClick={onClose} style={{ fontSize: 18, padding: "4px 10px" }}>✕</button>
        </div>
        <div style={{ flex: 1 }}>{children}</div>
      </div>
    </div>
  );
}

function PanelFooter({
  onClose, onSave, saving, saveLabel,
}: {
  onClose: () => void; onSave: () => void; saving: boolean; saveLabel: string;
}) {
  return (
    <div style={{
      padding: "14px 22px", borderTop: "1px solid var(--border)",
      display: "flex", gap: 10, justifyContent: "flex-end",
    }}>
      <button className="btn btn-ghost" onClick={onClose}>İptal</button>
      <button className="btn btn-primary" onClick={onSave} disabled={saving}>
        {saving ? "Kaydediliyor…" : saveLabel}
      </button>
    </div>
  );
}

const LS: React.CSSProperties = {
  display: "flex", flexDirection: "column", gap: 6,
  fontSize: 13, fontWeight: 500, color: "var(--text-muted)",
};

const arrowBtn: React.CSSProperties = {
  background: "none", border: "1px solid var(--border)",
  borderRadius: 4, cursor: "pointer", padding: "2px 7px",
  fontSize: 10, color: "var(--text-dim)",
  lineHeight: 1,
};
