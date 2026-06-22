"use client";

import { useState, useEffect } from "react";
import { formatTR } from "@/lib/date-utils";
import type { DeviceSnapshot } from "@/lib/db";

/* Kaç ms önce toplandı → "3 dakika önce" */
function ago(iso: string): string {
  const ms = Date.now() - new Date(iso).getTime();
  const min = Math.floor(ms / 60000);
  if (min < 2) return "Az önce";
  if (min < 60) return `${min} dk önce`;
  const h = Math.floor(min / 60);
  if (h < 24) return `${h} sa önce`;
  return `${Math.floor(h / 24)} gün önce`;
}

/* Veri kategorileri — anahtara göre grupla */
const SECTION_ORDER = [
  { key: "bilgisayar",   icon: "🖥", label: "Bilgisayar",       keys: ["computerName","windowsVersion","windowsActivation","systemUptime","timeZone","locale","domainName","loggedInUser"] },
  { key: "donanim",      icon: "🔧", label: "Donanım",           keys: ["cpu","ram","gpu","screenResolution","bios","motherboard"] },
  { key: "disk",         icon: "💾", label: "Disk",              keys: ["disks","diskInfo"] },
  { key: "ag",           icon: "🌐", label: "Ağ",                keys: ["mac","ip","publicIp","networkAdapters","wifiProfiles","bitlockerStatus"] },
  { key: "sistem",       icon: "⚙️", label: "Sistem",            keys: ["battery","printers","usbDevices","audioDevices","runningProcesses","installedSoftware"] },
];

const KEY_LABELS: Record<string, string> = {
  computerName: "Bilgisayar Adı", windowsVersion: "Windows Sürümü",
  windowsActivation: "Aktivasyon", systemUptime: "Çalışma Süresi",
  timeZone: "Saat Dilimi", locale: "Yerel Ayarlar", domainName: "Domain",
  loggedInUser: "Kullanıcı", cpu: "İşlemci", ram: "Bellek",
  gpu: "Ekran Kartı", screenResolution: "Ekran Çözünürlüğü",
  bios: "BIOS", motherboard: "Anakart", disks: "Diskler", diskInfo: "Disk Bilgisi",
  mac: "MAC Adresi", ip: "IP Adresi", publicIp: "Dış IP",
  networkAdapters: "Ağ Adaptörleri", wifiProfiles: "Wi-Fi Profilleri",
  bitlockerStatus: "BitLocker", battery: "Pil", printers: "Yazıcılar",
  usbDevices: "USB Aygıtlar", audioDevices: "Ses Aygıtları",
  runningProcesses: "Çalışan İşlemler", installedSoftware: "Yazılımlar",
};

function DataValue({ v }: { v: unknown }) {
  if (v === null || v === undefined) return <span style={{ color: "var(--text-dim)" }}>—</span>;
  if (Array.isArray(v)) {
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 2, maxHeight: 160, overflowY: "auto" }}>
        {v.map((item, i) => (
          <div key={i} style={{ fontSize: 11, color: "var(--text-muted)", padding: "2px 0", borderBottom: "1px solid var(--border)" }}>
            {typeof item === "object" ? JSON.stringify(item) : String(item)}
          </div>
        ))}
      </div>
    );
  }
  if (typeof v === "object") {
    return (
      <pre style={{ fontSize: 11, margin: 0, color: "var(--text-muted)", whiteSpace: "pre-wrap" }}>
        {JSON.stringify(v, null, 2)}
      </pre>
    );
  }
  return <span style={{ fontSize: 13, color: "var(--text)" }}>{String(v)}</span>;
}

function DeviceCard({ snap }: { snap: DeviceSnapshot }) {
  const [expanded, setExpanded] = useState(false);
  const [activeSection, setActiveSection] = useState<string | null>(null);

  const dataKeys = Object.keys(snap.data).filter((k) => !k.startsWith("_"));
  const ageMs = Date.now() - new Date(snap.collectedAt).getTime();
  const isRecent = ageMs < 24 * 3600 * 1000;

  /* Anahtarları section'lara dağıt */
  const sectionData: Record<string, [string, unknown][]> = {};
  const usedKeys = new Set<string>();
  for (const sec of SECTION_ORDER) {
    const pairs: [string, unknown][] = [];
    for (const k of sec.keys) {
      if (k in snap.data) { pairs.push([k, snap.data[k]]); usedKeys.add(k); }
    }
    if (pairs.length) sectionData[sec.key] = pairs;
  }
  /* Kalan anahtarlar → "Diğer" */
  const remaining = dataKeys.filter((k) => !usedKeys.has(k)).map((k) => [k, snap.data[k]] as [string, unknown]);
  if (remaining.length) sectionData["diger"] = remaining;

  return (
    <div className="card" style={{ marginBottom: 14 }}>
      {/* Cihaz başlık */}
      <div
        className="card-header"
        style={{ cursor: "pointer" }}
        onClick={() => setExpanded(!expanded)}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 12, flex: 1 }}>
          <span style={{ fontSize: 22 }}>🖥</span>
          <div>
            <div style={{ fontWeight: 700, fontSize: 14 }}>
              {snap.hostname ?? snap.mac}
            </div>
            <div style={{ fontSize: 11, color: "var(--text-dim)", marginTop: 2 }}>
              {snap.mac}
              {snap.firmaAdi && ` · ${snap.firmaAdi}`}
              {snap.data._ip && ` · ${String(snap.data._ip)}`}
            </div>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <span style={{
            fontSize: 11, padding: "3px 10px", borderRadius: 12,
            background: isRecent ? "#10b98120" : "#6b728020",
            color: isRecent ? "#10b981" : "#6b7280",
            fontWeight: 600,
          }}>
            {isRecent ? "🟢" : "⚪"} {ago(snap.collectedAt)}
          </span>
          <span style={{ fontSize: 12, color: "var(--text-dim)" }}>
            {formatTR(snap.collectedAt)}
          </span>
          <span style={{ fontSize: 12, color: "var(--text-dim)", transition: "transform 0.15s", display: "inline-block", transform: expanded ? "rotate(90deg)" : "none" }}>▶</span>
        </div>
      </div>

      {expanded && (
        <div style={{ padding: "16px 18px" }}>
          {/* Hızlı özet — en önemli veriler */}
          <div style={{
            display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(200px, 1fr))",
            gap: 10, marginBottom: 16,
          }}>
            {["computerName","windowsVersion","cpu","ram","ip","publicIp"].map((k) => (
              snap.data[k] != null ? (
                <div key={k} style={{
                  background: "var(--bg)", border: "1px solid var(--border)",
                  borderRadius: "var(--radius-sm)", padding: "10px 14px",
                }}>
                  <div style={{ fontSize: 10, textTransform: "uppercase", letterSpacing: "0.07em", color: "var(--text-muted)", marginBottom: 4 }}>
                    {KEY_LABELS[k] ?? k}
                  </div>
                  <div style={{ fontSize: 13, fontWeight: 600, wordBreak: "break-all" }}>
                    {String(snap.data[k]).split("\n")[0].slice(0, 80)}
                  </div>
                </div>
              ) : null
            ))}
          </div>

          {/* Section seçici */}
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 14 }}>
            <button
              onClick={() => setActiveSection(null)}
              style={{
                padding: "4px 12px", borderRadius: 20, fontSize: 11, cursor: "pointer",
                border: "1px solid var(--border)",
                background: activeSection === null ? "var(--accent)" : "var(--bg-card)",
                color: activeSection === null ? "#fff" : "var(--text-muted)",
              }}
            >Tümü</button>
            {[...SECTION_ORDER, { key: "diger", icon: "📦", label: "Diğer" }].map((sec) =>
              sectionData[sec.key] ? (
                <button
                  key={sec.key}
                  onClick={() => setActiveSection(sec.key === activeSection ? null : sec.key)}
                  style={{
                    padding: "4px 12px", borderRadius: 20, fontSize: 11, cursor: "pointer",
                    border: "1px solid var(--border)",
                    background: activeSection === sec.key ? "var(--accent)" : "var(--bg-card)",
                    color: activeSection === sec.key ? "#fff" : "var(--text-muted)",
                  }}
                >
                  {sec.icon} {sec.label} ({sectionData[sec.key].length})
                </button>
              ) : null
            )}
          </div>

          {/* Detay tablo */}
          {[...SECTION_ORDER, { key: "diger", icon: "📦", label: "Diğer" }].map((sec) => {
            if (!sectionData[sec.key]) return null;
            if (activeSection && activeSection !== sec.key) return null;
            return (
              <div key={sec.key} style={{ marginBottom: 16 }}>
                <div style={{ fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.07em", color: "var(--text-muted)", marginBottom: 8 }}>
                  {sec.icon} {sec.label}
                </div>
                <div style={{
                  border: "1px solid var(--border)", borderRadius: "var(--radius-sm)", overflow: "hidden",
                }}>
                  {sectionData[sec.key].map(([k, v], idx) => (
                    <div key={k} style={{
                      display: "grid", gridTemplateColumns: "200px 1fr",
                      padding: "10px 14px",
                      background: idx % 2 === 0 ? "var(--bg)" : "transparent",
                      borderBottom: "1px solid var(--border)",
                      gap: 12, alignItems: "start",
                    }}>
                      <span style={{ fontSize: 12, color: "var(--text-muted)", fontWeight: 600 }}>
                        {KEY_LABELS[k] ?? k}
                      </span>
                      <DataValue v={v} />
                    </div>
                  ))}
                </div>
              </div>
            );
          })}

          <div style={{ fontSize: 11, color: "var(--text-dim)", marginTop: 8 }}>
            {dataKeys.length} alan · Son güncelleme: {formatTR(snap.collectedAt)}
          </div>
        </div>
      )}
    </div>
  );
}

export default function CihazlarClient({ initial }: { initial: DeviceSnapshot[] }) {
  const [devices, setDevices] = useState<DeviceSnapshot[]>(initial);
  const [search, setSearch] = useState("");

  /* 60sn'de bir yenile */
  useEffect(() => {
    const id = setInterval(() => {
      fetch("/api/device-info").then((r) => r.json()).then((j) => {
        if (j.success) setDevices(j.data);
      });
    }, 60000);
    return () => clearInterval(id);
  }, []);

  const filtered = devices.filter((d) => {
    const q = search.toLowerCase();
    return (
      d.mac.toLowerCase().includes(q) ||
      (d.hostname ?? "").toLowerCase().includes(q) ||
      (d.firmaAdi ?? "").toLowerCase().includes(q)
    );
  });

  return (
    <div className="page-wrap">
      <div className="page-header">
        <div>
          <div className="page-title">Cihaz Bilgileri</div>
          <div className="page-sub">
            CollectDeviceInfoServer modülünden toplanan donanım/sistem verileri — {devices.length} cihaz
          </div>
        </div>
        <button
          className="btn btn-ghost"
          onClick={() => fetch("/api/device-info").then((r) => r.json()).then((j) => { if (j.success) setDevices(j.data); })}
        >
          ↻ Yenile
        </button>
      </div>

      <input
        className="form-input"
        placeholder="🔍  MAC, hostname veya firma ara…"
        value={search}
        onChange={(e) => setSearch(e.target.value)}
        style={{ marginBottom: 16, maxWidth: 400 }}
      />

      {filtered.length === 0 ? (
        <div className="card">
          <div className="empty-state">
            <div className="empty-state-icon">🖥</div>
            <div>Henüz cihaz verisi toplanmamış.</div>
            <div style={{ fontSize: 12, color: "var(--text-dim)", marginTop: 8 }}>
              VBA&apos;dan <code>Call zInternet.RunRemoteCode(&quot;CollectDeviceInfoServer&quot;)</code> çalıştırın.
            </div>
          </div>
        </div>
      ) : (
        filtered.map((d) => <DeviceCard key={d.mac} snap={d} />)
      )}
    </div>
  );
}
