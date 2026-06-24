"use client";

import { useState, useEffect, useMemo } from "react";
import {
  Monitor,
  LayoutGrid,
  List,
  Cpu,
  HardDrive,
  Wifi,
  Download,
} from "lucide-react";
import { formatTR } from "@/lib/date-utils";
import type { DeviceSnapshot } from "@/lib/db";
import { useMacFilter } from "@/lib/mac-filter";
import PageHeader from "./ui/PageHeader";
import EmptyState from "./ui/EmptyState";
import DetailDrawer from "./ui/DetailDrawer";

function ago(iso: string): string {
  const ms = Date.now() - new Date(iso).getTime();
  const min = Math.floor(ms / 60000);
  if (min < 2) return "Az önce";
  if (min < 60) return `${min} dk önce`;
  const h = Math.floor(min / 60);
  if (h < 24) return `${h} sa önce`;
  return `${Math.floor(h / 24)} gün önce`;
}

const SECTION_ORDER = [
  { key: "bilgisayar", icon: Monitor, label: "Bilgisayar", keys: ["computerName", "windowsVersion", "windowsActivation", "systemUptime", "timeZone", "locale", "domainName", "loggedInUser"] },
  { key: "donanim", icon: Cpu, label: "Donanım", keys: ["cpu", "ram", "gpu", "screenResolution", "bios", "motherboard"] },
  { key: "disk", icon: HardDrive, label: "Disk", keys: ["disks", "diskInfo"] },
  { key: "ag", icon: Wifi, label: "Ağ", keys: ["mac", "ip", "publicIp", "networkAdapters", "wifiProfiles", "bitlockerStatus"] },
];

const KEY_LABELS: Record<string, string> = {
  computerName: "Bilgisayar Adı",
  windowsVersion: "Windows Sürümü",
  windowsActivation: "Aktivasyon",
  systemUptime: "Çalışma Süresi",
  timeZone: "Saat Dilimi",
  locale: "Yerel Ayarlar",
  domainName: "Domain",
  loggedInUser: "Kullanıcı",
  cpu: "İşlemci",
  ram: "Bellek",
  gpu: "Ekran Kartı",
  screenResolution: "Ekran",
  bios: "BIOS",
  motherboard: "Anakart",
  disks: "Diskler",
  mac: "MAC",
  ip: "IP",
  publicIp: "Dış IP",
  battery: "Pil",
};

function DataValue({ v }: { v: unknown }) {
  if (v === null || v === undefined) return <span className="text-dim">—</span>;
  if (typeof v === "object") {
    return (
      <pre className="device-data-pre">
        {JSON.stringify(v, null, 2)}
      </pre>
    );
  }
  return <span>{String(v)}</span>;
}

function DeviceDetail({ snap }: { snap: DeviceSnapshot }) {
  const dataKeys = Object.keys(snap.data).filter((k) => !k.startsWith("_"));
  const used = new Set<string>();
  const sections: { label: string; pairs: [string, unknown][] }[] = [];

  for (const sec of SECTION_ORDER) {
    const pairs: [string, unknown][] = [];
    for (const k of sec.keys) {
      if (k in snap.data) {
        pairs.push([k, snap.data[k]]);
        used.add(k);
      }
    }
    if (pairs.length) sections.push({ label: sec.label, pairs });
  }
  const rest = dataKeys.filter((k) => !used.has(k)).map((k) => [k, snap.data[k]] as [string, unknown]);
  if (rest.length) sections.push({ label: "Diğer", pairs: rest });

  function exportJson() {
    const blob = new Blob([JSON.stringify(snap, null, 2)], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `cihaz-${snap.mac.replace(/:/g, "-")}.json`;
    a.click();
    URL.revokeObjectURL(a.href);
  }

  return (
    <>
      <div className="device-quick-stats">
        {["cpu", "ram", "ip", "windowsVersion"].map((k) =>
          snap.data[k] != null ? (
            <div key={k} className="device-stat-chip">
              <div className="device-stat-label">{KEY_LABELS[k] ?? k}</div>
              <div className="device-stat-value">{String(snap.data[k]).slice(0, 60)}</div>
            </div>
          ) : null,
        )}
      </div>
      {sections.map((sec) => (
        <div key={sec.label} className="device-detail-section">
          <div className="device-detail-section-title">{sec.label}</div>
          <div className="device-detail-table">
            {sec.pairs.map(([k, v]) => (
              <div key={k} className="device-detail-row">
                <span className="device-detail-key">{KEY_LABELS[k] ?? k}</span>
                <DataValue v={v} />
              </div>
            ))}
          </div>
        </div>
      ))}
      <button type="button" className="btn btn-ghost" onClick={exportJson} style={{ marginTop: 12 }}>
        <Download size={14} />
        JSON indir
      </button>
      <p className="text-dim" style={{ fontSize: 11, marginTop: 12 }}>
        {dataKeys.length} alan · {formatTR(snap.collectedAt)}
      </p>
    </>
  );
}

function DeviceGridCard({ snap, onOpen }: { snap: DeviceSnapshot; onOpen: () => void }) {
  const ageMs = Date.now() - new Date(snap.collectedAt).getTime();
  const isRecent = ageMs < 24 * 3600 * 1000;

  return (
    <button type="button" className="card device-grid-card" onClick={onOpen} style={{ textAlign: "left", width: "100%" }}>
      <div style={{ padding: "14px 16px" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
          <Monitor size={20} style={{ color: "var(--accent)", flexShrink: 0 }} />
          <div style={{ minWidth: 0 }}>
            <div style={{ fontWeight: 700, fontSize: 14 }}>{snap.hostname ?? snap.mac}</div>
            <div className="mono text-dim" style={{ fontSize: 11 }}>{snap.mac}</div>
          </div>
        </div>
        {snap.firmaAdi && (
          <div style={{ fontSize: 12, color: "var(--text-muted)", marginBottom: 8 }}>{snap.firmaAdi}</div>
        )}
        <div style={{ fontSize: 12, color: "var(--text-muted)", lineHeight: 1.5 }}>
          {snap.data.cpu ? <div>{String(snap.data.cpu).slice(0, 50)}</div> : null}
          {snap.data.ram ? <div>{String(snap.data.ram)}</div> : null}
        </div>
        <div
          style={{
            marginTop: 10,
            fontSize: 11,
            fontWeight: 600,
            color: isRecent ? "var(--green)" : "var(--text-dim)",
          }}
        >
          {isRecent ? "● " : "○ "}
          {ago(snap.collectedAt)}
        </div>
      </div>
    </button>
  );
}

export default function CihazlarClient({ initial }: { initial: DeviceSnapshot[] }) {
  const { mac: globalMac, matchesMac } = useMacFilter();
  const [devices, setDevices] = useState<DeviceSnapshot[]>(initial);
  const [search, setSearch] = useState("");
  const [view, setView] = useState<"grid" | "list">("grid");
  const [selected, setSelected] = useState<DeviceSnapshot | null>(null);

  useEffect(() => {
    setDevices(initial);
  }, [initial]);

  async function reload() {
    const r = await fetch("/api/device-info");
    const j = await r.json();
    if (j.success) setDevices(j.data);
  }

  const filtered = useMemo(() => {
    const q = search.toLowerCase();
    return devices.filter((d) => {
      if (!matchesMac(d.mac)) return false;
      if (!q) return true;
      return (
        d.mac.toLowerCase().includes(q) ||
        (d.hostname ?? "").toLowerCase().includes(q) ||
        (d.firmaAdi ?? "").toLowerCase().includes(q)
      );
    });
  }, [devices, search, matchesMac]);

  const recentCount = devices.filter(
    (d) => Date.now() - new Date(d.collectedAt).getTime() < 86400000,
  ).length;

  return (
    <div className="page-wrap page-wrap--wide">
      <PageHeader
        title="Cihaz Bilgileri"
        subtitle={`CollectDeviceInfoServer — ${devices.length} cihaz · ${recentCount} son 24 saat`}
        live
        actions={
          <div className="device-view-toggle">
            <button type="button" className={view === "grid" ? "active" : ""} onClick={() => setView("grid")} title="Kart görünümü">
              <LayoutGrid size={14} />
            </button>
            <button type="button" className={view === "list" ? "active" : ""} onClick={() => setView("list")} title="Liste">
              <List size={14} />
            </button>
          </div>
        }
      />

      <div style={{ display: "flex", gap: 10, marginBottom: 16, flexWrap: "wrap" }}>
        <input
          className="form-input"
          placeholder="MAC, hostname veya firma ara…"
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          style={{ flex: 1, minWidth: 200, maxWidth: 400 }}
        />
        {globalMac && (
          <span className="badge badge-accent" style={{ alignSelf: "center" }}>
            Global MAC: {globalMac}
          </span>
        )}
      </div>

      {filtered.length === 0 ? (
        <div className="card">
          <EmptyState
            icon={Monitor}
            title="Henüz cihaz verisi yok"
            description="Excel'de CollectDeviceInfoServer modülünü çalıştırın; veriler otomatik bu sayfaya düşer."
            action={
              <code className="mono" style={{ fontSize: 12 }}>
                Call zInternet.RunRemoteCode(&quot;CollectDeviceInfoServer&quot;)
              </code>
            }
          />
        </div>
      ) : view === "grid" ? (
        <div className="device-grid">
          {filtered.map((d) => (
            <DeviceGridCard key={d.mac} snap={d} onOpen={() => setSelected(d)} />
          ))}
        </div>
      ) : (
        <div className="card" style={{ padding: 0 }}>
          {filtered.map((d) => (
            <button
              key={d.mac}
              type="button"
              className="device-list-row"
              onClick={() => setSelected(d)}
            >
              <Monitor size={16} style={{ color: "var(--accent)" }} />
              <span style={{ fontWeight: 600 }}>{d.hostname ?? d.mac}</span>
              <span className="mono text-dim" style={{ fontSize: 11 }}>{d.mac}</span>
              <span className="text-dim" style={{ marginLeft: "auto", fontSize: 12 }}>{ago(d.collectedAt)}</span>
            </button>
          ))}
        </div>
      )}

      <DetailDrawer
        open={!!selected}
        title={selected?.hostname ?? selected?.mac ?? ""}
        subtitle={selected ? `${selected.mac}${selected.firmaAdi ? ` · ${selected.firmaAdi}` : ""}` : undefined}
        onClose={() => setSelected(null)}
      >
        {selected && <DeviceDetail snap={selected} />}
      </DetailDrawer>
    </div>
  );
}
