"use client";

import { useRouter } from "next/navigation";
import { useState, useTransition } from "react";
import type { LicenseRecord } from "@/lib/types";
import { LicenseBadge } from "@/lib/status-badges";
import { useToast } from "./ToastProvider";
import { AlertTriangle } from "lucide-react";

export default function LisanslarTable({ licenses }: { licenses: LicenseRecord[] }) {
  const router = useRouter();
  const { toast } = useToast();
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [isPending, startTransition] = useTransition();

  const toggleSelect = (mac: string) =>
    setSelected((prev) => {
      const next = new Set(prev);
      next.has(mac) ? next.delete(mac) : next.add(mac);
      return next;
    });

  const selectAll = () =>
    setSelected(
      selected.size === licenses.length
        ? new Set()
        : new Set(licenses.map((l) => l.macAdresi)),
    );

  async function bulkAction(action: "activate" | "deactivate") {
    if (selected.size === 0) return;
    const macs = Array.from(selected);

    const res = await fetch("/api/license/bulk", {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ macs, action }),
    });

    if (res.ok) {
      setSelected(new Set());
      toast(`${macs.length} lisans ${action === "activate" ? "aktifleştirildi" : "pasifleştirildi"}`, "success");
      startTransition(() => router.refresh());
    } else {
      toast("Toplu işlem başarısız", "error");
    }
  }

  async function singleToggle(mac: string, current: string) {
    const isActive = ["true", "1", "active", "evet"].includes(current.toLowerCase());
    const res = await fetch(`/api/license/${encodeURIComponent(mac)}`, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ action: isActive ? "deactivate" : "activate" }),
    });
    if (res.ok) {
      toast(isActive ? "Lisans pasifleştirildi" : "Lisans aktifleştirildi", "success");
      startTransition(() => router.refresh());
    } else {
      toast("Değişiklik kaydedilemedi", "error");
    }
  }

  const allSelected = licenses.length > 0 && selected.size === licenses.length;
  const someSelected = selected.size > 0;

  return (
    <>
      {/* Toplu işlem araç çubuğu */}
      <div
        className="bulk-toolbar"
        style={{ opacity: someSelected ? 1 : 0.4, pointerEvents: someSelected ? "auto" : "none" }}
      >
        <span className="bulk-count">{selected.size} seçili</span>
        <button
          className="btn btn-green"
          onClick={() => bulkAction("activate")}
          disabled={isPending || !someSelected}
        >
          ✔ Toplu Aktifleştir
        </button>
        <button
          className="btn btn-red"
          onClick={() => bulkAction("deactivate")}
          disabled={isPending || !someSelected}
        >
          ✖ Toplu Pasifleştir
        </button>
        <button
          className="btn btn-ghost"
          onClick={() => setSelected(new Set())}
          disabled={!someSelected}
        >
          Seçimi Temizle
        </button>
      </div>

      <div className="table-wrap">
        <table className="data-table">
          <thead>
            <tr>
              <th style={{ width: 36 }}>
                <input
                  type="checkbox"
                  checked={allSelected}
                  onChange={selectAll}
                  title="Tümünü seç"
                />
              </th>
              <th>MAC Adresi</th>
              <th>Firma</th>
              <th>Kullanıcı</th>
              <th>Dosya</th>
              <th>IP Adresi</th>
              <th>Lisans</th>
              <th>Oluşturma</th>
              <th>Güncelleme</th>
              <th>İşlem</th>
            </tr>
          </thead>
          <tbody>
            {licenses.map((item) => {
              const isActive = ["true", "1", "active", "evet"].includes(
                item.license.toLowerCase(),
              );
              const isViolation =
                item.dosyaAdi && item.dosyaAdi.toLowerCase() !== "teklif.xlam";

              return (
                <tr
                  key={item.macAdresi}
                  className={isViolation ? "row-violation" : ""}
                  style={selected.has(item.macAdresi) ? { background: "var(--hover)" } : {}}
                >
                  <td>
                    <input
                      type="checkbox"
                      checked={selected.has(item.macAdresi)}
                      onChange={() => toggleSelect(item.macAdresi)}
                    />
                  </td>
                  <td>
                    <span className="mono">{item.macAdresi}</span>
                  </td>
                  <td>{item.firmaAdi ?? <span style={{ color: "var(--text-dim)" }}>—</span>}</td>
                  <td>{item.userAdi ?? <span style={{ color: "var(--text-dim)" }}>—</span>}</td>
                  <td>
                    <span className={`mono ${isViolation ? "mono-warn" : ""}`}>
                      {item.dosyaAdi ?? "—"}
                    </span>
                    {isViolation && (
                      <span title="İzinsiz kopya! dosyaAdi teklif.xlam değil." style={{ marginLeft: 4, color: "var(--red)" }}>
                        ⚠
                      </span>
                    )}
                  </td>
                  <td>
                    <span className="mono">{item.ipAdresi ?? "—"}</span>
                  </td>
                  <td>
                    <LicenseBadge value={item.license} />
                  </td>
                  <td style={{ color: "var(--text-muted)", fontSize: 12, whiteSpace: "nowrap" }}>
                    {new Date(item.createdAt).toLocaleString("tr-TR")}
                  </td>
                  <td style={{ color: "var(--text-muted)", fontSize: 12, whiteSpace: "nowrap" }}>
                    {new Date(item.updatedAt).toLocaleString("tr-TR")}
                  </td>
                  <td>
                    <button
                      className={`toggle-btn ${isActive ? "toggle-btn-deactivate" : "toggle-btn-activate"}`}
                      onClick={() => singleToggle(item.macAdresi, item.license)}
                      disabled={isPending}
                    >
                      {isActive ? "Pasifleştir" : "Aktifleştir"}
                    </button>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </>
  );
}
