"use client";

import { useRouter } from "next/navigation";
import { useState } from "react";

interface Props {
  mac: string;
  current: string;
}

export default function ToggleLicenseButton({ mac, current }: Props) {
  const router = useRouter();
  const [loading, setLoading] = useState(false);

  const isActive = ["true", "1", "active", "evet"].includes(current.toLowerCase());
  const next = isActive ? "false" : "true";
  const label = isActive ? "Pasifleştir" : "Aktifleştir";

  async function handleClick() {
    setLoading(true);
    try {
      await fetch(`/api/license/${encodeURIComponent(mac)}/`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ license: next }),
      });
      router.refresh();
    } finally {
      setLoading(false);
    }
  }

  return (
    <button
      onClick={handleClick}
      disabled={loading}
      className={`toggle-btn ${isActive ? "toggle-btn--deactivate" : "toggle-btn--activate"}`}
    >
      {loading ? "..." : label}
    </button>
  );
}
