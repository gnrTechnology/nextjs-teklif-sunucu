"use client";

import { useRouter } from "next/navigation";
import { useEffect, useState } from "react";

const INTERVAL = 10_000; // 10 saniye

export default function Refresher() {
  const router = useRouter();
  const [countdown, setCountdown] = useState(INTERVAL / 1000);
  const [lastRefresh, setLastRefresh] = useState<string>("");

  useEffect(() => {
    setLastRefresh(new Date().toLocaleTimeString("tr-TR"));

    const tick = setInterval(() => {
      setCountdown((c) => {
        if (c <= 1) {
          router.refresh();
          setLastRefresh(new Date().toLocaleTimeString("tr-TR"));
          return INTERVAL / 1000;
        }
        return c - 1;
      });
    }, 1000);

    return () => clearInterval(tick);
  }, [router]);

  function handleManualRefresh() {
    router.refresh();
    setLastRefresh(new Date().toLocaleTimeString("tr-TR"));
    setCountdown(INTERVAL / 1000);
  }

  return (
    <div className="refresher">
      <span className="refresher-dot" title="Canlı veri" />
      <span className="refresher-label">
        Son: {lastRefresh || "—"} &nbsp;·&nbsp; {countdown}s
      </span>
      <button className="refresher-btn" onClick={handleManualRefresh} title="Şimdi yenile">
        ↻
      </button>
    </div>
  );
}
