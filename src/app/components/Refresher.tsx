"use client";

import { useRouter } from "next/navigation";
import { useEffect, useRef, useState } from "react";

export default function Refresher() {
  const router = useRouter();
  const [lastRefresh, setLastRefresh] = useState<string>("");
  const [status, setStatus] = useState<"connecting" | "live" | "error">("connecting");
  const esRef = useRef<EventSource | null>(null);

  function doRefresh() {
    router.refresh();
    setLastRefresh(new Date().toLocaleTimeString("tr-TR"));
  }

  useEffect(() => {
    setLastRefresh(new Date().toLocaleTimeString("tr-TR"));

    function connect() {
      if (esRef.current) esRef.current.close();

      const es = new EventSource("/api/events");
      esRef.current = es;

      es.onopen = () => setStatus("live");

      es.addEventListener("update", () => {
        doRefresh();
      });

      // SSE 60s sonra kapanır → otomatik yeniden bağlan
      es.onerror = () => {
        setStatus("connecting");
        es.close();
        setTimeout(connect, 2000);
      };
    }

    connect();

    return () => {
      esRef.current?.close();
    };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  return (
    <div className="refresher">
      <span
        className="refresher-dot"
        title={status === "live" ? "Gerçek zamanlı (SSE)" : "Bağlanıyor…"}
        style={{ background: status === "live" ? "var(--green)" : status === "error" ? "var(--red)" : "var(--yellow)" }}
      />
      <span className="refresher-label">
        {status === "live" ? "Canlı" : status === "error" ? "Hata" : "Bağlanıyor…"}
        {lastRefresh && <>&nbsp;·&nbsp;Son: {lastRefresh}</>}
      </span>
      <button className="refresher-btn" onClick={doRefresh} title="Şimdi yenile">
        ↻
      </button>
    </div>
  );
}
