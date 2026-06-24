"use client";

import { usePathname, useRouter } from "next/navigation";
import { useEffect, useRef } from "react";

const OPS_PREFIXES = [
  "/",
  "/heartbeats",
  "/komutlar",
  "/modul-ciktilari",
  "/cihazlar",
  "/loglar",
  "/klasor-izleme",
  "/lisanslar",
];

function isOpsPage(pathname: string) {
  if (pathname === "/") return true;
  return OPS_PREFIXES.some((p) => p !== "/" && pathname.startsWith(p));
}

/** Operasyon sayfalarında SSE ile server component verisini yeniler */
export default function LiveRefresh() {
  const pathname = usePathname();
  const router = useRouter();
  const esRef = useRef<EventSource | null>(null);

  useEffect(() => {
    if (!isOpsPage(pathname)) return;

    function connect() {
      esRef.current?.close();
      const es = new EventSource("/api/events");
      esRef.current = es;
      es.addEventListener("update", () => router.refresh());
      es.onerror = () => {
        es.close();
        setTimeout(connect, 3000);
      };
    }

    connect();
    return () => esRef.current?.close();
  }, [pathname, router]);

  return null;
}
