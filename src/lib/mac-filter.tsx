"use client";

import { createContext, useCallback, useContext, useEffect, useRef, useState } from "react";
import { useRouter, useSearchParams } from "next/navigation";
import { useUserPreferences } from "@/lib/user-preferences";

const STORAGE_KEY = "teklif_mac_filter";

type MacFilterCtx = {
  mac: string;
  setMac: (mac: string) => void;
  clearMac: () => void;
  matchesMac: (value: string | null | undefined) => boolean;
};

const MacFilterContext = createContext<MacFilterCtx | null>(null);

export function MacFilterProvider({ children }: { children: React.ReactNode }) {
  const router = useRouter();
  const searchParams = useSearchParams();
  const urlMac = searchParams.get("mac") ?? "";
  const { prefs, updatePrefs, ready } = useUserPreferences();
  const [mac, setMacState] = useState("");
  const hydrated = useRef(false);

  useEffect(() => {
    if (urlMac) {
      setMacState(urlMac);
      try {
        sessionStorage.setItem(STORAGE_KEY, urlMac);
      } catch {
        /* ignore */
      }
      if (ready && urlMac !== prefs.macFilter) {
        updatePrefs({ macFilter: urlMac });
      }
      return;
    }
    if (ready && !hydrated.current) {
      hydrated.current = true;
      const fromDb = prefs.macFilter ?? "";
      if (fromDb) {
        setMacState(fromDb);
        try {
          sessionStorage.setItem(STORAGE_KEY, fromDb);
        } catch {
          /* ignore */
        }
        return;
      }
    }
    try {
      const stored = sessionStorage.getItem(STORAGE_KEY) ?? "";
      if (stored) setMacState(stored);
    } catch {
      /* ignore */
    }
  }, [urlMac, ready, prefs.macFilter, updatePrefs]);

  const setMac = useCallback(
    (value: string) => {
      const v = value.trim();
      setMacState(v);
      try {
        if (v) sessionStorage.setItem(STORAGE_KEY, v);
        else sessionStorage.removeItem(STORAGE_KEY);
      } catch {
        /* ignore */
      }
      updatePrefs({ macFilter: v });
    },
    [updatePrefs],
  );

  const clearMac = useCallback(() => {
    setMac("");
    try {
      sessionStorage.removeItem(STORAGE_KEY);
    } catch {
      /* ignore */
    }
    const params = new URLSearchParams(searchParams.toString());
    params.delete("mac");
    const q = params.toString();
    router.replace(q ? `?${q}` : window.location.pathname);
  }, [router, searchParams, setMac]);

  const matchesMac = useCallback(
    (value: string | null | undefined) => {
      if (!mac) return true;
      return (value ?? "").toLowerCase().includes(mac.toLowerCase());
    },
    [mac],
  );

  return (
    <MacFilterContext.Provider value={{ mac, setMac, clearMac, matchesMac }}>
      {children}
    </MacFilterContext.Provider>
  );
}

export function useMacFilter() {
  const ctx = useContext(MacFilterContext);
  if (!ctx) throw new Error("useMacFilter MacFilterProvider içinde kullanılmalı");
  return ctx;
}

export function MacFilterBar() {
  const { mac, clearMac } = useMacFilter();
  if (!mac) return null;

  return (
    <div className="mac-filter-bar" role="status">
      <span>
        MAC filtresi: <code className="mono">{mac}</code>
      </span>
      <button type="button" className="btn btn-ghost mac-filter-clear" onClick={clearMac}>
        Temizle
      </button>
    </div>
  );
}
