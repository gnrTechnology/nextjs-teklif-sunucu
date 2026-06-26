"use client";

import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import {
  DEFAULT_USER_PREFERENCES,
  mergeUserPreferences,
  type CihazlarPagePrefs,
  type ThemePreference,
  type UserPreferences,
} from "@/lib/user-preferences-types";

type UserPreferencesContextValue = {
  ready: boolean;
  prefs: UserPreferences;
  updatePrefs: (patch: Partial<UserPreferences>) => void;
  updatePage: <K extends keyof NonNullable<UserPreferences["pages"]>>(
    page: K,
    patch: NonNullable<UserPreferences["pages"]>[K],
  ) => void;
};

const UserPreferencesContext = createContext<UserPreferencesContextValue | null>(
  null,
);

export function UserPreferencesProvider({
  children,
}: {
  children: React.ReactNode;
}) {
  const [prefs, setPrefs] = useState<UserPreferences>(DEFAULT_USER_PREFERENCES);
  const [ready, setReady] = useState(false);
  const saveTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const prefsRef = useRef(prefs);
  prefsRef.current = prefs;

  useEffect(() => {
    let cancelled = false;
    fetch("/api/user-preferences")
      .then((r) => (r.ok ? r.json() : null))
      .then((j) => {
        if (cancelled || !j?.success || !j.prefs) return;
        setPrefs(mergeUserPreferences(DEFAULT_USER_PREFERENCES, j.prefs));
      })
      .catch(() => {})
      .finally(() => {
        if (!cancelled) setReady(true);
      });
    return () => {
      cancelled = true;
    };
  }, []);

  const persist = useCallback((next: UserPreferences) => {
    if (saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(() => {
      fetch("/api/user-preferences", {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(next),
      }).catch(() => {});
    }, 350);
  }, []);

  const updatePrefs = useCallback(
    (patch: Partial<UserPreferences>) => {
      setPrefs((prev) => {
        const next = mergeUserPreferences(prev, patch);
        persist(next);
        return next;
      });
    },
    [persist],
  );

  const updatePage = useCallback(
    <K extends keyof NonNullable<UserPreferences["pages"]>>(
      page: K,
      patch: NonNullable<UserPreferences["pages"]>[K],
    ) => {
      setPrefs((prev) => {
        const next = mergeUserPreferences(prev, {
          pages: { [page]: patch } as UserPreferences["pages"],
        });
        persist(next);
        return next;
      });
    },
    [persist],
  );

  const value = useMemo(
    () => ({ ready, prefs, updatePrefs, updatePage }),
    [ready, prefs, updatePrefs, updatePage],
  );

  return (
    <UserPreferencesContext.Provider value={value}>
      {children}
    </UserPreferencesContext.Provider>
  );
}

export function useUserPreferences() {
  const ctx = useContext(UserPreferencesContext);
  if (!ctx) {
    throw new Error("useUserPreferences UserPreferencesProvider içinde kullanılmalı");
  }
  return ctx;
}

export function useThemePreference() {
  const { prefs, updatePrefs, ready } = useUserPreferences();
  const setTheme = useCallback(
    (theme: ThemePreference) => updatePrefs({ theme }),
    [updatePrefs],
  );
  return {
    theme: prefs.theme ?? "system",
    setTheme,
    ready,
  };
}

export function useCihazlarPrefs() {
  const { prefs, updatePage, ready } = useUserPreferences();
  const page = prefs.pages?.cihazlar ?? DEFAULT_USER_PREFERENCES.pages!.cihazlar!;
  const setPage = useCallback(
    (patch: Partial<CihazlarPagePrefs>) => updatePage("cihazlar", patch),
    [updatePage],
  );
  return { page, setPage, ready };
}
