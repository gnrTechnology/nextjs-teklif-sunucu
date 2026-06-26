export type ThemePreference = "dark" | "light" | "system";

export type CihazlarPagePrefs = {
  view?: "grid" | "list";
  search?: string;
  selectedMac?: string | null;
};

export type UserPreferences = {
  theme?: ThemePreference;
  macFilter?: string;
  pages?: {
    cihazlar?: CihazlarPagePrefs;
    [key: string]: Record<string, unknown> | undefined;
  };
};

export const DEFAULT_USER_PREFERENCES: UserPreferences = {
  theme: "system",
  macFilter: "",
  pages: {
    cihazlar: {
      view: "grid",
      search: "",
      selectedMac: null,
    },
  },
};

export function mergeUserPreferences(
  base: UserPreferences,
  patch: Partial<UserPreferences>,
): UserPreferences {
  const next: UserPreferences = { ...base, ...patch };
  if (patch.pages) {
    next.pages = { ...base.pages };
    for (const [key, val] of Object.entries(patch.pages)) {
      if (val && typeof val === "object") {
        next.pages[key] = { ...(base.pages?.[key] ?? {}), ...val };
      }
    }
  }
  return next;
}
