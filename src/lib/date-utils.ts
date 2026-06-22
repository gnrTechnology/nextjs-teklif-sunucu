/** UTC ISO string'i Türkiye saatine (UTC+3) çevirerek gösterim formatı döndürür */
export function formatTR(isoString?: string | null): string {
  if (!isoString) return "—";
  try {
    return new Date(isoString).toLocaleString("tr-TR", {
      timeZone: "Europe/Istanbul",
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    });
  } catch {
    return isoString;
  }
}

/** Şu anki Türkiye saatini ISO string olarak döndürür (db'ye yazmak için değil, gösterim için) */
export function nowTR(): string {
  return new Date().toLocaleString("tr-TR", { timeZone: "Europe/Istanbul" });
}
