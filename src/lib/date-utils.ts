/** UTC ISO string'i Türkiye saatine (UTC+3) çevirerek gösterim formatı döndürür */
export function formatTR(isoString?: string | null): string {
  if (!isoString) return "—";
  try {
    return parseDbTimestamp(isoString).toLocaleString("tr-TR", {
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

/** DB'den gelen zaman damgasını Date'e çevirir (eski nowTR +3h yazımını düzeltir) */
export function parseDbTimestamp(isoString: string): Date {
  const d = new Date(isoString);
  // Eski kayitlar: UTC'ye +3 saat eklenmis saat yazilmisti → gelecekte gorunur
  if (d.getTime() - Date.now() > 60_000) {
    return new Date(d.getTime() - 3 * 60 * 60 * 1000);
  }
  return d;
}

/** Saniyeyi "3 dk 12 sn" gibi okunabilir süreye çevirir */
export function formatDurationSec(totalSec: number): string {
  const sec = Math.max(0, Math.floor(totalSec));
  if (sec < 60) return `${sec} sn`;
  const min = Math.floor(sec / 60);
  const rem = sec % 60;
  if (min < 60) return rem > 0 ? `${min} dk ${rem} sn` : `${min} dk`;
  const h = Math.floor(min / 60);
  const remMin = min % 60;
  return remMin > 0 ? `${h} sa ${remMin} dk` : `${h} sa`;
}

/** İki zaman damgası arasındaki süreyi saniye olarak döndürür */
export function elapsedSecSince(isoString?: string | null): number | null {
  if (!isoString) return null;
  return Math.max(0, Math.floor((Date.now() - parseDbTimestamp(isoString).getTime()) / 1000));
}

/** "3 dk önce" gibi goreceli zaman metni */
export function timeAgo(isoString?: string | null): string {
  if (!isoString) return "—";
  const diffMs = Math.max(0, Date.now() - parseDbTimestamp(isoString).getTime());
  const diffSec = Math.floor(diffMs / 1000);
  if (diffSec < 10) return "az önce";
  if (diffSec < 60) return `${diffSec} sn önce`;
  const diffMin = Math.floor(diffSec / 60);
  if (diffMin < 60) return `${diffMin} dk önce`;
  const diffH = Math.floor(diffMin / 60);
  if (diffH < 24) return `${diffH} sa önce`;
  return `${Math.floor(diffH / 24)} gün önce`;
}

/** Heartbeat nabiz durumu — TeklifAgent varsayilan araligi ~1-5 dk */
export type HeartbeatStatus = "online" | "idle" | "offline";

export function getHeartbeatStatus(isoString?: string | null): HeartbeatStatus {
  if (!isoString) return "offline";
  const diffMs = Math.max(0, Date.now() - parseDbTimestamp(isoString).getTime());
  const diffMin = diffMs / 60_000;
  if (diffMin < 5) return "online";
  if (diffMin < 60) return "idle";
  return "offline";
}

/** Şu anki Türkiye saatini ISO string olarak döndürür (db'ye yazmak için değil, gösterim için) */
export function nowTR(): string {
  return new Date().toLocaleString("tr-TR", { timeZone: "Europe/Istanbul" });
}
