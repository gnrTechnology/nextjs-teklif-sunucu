const CACHE = new Map<string, { country: string; city?: string; expires: number }>();

export async function lookupGeo(ip: string): Promise<{ country: string; city?: string } | null> {
  const clean = ip.trim();
  if (!clean || clean.startsWith("127.") || clean.startsWith("10.") || clean.startsWith("192.168.")) {
    return null;
  }

  const cached = CACHE.get(clean);
  if (cached && cached.expires > Date.now()) {
    return { country: cached.country, city: cached.city };
  }

  try {
    const res = await fetch(`http://ip-api.com/json/${encodeURIComponent(clean)}?fields=status,country,city`, {
      next: { revalidate: 86400 },
    });
    const j = await res.json() as { status?: string; country?: string; city?: string };
    if (j.status !== "success" || !j.country) return null;
    CACHE.set(clean, { country: j.country, city: j.city, expires: Date.now() + 86400000 });
    return { country: j.country, city: j.city };
  } catch {
    return null;
  }
}
