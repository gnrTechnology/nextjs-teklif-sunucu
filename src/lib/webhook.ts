/**
 * Opsiyonel webhook bildirimi — WEBHOOK_URL env ile yapılandırılır.
 */
export async function sendWebhook(payload: {
  event: string;
  message: string;
  data?: Record<string, unknown>;
}): Promise<void> {
  const url = process.env.WEBHOOK_URL?.trim();
  if (!url) return;

  try {
    await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        ...payload,
        timestamp: new Date().toISOString(),
        source: "teklif-sunucu",
      }),
    });
  } catch (err) {
    console.error("[webhook]", err);
  }
}
