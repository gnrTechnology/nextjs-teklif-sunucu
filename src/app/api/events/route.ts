export const dynamic = "force-dynamic";
export const runtime = "nodejs";

import { getLatestLogId } from "@/lib/db";

const POLL_INTERVAL_MS = 3000;
const MAX_CYCLES = 20; // 60 saniye sonra client yeniden bağlanır

/**
 * GET /api/events
 * Server-Sent Events: license_logs tablosunda yeni kayıt oluştuğunda "update" eventi gönderir.
 * Vercel'de max ~60s çalışır; EventSource client otomatik yeniden bağlanır.
 */
export async function GET() {
  const encoder = new TextEncoder();

  let initialId: number;
  try {
    initialId = await getLatestLogId();
  } catch {
    initialId = 0;
  }

  const stream = new ReadableStream({
    async start(controller) {
      // Bağlantı kuruldu bildirimi
      controller.enqueue(encoder.encode(`: connected\n\n`));

      let lastId = initialId;

      for (let cycle = 0; cycle < MAX_CYCLES; cycle++) {
        await new Promise<void>((resolve) => setTimeout(resolve, POLL_INTERVAL_MS));

        try {
          const latestId = await getLatestLogId();
          if (latestId !== lastId) {
            lastId = latestId;
            const payload = JSON.stringify({ latestLogId: latestId, ts: Date.now() });
            controller.enqueue(
              encoder.encode(`event: update\ndata: ${payload}\n\n`),
            );
          } else {
            // Bağlantıyı canlı tut
            controller.enqueue(encoder.encode(`: heartbeat\n\n`));
          }
        } catch {
          // DB hatasında sessizce devam et
        }
      }

      controller.close();
    },
  });

  return new Response(stream, {
    headers: {
      "Content-Type": "text/event-stream",
      "Cache-Control": "no-cache, no-transform",
      Connection: "keep-alive",
      "X-Accel-Buffering": "no",
    },
  });
}
