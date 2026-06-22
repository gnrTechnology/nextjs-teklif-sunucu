import { jsonResponse, errorResponse } from "@/lib/api-response";
import { seedModulesFromJson } from "@/lib/seed-modules";

/** GET|POST /api/seed-modules/ — modules.json → Neon (kolay erişim, trailing slash uyumlu) */
export async function GET() {
  return runSeed();
}

export async function POST() {
  return runSeed();
}

async function runSeed() {
  try {
    const { seeded, errors } = await seedModulesFromJson();
    return jsonResponse({
      success: true,
      seeded,
      errors: errors.length > 0 ? errors : undefined,
      message: `${seeded} modül Neon DB'ye aktarıldı.`,
    });
  } catch (err) {
    console.error("[/api/seed-modules]", err);
    return errorResponse(String(err), 500);
  }
}
