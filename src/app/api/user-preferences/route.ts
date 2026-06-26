import { NextResponse } from "next/server";
import { getUserKey, isAdminAuthenticated } from "@/lib/admin-auth";
import { getUserPreferences, saveUserPreferences } from "@/lib/db";
import type { UserPreferences } from "@/lib/user-preferences-types";

export async function GET() {
  if (!(await isAdminAuthenticated())) {
    return NextResponse.json({ success: false, message: "Yetkisiz" }, { status: 401 });
  }
  const userKey = await getUserKey();
  const prefs = await getUserPreferences(userKey);
  return NextResponse.json({ success: true, userKey, prefs });
}

export async function PATCH(request: Request) {
  if (!(await isAdminAuthenticated())) {
    return NextResponse.json({ success: false, message: "Yetkisiz" }, { status: 401 });
  }
  const body = (await request.json().catch(() => ({}))) as Partial<UserPreferences>;
  const userKey = await getUserKey();
  const prefs = await saveUserPreferences(userKey, body);
  return NextResponse.json({ success: true, userKey, prefs });
}
