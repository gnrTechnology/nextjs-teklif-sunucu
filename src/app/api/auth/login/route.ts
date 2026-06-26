import { NextResponse } from "next/server";

export async function POST(request: Request) {
  const adminPassword = process.env.ADMIN_PASSWORD;
  if (!adminPassword) {
    return NextResponse.json({ success: true, message: "Auth devre dışı" });
  }
  const body = await request.json().catch(() => ({}));
  const password = String(body.password ?? "");
  if (password !== adminPassword) {
    return NextResponse.json({ success: false }, { status: 401 });
  }
  const res = NextResponse.json({ success: true });
  const userKey = process.env.ADMIN_USER_ID ?? "admin";
  const cookieOpts = {
    httpOnly: true,
    sameSite: "lax" as const,
    path: "/",
    maxAge: 60 * 60 * 24 * 7,
  };
  res.cookies.set("teklif_admin", adminPassword, cookieOpts);
  res.cookies.set("teklif_user", userKey, cookieOpts);
  return res;
}
