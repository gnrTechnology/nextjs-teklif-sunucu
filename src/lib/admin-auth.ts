import { cookies } from "next/headers";

export async function isAdminAuthenticated(): Promise<boolean> {
  const adminPassword = process.env.ADMIN_PASSWORD;
  if (!adminPassword) return true;
  const jar = await cookies();
  return jar.get("teklif_admin")?.value === adminPassword;
}

/** Oturum bazlı tercih anahtarı (çoklu kullanıcıya hazır) */
export async function getUserKey(): Promise<string> {
  const jar = await cookies();
  return jar.get("teklif_user")?.value ?? "default";
}
