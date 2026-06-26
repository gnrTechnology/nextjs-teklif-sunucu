/**
 * Neon'da kalan MsgBox sayisini raporlar.
 * node scripts/count-msgbox-neon.mjs
 */
import fs from "fs";
import { neon } from "@neondatabase/serverless";

const env = fs.readFileSync(".env.local", "utf8");
const sql = neon(env.match(/^DATABASE_URL=(.+)$/m)[1].trim());
const rows = await sql`SELECT method_name, code FROM modules`;
const hits = rows.filter((r) => /MsgBox/i.test(r.code));
console.log("MsgBox iceren modul:", hits.length);
hits.forEach((r) => console.log(" -", r.method_name));
