/**
 * data/audit/*.bas dosyalarini Neon'a yazar.
 * node scripts/upsert-audit-modules.mjs
 */
import fs from "fs";
import path from "path";
import { neon } from "@neondatabase/serverless";

const auditDir = path.join(process.cwd(), "data", "audit");
const env = fs.readFileSync(".env.local", "utf8");
const url = env.match(/^DATABASE_URL=(.+)$/m)[1].trim();
const sql = neon(url);

const files = fs.readdirSync(auditDir).filter((f) => f.endsWith(".bas"));
for (const file of files) {
  const methodName = file.replace(/\.bas$/i, "");
  const code = fs.readFileSync(path.join(auditDir, file), "utf8");
  await sql`
    UPDATE modules SET code = ${code}, updated_at = NOW()
    WHERE method_name = ${methodName}
  `;
  console.log("upsert:", methodName, "(" + code.length + " char)");
}

console.log("done", files.length);
