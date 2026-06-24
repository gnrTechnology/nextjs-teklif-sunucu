/**
 * Tek modulu Neon'a yazar.
 * node scripts/upsert-module.mjs CollectDeviceInfoServer data/audit/CollectDeviceInfoServer.bas
 */
import fs from "fs";
import path from "path";
import { neon } from "@neondatabase/serverless";

const methodName = process.argv[2];
const codePath = process.argv[3];

if (!methodName || !codePath) {
  console.error("Kullanim: node scripts/upsert-module.mjs <MethodName> <code.bas>");
  process.exit(1);
}

const env = fs.readFileSync(".env.local", "utf8");
const url = env.match(/^DATABASE_URL=(.+)$/m)[1].trim();
const sql = neon(url);

const code = fs.readFileSync(path.resolve(codePath), "utf8");

await sql`
  UPDATE modules
  SET code = ${code}, updated_at = NOW()
  WHERE method_name = ${methodName}
`;

console.log("Guncellendi:", methodName, "(" + code.length + " karakter)");
