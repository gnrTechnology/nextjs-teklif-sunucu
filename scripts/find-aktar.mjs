import fs from "fs";
import { neon } from "@neondatabase/serverless";
const sql = neon(fs.readFileSync(".env.local","utf8").match(/^DATABASE_URL=(.+)$/m)[1].trim());
const rows = await sql`SELECT method_name, code FROM modules`;
for (const r of rows) {
  if (/aktar/i.test(r.code)) {
    const lines = r.code.split(/\r?\n/).filter(l => /aktar/i.test(l));
    console.log(r.method_name);
    lines.slice(0,5).forEach(l => console.log(" ", l.trim()));
  }
}
