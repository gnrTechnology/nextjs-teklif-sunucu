import fs from "fs";
import { neon } from "@neondatabase/serverless";
const sql = neon(
  fs.readFileSync(".env.local", "utf8").match(/^DATABASE_URL=(.+)$/m)[1].trim(),
);
const rows = await sql`SELECT method_name, code, length(code) as len FROM modules`;
const h = rows.filter((r) => r.code.includes("ORTAK YARDIMCILAR"));
const af = rows.filter((r) => /Columns\.AutoFit/i.test(r.code));
const clr = rows.filter((r) => /Cells\.ClearContents/i.test(r.code));
const huge = rows.filter((r) => r.code.length > 8000);
console.log("ORTAK YARDIMCILAR:", h.length);
console.log("AutoFit:", af.length);
console.log("ClearContents:", clr.length);
console.log("code > 8KB:", huge.length);
console.log(
  "avg code len:",
  Math.round(rows.reduce((s, r) => s + r.code.length, 0) / rows.length),
);
