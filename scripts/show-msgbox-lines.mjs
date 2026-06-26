import fs from "fs";
import { neon } from "@neondatabase/serverless";

const sql = neon(
  fs.readFileSync(".env.local", "utf8").match(/^DATABASE_URL=(.+)$/m)[1].trim(),
);
const names = [
  "KillProcessByName",
  "ShowYesNoCancelDialog",
  "DeleteFolder",
  "CleanTempFolder",
  "DeleteFile",
  "RestartWindowsService",
];
for (const n of names) {
  const r = await sql`SELECT code FROM modules WHERE method_name = ${n}`;
  if (!r[0]) continue;
  const lines = r[0].code.split(/\r?\n/).filter((l) => /MsgBox/i.test(l));
  console.log("\n" + n + ":");
  lines.forEach((l) => console.log(" ", l.trim()));
}
