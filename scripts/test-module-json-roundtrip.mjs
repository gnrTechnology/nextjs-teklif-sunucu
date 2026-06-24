import fs from "fs";
import { neon } from "@neondatabase/serverless";

const env = fs.readFileSync(".env.local", "utf8");
const url = env.match(/^DATABASE_URL=(.+)$/m)[1].trim();
const sql = neon(url);

const rows = await sql`SELECT code FROM modules WHERE method_name = 'CollectDeviceInfoServer'`;
const code = rows[0].code;

function extractCodeFromJSON(jsonText) {
  let p1 = jsonText.toLowerCase().indexOf('"code"');
  if (p1 < 0) return jsonText;
  p1 = jsonText.indexOf(":", p1);
  p1 = jsonText.indexOf('"', p1) + 1;
  const p2 = jsonText.lastIndexOf('"');
  if (p2 <= p1) return "";
  let tempStr = jsonText.slice(p1, p2);
  tempStr = tempStr.replace(/\\"/g, '"');
  tempStr = tempStr.replace(/\\r\\n/g, "\r\n");
  tempStr = tempStr.replace(/\\n/g, "\r\n");
  tempStr = tempStr.replace(/\\t/g, "\t");
  tempStr = tempStr.replace(/\\\\/g, "\\");
  return tempStr;
}

const wrapped = JSON.stringify({ success: true, code });
const extracted = extractCodeFromJSON(wrapped);

console.log("Original length:", code.length);
console.log("Extracted length:", extracted.length);
console.log("Match:", code === extracted);
if (code !== extracted) {
  for (let i = 0; i < Math.max(code.length, extracted.length); i++) {
    if (code[i] !== extracted[i]) {
      console.log("First diff at", i);
      console.log("Orig:", JSON.stringify(code.slice(i, i + 80)));
      console.log("Extr:", JSON.stringify(extracted.slice(i, i + 80)));
      break;
    }
  }
}
