import fs from "fs";
import path from "path";
import { neon } from "@neondatabase/serverless";

const env = fs.readFileSync(".env.local", "utf8");
const url = env.match(/^DATABASE_URL=(.+)$/m)[1].trim();
const sql = neon(url);

const metaPath = path.join(process.cwd(), "data", "modules-meta.json");
const META = JSON.parse(fs.readFileSync(metaPath, "utf8"));

const dir = path.join(process.cwd(), "data", "modules-staging");
const files = fs.readdirSync(dir).filter((f) => f.endsWith(".bas"));

for (const file of files) {
  const methodName = file.replace(/\.bas$/i, "");
  const code = fs.readFileSync(path.join(dir, file), "utf8");
  const meta = META[methodName] ?? { description: methodName, category: "genel" };
  await sql`
    INSERT INTO modules (method_name, description, category, active, code, created_at, updated_at)
    VALUES (
      ${methodName},
      ${meta.description},
      ${meta.category},
      true,
      ${code},
      NOW(),
      NOW()
    )
    ON CONFLICT (method_name) DO UPDATE SET
      description = EXCLUDED.description,
      category = EXCLUDED.category,
      active = true,
      code = EXCLUDED.code,
      updated_at = NOW()
  `;
  console.log("upserted:", methodName);
}

console.log("done", files.length, "modules");
