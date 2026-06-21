import { readFileSync, writeFileSync } from 'fs';

const basContent = readFileSync('data/modules-source/getLicense.bas', 'utf8');

// Normalize line endings to \r\n then encode
const normalized = basContent.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
const lines = normalized.split('\n');
const jsonSafe = lines.join('\r\n')
  .replace(/\\/g, '\\\\')
  .replace(/"/g, '\\"')
  .replace(/\r\n/g, '\\r\\n')
  .replace(/\t/g, '\\t');

const modules = JSON.parse(readFileSync('data/modules.json', 'utf8'));
const idx = modules.findIndex(m => m.methodName === 'getLicense');
if (idx === -1) { console.error('getLicense not found!'); process.exit(1); }

modules[idx].code = jsonSafe;
modules[idx].description = 'Lisans kontrol ve kayit; firmaAdi (mdip) ve userAdi (TBveren) registry den okunur';

writeFileSync('data/modules.json', JSON.stringify(modules, null, 2), 'utf8');
console.log('modules.json guncellendi. getLicense code uzunlugu:', jsonSafe.length, 'karakter');
