import { readFileSync, writeFileSync } from 'fs';

const basContent = readFileSync('data/modules-source/getLicense.bas', 'utf8');

// Sadece satir sonlarini normalize et; JSON.stringify geri kalan escaping'i yapar.
// Manuel escape YAPMA — JSON.stringify zaten " ve \ gibi karakterleri dogru escape eder.
const normalized = basContent.replace(/\r\n/g, '\n').replace(/\r/g, '\n');

const modules = JSON.parse(readFileSync('data/modules.json', 'utf8'));
const idx = modules.findIndex(m => m.methodName === 'getLicense');
if (idx === -1) { console.error('getLicense not found!'); process.exit(1); }

modules[idx].code = normalized;
modules[idx].description = 'Lisans kontrol ve kayit; firmaAdi (mdip) ve userAdi (TBveren) registry den okunur';

writeFileSync('data/modules.json', JSON.stringify(modules, null, 2), 'utf8');
console.log('modules.json guncellendi. getLicense code uzunlugu:', normalized.length, 'karakter');
