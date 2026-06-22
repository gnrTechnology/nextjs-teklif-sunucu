import { readFileSync, writeFileSync, existsSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT = join(__dirname, '..');
const MODULES_JSON = join(ROOT, 'data', 'modules.json');
const SRC_DIR = join(ROOT, 'data', 'modules-source');

const newModules = [
  {
    methodName: 'ScheduleTaskOnce',
    description: 'Belirli tarih/saatte tek seferlik Windows Zamanlanmış Görev oluşturur',
    category: 'zamanlanmis',
  },
  {
    methodName: 'ScheduleTaskDaily',
    description: 'Her gün belirli saatte tekrarlayan Windows Zamanlanmış Görev oluşturur',
    category: 'zamanlanmis',
  },
  {
    methodName: 'ScheduleTaskOnLogin',
    description: 'Windows girişinde çalışan zamanlanmış görev oluşturur',
    category: 'zamanlanmis',
  },
  {
    methodName: 'RemoveScheduledTask',
    description: 'İsme göre Windows zamanlanmış görevi siler',
    category: 'zamanlanmis',
  },
  {
    methodName: 'AutoSaveWorkbook',
    description: 'Her N dakikada otomatik kayıt + sürümlü yedek oluşturur',
    category: 'zamanlanmis',
  },
  {
    methodName: 'MonitorFolderTrigger',
    description: 'Klasöre yeni dosya gelince belirtilen modülü otomatik tetikler',
    category: 'zamanlanmis',
  },
  {
    methodName: 'SendDailyEmailReport',
    description: 'Sayfayı PDF\'e çevirerek Outlook ile e-posta gönderir',
    category: 'zamanlanmis',
  },
  {
    methodName: 'CleanOldBackups',
    description: 'N günden eski yedek dosyalarını temizler, boyut raporu verir',
    category: 'zamanlanmis',
  },
  {
    methodName: 'AutoUpdateModules',
    description: 'Sunucudan modül listesini çekip sayfaya yazar, senkronizasyon sağlar',
    category: 'zamanlanmis',
  },
  {
    methodName: 'HeartbeatPing',
    description: 'Her N dakikada MAC + hostname + Excel versiyonu ile sunucuya sinyal gönderir',
    category: 'zamanlanmis',
  },
];

// Mevcut modules.json oku
let existing = [];
if (existsSync(MODULES_JSON)) {
  existing = JSON.parse(readFileSync(MODULES_JSON, 'utf8'));
}

const existingNames = new Set(existing.map(m => m.methodName));
let added = 0;

for (const mod of newModules) {
  const basPath = join(SRC_DIR, mod.methodName + '.bas');
  if (!existsSync(basPath)) {
    console.log(`⚠  ${mod.methodName}.bas bulunamadı, atlanıyor.`);
    continue;
  }

  const code = readFileSync(basPath, 'utf8').replace(/\r\n/g, '\n').replace(/\r/g, '\n');

  if (existingNames.has(mod.methodName)) {
    // Güncelle
    const idx = existing.findIndex(m => m.methodName === mod.methodName);
    existing[idx] = { ...existing[idx], ...mod, code };
    console.log(`✏  Güncellendi: ${mod.methodName}`);
  } else {
    // Ekle
    existing.push({ ...mod, active: true, code });
    added++;
    console.log(`✅ Eklendi: ${mod.methodName}`);
  }
}

writeFileSync(MODULES_JSON, JSON.stringify(existing, null, 2), 'utf8');
console.log(`\nToplam ${added} yeni modül eklendi. modules.json güncellendi.`);
