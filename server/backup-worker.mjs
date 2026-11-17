import { existsSync, mkdirSync, readFileSync, renameSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

// This small process is started by Windows Task Scheduler. It has no UI and
// reads the latest snapshot saved by the billing app.
const dataDirectory = process.argv[2] || process.env.CANTEEN_DATA_DIR || join(process.cwd(), 'data');
const snapshotPath = join(dataDirectory, 'backup-snapshot.json');
const backupDirectory = 'C:\\Delhi Canteen Daily Backup';

if (!existsSync(snapshotPath)) {
  console.error('No backup snapshot found. Open the billing app once first.');
  process.exit(1);
}

const now = new Date();
const date = [now.getFullYear(), String(now.getMonth() + 1).padStart(2, '0'), String(now.getDate()).padStart(2, '0')].join('-');
const snapshot = JSON.parse(readFileSync(snapshotPath, 'utf8'));
mkdirSync(backupDirectory, { recursive: true });
const output = join(backupDirectory, `delhi-canteen-backup-${date}.json`);
const temporaryOutput = `${output}.tmp`;
writeFileSync(temporaryOutput, JSON.stringify({ application: 'Delhi Canteen Billing Software', version: 1, createdAt: now.toISOString(), data: snapshot }, null, 2), 'utf8');
renameSync(temporaryOutput, output);
console.log(`Hourly backup updated: ${output}`);
