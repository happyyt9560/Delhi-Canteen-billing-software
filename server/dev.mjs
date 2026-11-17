import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const api = spawn(process.execPath, ['server/server.js'], { stdio: 'inherit' });
const vitePath = fileURLToPath(new URL('../node_modules/vite/bin/vite.js', import.meta.url));
const vite = spawn(process.execPath, [vitePath], { stdio: 'inherit' });

const stop = () => {
  if (!api.killed) api.kill();
  if (!vite.killed) vite.kill();
};

process.on('SIGINT', () => { stop(); process.exit(); });
process.on('SIGTERM', () => { stop(); process.exit(); });
api.on('exit', code => { if (code && !vite.killed) vite.kill(); });
vite.on('exit', code => { stop(); process.exit(code || 0); });
