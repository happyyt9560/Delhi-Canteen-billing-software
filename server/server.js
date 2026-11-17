import { createServer } from 'node:http';
import { mkdirSync, writeFileSync, existsSync, readFileSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { DatabaseSync } from 'node:sqlite';

const dataDirectory = process.env.CANTEEN_DATA_DIR || join(process.cwd(), 'data');
mkdirSync(dataDirectory, { recursive: true });
const database = new DatabaseSync(join(dataDirectory, 'canteen.sqlite'));
// WAL plus FULL synchronisation gives committed billing records durable
// SQLite transactions while still allowing safe reads during a backup.
database.exec('PRAGMA journal_mode = WAL; PRAGMA synchronous = FULL; PRAGMA busy_timeout = 5000;');
const snapshotPath = join(dataDirectory, 'backup-snapshot.json');
database.exec(`CREATE TABLE IF NOT EXISTS products (
  id TEXT PRIMARY KEY,
  payload TEXT NOT NULL,
  updated_at TEXT NOT NULL
)`);
database.exec(`CREATE TABLE IF NOT EXISTS app_state (
  id INTEGER PRIMARY KEY CHECK (id = 1),
  payload TEXT NOT NULL,
  updated_at TEXT NOT NULL
)`);

const readAppState = () => {
  const row = database.prepare('SELECT payload, updated_at FROM app_state WHERE id = 1').get();
  if (!row) return null;
  return { data: JSON.parse(row.payload), updatedAt: row.updated_at };
};

// The full application state and its product index are updated inside one
// SQLite transaction. A power loss cannot leave a new bill saved without its
// matching product data (or vice versa).
const saveAppState = state => {
  if (!state || typeof state !== 'object' || !Array.isArray(state.products)) throw new Error('Invalid application data.');
  const updatedAt = new Date().toISOString();
  const payload = JSON.stringify(state);
  const saveState = database.prepare(`INSERT INTO app_state (id, payload, updated_at) VALUES (1, ?, ?)
    ON CONFLICT(id) DO UPDATE SET payload = excluded.payload, updated_at = excluded.updated_at`);
  const saveProduct = database.prepare(`INSERT INTO products (id, payload, updated_at) VALUES (?, ?, ?)
    ON CONFLICT(id) DO UPDATE SET payload = excluded.payload, updated_at = excluded.updated_at`);
  database.exec('BEGIN IMMEDIATE');
  try {
    saveState.run(payload, updatedAt);
    database.exec('DELETE FROM products');
    for (const product of state.products) {
      if (product?.id && product?.name) saveProduct.run(product.id, JSON.stringify(product), updatedAt);
    }
    database.exec('COMMIT');
    return { updatedAt };
  } catch (error) {
    database.exec('ROLLBACK');
    throw error;
  }
};

const changeStoredProduct = (productId, value) => {
  const stored = readAppState();
  if (!stored) return;
  const products = value
    ? stored.data.products.map(product => product.id === productId ? value : product)
    : stored.data.products.filter(product => product.id !== productId);
  if (value && !stored.data.products.some(product => product.id === productId)) products.unshift(value);
  saveAppState({ ...stored.data, products });
};

const send = (response, status, body) => {
  response.writeHead(status, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
  response.end(JSON.stringify(body));
};

const readBody = request => new Promise((resolve, reject) => {
  let body = '';
  request.on('data', chunk => { body += chunk; if (body.length > 50 * 1024 * 1024) reject(new Error('Request is too large')); });
  request.on('end', () => resolve(body));
  request.on('error', reject);
});

const safeBackupName = value => String(value || 'manual-backup')
  .trim().replace(/[^a-z0-9 _-]/gi, '').replace(/\s+/g, '-').slice(0, 80) || 'manual-backup';
const backupDirectory = 'C:\\Delhi Canteen Daily Backup';
const automaticBackupFilename = (now = new Date()) => {
  const date = [now.getFullYear(), String(now.getMonth() + 1).padStart(2, '0'), String(now.getDate()).padStart(2, '0')].join('-');
  return `delhi-canteen-backup-${date}.json`;
};
const createBackup = (name, snapshot) => {
  if (!snapshot || typeof snapshot !== 'object') throw new Error('No app data is available to back up yet.');
  mkdirSync(backupDirectory, { recursive: true });
  const document = { application: 'Delhi Canteen Billing Software', version: 1, createdAt: new Date().toISOString(), data: snapshot };
  const filename = `${safeBackupName(name)}.json`;
  const outputPath = join(backupDirectory, filename);
  writeFileSync(outputPath, JSON.stringify(document, null, 2), 'utf8');
  return { filename, path: outputPath, createdAt: document.createdAt };
};

createServer((request, response) => {
  const { pathname } = new URL(request.url, 'http://localhost');
  if (request.method === 'OPTIONS') {
    response.writeHead(204, { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'GET,POST,PUT,DELETE,OPTIONS', 'Access-Control-Allow-Headers': 'Content-Type' });
    return response.end();
  }
  if (pathname === '/api/backup-snapshot' && request.method === 'POST') {
    readBody(request).then(body => {
      const snapshot = JSON.parse(body);
      if (!snapshot || typeof snapshot !== 'object') throw new Error('Invalid snapshot');
      writeFileSync(snapshotPath, JSON.stringify(snapshot), 'utf8');
      send(response, 200, { saved: true });
    }).catch(() => send(response, 400, { error: 'Could not save backup data.' }));
    return;
  }
  if (pathname === '/api/app-state' && request.method === 'GET') {
    try {
      const stored = readAppState();
      // `products` supports a one-time migration from the app's earlier
      // products-only SQLite database.
      const products = database.prepare('SELECT payload FROM products ORDER BY updated_at DESC').all().map(row => JSON.parse(row.payload));
      return send(response, 200, { ...stored, products });
    } catch {
      return send(response, 500, { error: 'Could not read billing data.' });
    }
  }
  if (pathname === '/api/app-state' && request.method === 'PUT') {
    readBody(request).then(body => {
      const state = JSON.parse(body || '{}');
      return send(response, 200, saveAppState(state));
    }).catch(error => send(response, 400, { error: error.message || 'Could not save billing data.' }));
    return;
  }
  if (pathname === '/api/backups' && request.method === 'POST') {
    readBody(request).then(body => {
      const { name, data } = JSON.parse(body || '{}');
      const snapshot = data && typeof data === 'object' ? data : (existsSync(snapshotPath) ? JSON.parse(readFileSync(snapshotPath, 'utf8')) : null);
      return send(response, 201, createBackup(name, snapshot));
    }).catch(error => send(response, 400, { error: error.message || 'Could not create backup.' }));
    return;
  }
  if (pathname === '/api/backups/status' && request.method === 'GET') {
    const filename = automaticBackupFilename();
    const outputPath = join(backupDirectory, filename);
    const exists = existsSync(outputPath);
    return send(response, 200, {
      filename,
      path: outputPath,
      exists,
      updatedAt: exists ? statSync(outputPath).mtime.toISOString() : null,
    });
  }
  if (pathname === '/api/backups/restore' && request.method === 'POST') {
    readBody(request).then(body => {
      const snapshot = JSON.parse(body || '{}');
      if (!snapshot || typeof snapshot !== 'object' || !Array.isArray(snapshot.products)) throw new Error('This is not a valid Delhi Canteen backup file.');
      writeFileSync(snapshotPath, JSON.stringify(snapshot), 'utf8');
      saveAppState(snapshot);
      send(response, 200, { restored: true });
    }).catch(error => send(response, 400, { error: error.message || 'Could not restore this backup.' }));
    return;
  }
  if (pathname === '/api/products' && request.method === 'GET') {
    const products = database.prepare('SELECT payload FROM products ORDER BY updated_at DESC').all().map(row => JSON.parse(row.payload));
    return send(response, 200, products);
  }
  if (pathname === '/api/products' && request.method === 'POST') {
    let body = '';
    request.on('data', chunk => { body += chunk; });
    request.on('end', () => {
      try {
        const product = JSON.parse(body);
        if (!product.id || !product.name) return send(response, 400, { error: 'Product id and name are required.' });
        database.prepare(`INSERT INTO products (id, payload, updated_at) VALUES (?, ?, ?)
          ON CONFLICT(id) DO UPDATE SET payload = excluded.payload, updated_at = excluded.updated_at`)
          .run(product.id, JSON.stringify(product), new Date().toISOString());
        changeStoredProduct(product.id, product);
        send(response, 201, product);
      } catch { send(response, 400, { error: 'Invalid product data.' }); }
    });
    return;
  }
  if (pathname === '/api/products' && request.method === 'DELETE') {
    database.exec('DELETE FROM products');
    const stored = readAppState();
    if (stored) saveAppState({ ...stored.data, products: [] });
    return send(response, 200, { cleared: true });
  }
  if (pathname.startsWith('/api/products/') && request.method === 'DELETE') {
    const productId = decodeURIComponent(pathname.slice('/api/products/'.length));
    if (!productId) return send(response, 400, { error: 'Product id is required.' });
    database.prepare('DELETE FROM products WHERE id = ?').run(productId);
    changeStoredProduct(productId, null);
    return send(response, 200, { deleted: productId });
  }
  send(response, 404, { error: 'Route not found.' });
}).listen(3001, () => console.log('SQLite API running at http://localhost:3001'));
