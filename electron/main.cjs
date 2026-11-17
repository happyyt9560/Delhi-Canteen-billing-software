const { app, BrowserWindow, ipcMain, shell } = require('electron');
const { spawn, execFile } = require('node:child_process');
const path = require('node:path');
const fs = require('node:fs/promises');

let apiProcess;

// Saving straight to the C: drive fails on some managed Windows PCs (and on
// systems with Controlled Folder Access). Prefer the user's Documents folder,
// then fall back to Electron's own writable application-data directory.
const legacyReceiptDirectory = 'C:\\Delhi Canteen Bill';
const receiptDirectories = () => [...new Set([
  path.join(app.getPath('documents'), 'Delhi Canteen Bill'),
  path.join(app.getPath('userData'), 'receipts'),
  legacyReceiptDirectory,
].map(directory => path.resolve(directory)))];
const fileSafe = value => String(value || '').replace(/[^a-z0-9_-]/gi, '') || 'bill';
const receiptTimestamp = value => {
  const date = new Date(value);
  const safeDate = Number.isNaN(date.getTime()) ? new Date() : date;
  const part = number => String(number).padStart(2, '0');
  return `${safeDate.getFullYear()}-${part(safeDate.getMonth() + 1)}-${part(safeDate.getDate())}_${part(safeDate.getHours())}-${part(safeDate.getMinutes())}-${part(safeDate.getSeconds())}`;
};

const resolveReceiptPdfPath = filePath => {
  const resolvedPath = path.resolve(String(filePath || ''));
  const isReceiptPdf = receiptDirectories().some(directory => resolvedPath.toLowerCase().startsWith(`${directory.toLowerCase()}${path.sep}`));
  if (!isReceiptPdf || path.extname(resolvedPath).toLowerCase() !== '.pdf') {
    throw new Error('This PDF file cannot be shared.');
  }
  return resolvedPath;
};

const whatsappNumber = value => {
  let number = String(value || '').replace(/\D/g, '');
  if (number.startsWith('00')) number = number.slice(2);
  if (number.startsWith('0')) number = number.slice(1);
  // Bills in this app use Indian mobile numbers. Keep an already-entered
  // country code intact, otherwise add India's country code for WhatsApp.
  if (number.length === 10) number = `91${number}`;
  // WhatsApp accepts E.164 phone numbers only.  Do not claim that a chat was
  // selected when the billing phone field contains an incomplete number.
  return number.length >= 8 && number.length <= 15 ? number : '';
};

const powerShellLiteral = value => `'${String(value).replace(/'/g, "''")}'`;

// WhatsApp's public click-to-chat interface can prefill a recipient and text,
// but deliberately does not allow a website/app to attach a local file. Put
// the PDF on the Windows file clipboard so it can be pasted as a real document
// in the chat with Ctrl+V (rather than sending its filename as text).
const copyFileToClipboard = filePath => new Promise((resolve, reject) => {
  if (process.platform !== 'win32') {
    reject(new Error('Copying a PDF to the clipboard is available on Windows only.'));
    return;
  }
  const script = [
    'Add-Type -AssemblyName System.Windows.Forms',
    '$files = New-Object System.Collections.Specialized.StringCollection',
    `$files.Add(${powerShellLiteral(filePath)})`,
    '[System.Windows.Forms.Clipboard]::SetFileDropList($files)',
  ].join('; ');
  execFile('powershell.exe', ['-NoProfile', '-NonInteractive', '-STA', '-Command', script], error => {
    if (error) reject(new Error('PDF could not be copied to the clipboard.'));
    else resolve();
  });
});

ipcMain.handle('receipt:save-pdf', async (_event, receipt) => {
  if (!receipt || typeof receipt.html !== 'string' || !receipt.html.trim()) throw new Error('Receipt could not be prepared for PDF download.');
  const filename = `Bill-${fileSafe(receipt.billNumber)}-${receiptTimestamp(receipt.createdAt)}.pdf`;

  const receiptWindow = new BrowserWindow({
    show: false,
    webPreferences: { contextIsolation: true, nodeIntegration: false, sandbox: true },
  });
  try {
    await receiptWindow.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(receipt.html)}`);
    const pdf = await receiptWindow.webContents.printToPDF({
      printBackground: true,
      margins: { marginType: 'none' },
      pageSize: { width: 80000, height: 297000 },
    });
    let lastError;
    for (const directory of receiptDirectories()) {
      const outputPath = path.join(directory, filename);
      try {
        await fs.mkdir(directory, { recursive: true });
        await fs.writeFile(outputPath, pdf);
        return { path: outputPath, filename };
      } catch (error) {
        lastError = error;
      }
    }
    throw new Error(`PDF could not be saved to Documents or app storage.${lastError?.message ? ` ${lastError.message}` : ''}`);
  } finally {
    if (!receiptWindow.isDestroyed()) receiptWindow.destroy();
  }
});

ipcMain.handle('receipt:open-pdf', async (_event, filePath) => {
  const resolvedPath = resolveReceiptPdfPath(filePath);
  const error = await shell.openPath(resolvedPath);
  if (error) throw new Error(error);
});

ipcMain.handle('receipt:share-whatsapp', async (_event, details) => {
  const phone = whatsappNumber(details?.phone);
  const text = String(details?.text || 'Your bill receipt is attached.');
  const url = phone
    ? `https://wa.me/${phone}?${new URLSearchParams({ text }).toString()}`
    : `https://web.whatsapp.com/send?${new URLSearchParams({ text }).toString()}`;

  // A text-only bill needs no temporary PDF. This is the reliable share path
  // for WhatsApp: it opens the selected customer chat with the entire receipt
  // ready in the message box.
  if (!details?.filePath) {
    await shell.openExternal(url);
    return { hasRecipient: Boolean(phone), textOnly: true };
  }

  const filePath = resolveReceiptPdfPath(details.filePath);
  let copiedToClipboard = true;
  try {
    await copyFileToClipboard(filePath);
  } catch {
    // Opening the customer chat is still useful if another Windows program
    // currently owns the clipboard. In that case reveal the PDF as a manual
    // attachment fallback instead of cancelling the entire share action.
    copiedToClipboard = false;
  }
  await shell.openExternal(url);
  if (!copiedToClipboard) shell.showItemInFolder(filePath);
  return { filePath, hasRecipient: Boolean(phone), copiedToClipboard };
});

function installHourlyBackupTask() {
  if (process.platform !== 'win32') return;
  const workerPath = app.isPackaged
    ? path.join(process.resourcesPath, 'server', 'backup-worker.mjs')
    : path.join(__dirname, '..', 'server', 'backup-worker.mjs');
  const dataPath = path.join(app.getPath('userData'), 'data');
  // Each trigger runs at an exact clock hour. StartWhenAvailable makes Windows
  // run a missed backup after the next sign-in/startup; a switched-off computer
  // cannot run a scheduled task until Windows starts again.
  const workerCommand = `$env:ELECTRON_RUN_AS_NODE='1'; & '${process.execPath.replace(/'/g, "''")}' '${workerPath.replace(/'/g, "''")}' '${dataPath.replace(/'/g, "''")}'`;
  const encodedWorkerCommand = Buffer.from(workerCommand, 'utf16le').toString('base64');
  const script = [
    `$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-NoProfile -NonInteractive -EncodedCommand ${encodedWorkerCommand}'`,
    "$triggers = 0..23 | ForEach-Object { New-ScheduledTaskTrigger -Daily -At ('{0:00}:00' -f $_) }",
    '$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable',
    "Register-ScheduledTask -TaskName 'Delhi Canteen Daily Backup' -Action $action -Trigger $triggers -Settings $settings -Force | Out-Null",
  ].join('; ');
  execFile('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', script], error => {
    if (error) console.warn('Could not install hourly backup task:', error.message);
  });
}

function startApi() {
  const serverPath = app.isPackaged
    ? path.join(process.resourcesPath, 'server', 'server.js')
    : path.join(__dirname, '..', 'server', 'server.js');
  apiProcess = spawn(process.execPath, [serverPath], {
    env: { ...process.env, ELECTRON_RUN_AS_NODE: '1', CANTEEN_DATA_DIR: path.join(app.getPath('userData'), 'data') },
    stdio: 'inherit',
  });
}

function createWindow() {
  const window = new BrowserWindow({ width: 1280, height: 800, minWidth: 960, minHeight: 650, autoHideMenuBar: true, webPreferences: { contextIsolation: true, nodeIntegration: false, preload: path.join(__dirname, 'preload.cjs') } });
  // On Windows, Ctrl + is usually reported as Ctrl + Shift + =. Electron's
  // default zoom shortcut does not handle that layout consistently in the
  // packaged app, so support both main-keyboard and numpad variants here.
  window.webContents.on('before-input-event', (event, input) => {
    if (input.type !== 'keyDown' || !input.control || input.alt || input.meta) return;
    if (input.key.toLowerCase() === 'w') {
      // Ctrl+W is Electron's default window-close command. Reserve it for
      // the Billing screen's Edit product shortcut instead.
      event.preventDefault();
      window.webContents.send('billing:shortcut', 'edit-product');
      return;
    }
    const zoomIn = input.key === '+' || input.key === '=' || input.key === 'Add' || input.code === 'NumpadAdd';
    const zoomOut = input.key === '-' || input.key === 'Subtract' || input.code === 'NumpadSubtract';
    const resetZoom = input.key === '0' || input.key === 'Num0' || input.code === 'Numpad0';
    if (!zoomIn && !zoomOut && !resetZoom) return;
    event.preventDefault();
    if (resetZoom) window.webContents.setZoomLevel(0);
    else window.webContents.setZoomLevel(window.webContents.getZoomLevel() + (zoomIn ? 1 : -1));
  });
  // The desktop command loads the freshly built renderer.  A Vite URL is only
  // used when it is deliberately supplied for live development.
  if (process.env.ELECTRON_RENDERER_URL) window.loadURL(process.env.ELECTRON_RENDERER_URL);
  else window.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));
}

app.whenReady().then(() => { startApi(); installHourlyBackupTask(); createWindow(); app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); }); });
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
app.on('before-quit', () => apiProcess?.kill());
