const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('canteen', {
  saveReceiptPdf: receipt => ipcRenderer.invoke('receipt:save-pdf', receipt),
  openReceiptPdf: filePath => ipcRenderer.invoke('receipt:open-pdf', filePath),
  shareReceiptOnWhatsApp: details => ipcRenderer.invoke('receipt:share-whatsapp', details),
  onBillingShortcut: callback => {
    const listener = (_event, shortcut) => callback(shortcut);
    ipcRenderer.on('billing:shortcut', listener);
    return () => ipcRenderer.removeListener('billing:shortcut', listener);
  },
});
