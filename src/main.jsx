import React, { createContext, useContext, useEffect, useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { BarChart3, Boxes, ClipboardList, LayoutDashboard, PackagePlus, Plus, Search, Trash2, Pencil, Eye, X, Upload, Printer, Download, Minus, ShoppingCart, AlertTriangle, IndianRupee, Menu, FileText, CheckCircle2, XCircle, FileSpreadsheet, Store as StoreIcon, Sun, Moon } from 'lucide-react';
import { BarChart, Bar, CartesianGrid, LabelList, ResponsiveContainer, Tooltip, XAxis, YAxis, LineChart, Line } from 'recharts';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import './styles.css';

const Store = createContext();
const money = n => `Rs. ${Number(n || 0).toFixed(2)}`;
const roundAmount = n => Math.round((Number(n) + Number.EPSILON) * 100) / 100;
const neatNumber = n => String(roundAmount(n));
const whatsappText = value => String(value ?? '').replace(/\s+/g, ' ').trim();
const whatsappBillText = (bill, shop) => {
  const itemLines = (bill.items || []).flatMap((item, index) => {
    const quantity = neatNumber(item.quantity);
    const price = Number(item.sellingPrice || 0).toFixed(2);
    const total = Number(item.total ?? Number(item.quantity || 0) * Number(item.sellingPrice || 0)).toFixed(2);
    return [`${index + 1}. ${whatsappText(item.productName || item.name || 'Product')}`, `   ${quantity} x Rs. ${price} = Rs. ${total}`];
  });
  const lines = [
    `*${whatsappText(shop.ownerName || 'Delhi Canteen')}*`,
    '*BILL RECEIPT*',
    `Bill No: ${whatsappText(bill.number)}`,
    `Date: ${new Date(bill.createdAt).toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' })}`,
    bill.customerName ? `Customer: ${whatsappText(bill.customerName)}` : null,
    '',
    '*Items*',
    ...itemLines,
    '',
    `Total MRP: Rs. ${Number(bill.subtotal || 0).toFixed(2)}`,
    `Discount: Rs. ${Number(bill.totalDiscount || 0).toFixed(2)}`,
    `*Total Bill: Rs. ${Number(bill.grandTotal || 0).toFixed(2)}*`,
    `Paid: Rs. ${Number(bill.paidAmount || 0).toFixed(2)}`,
    `Remaining: Rs. ${Number(bill.remainingAmount || 0).toFixed(2)}`,
    `Payment: ${whatsappText(bill.payment || 'Cash')}`,
    '',
    'Thank you for shopping with us.',
  ];
  return lines.filter(line => line !== null).join('\n');
};
// Bill totals are collected in whole rupees: amounts ending in .50 or more
// round up (for example, Rs. 20.50 becomes Rs. 21).
const roundBillTotal = n => Math.round(Number(n) || 0);
const uid = () => crypto.randomUUID?.() || `${Date.now()}-${Math.random()}`;
const BILL_NUMBER_START = 10001;
const isFiveDigitBillNumber = value => /^\d{5}$/.test(String(value));
// The next sequence value is stored separately from the bill list. Deleting a
// bill therefore cannot make its number available for another bill.
const billNumber = data => {
  const used = new Set((data?.bills || []).map(bill => String(bill.number)));
  let next = Math.max(BILL_NUMBER_START, Math.trunc(Number(data?.nextBillNumber) || BILL_NUMBER_START));
  while (used.has(String(next).padStart(5, '0'))) next += 1;
  return String(next).padStart(5, '0');
};
const dateKey = d => new Date(d).toISOString().slice(0, 10);
const GENERAL_CATEGORY = { id: 'general', name: 'General' };
const defaultShopProfile = { ownerName: 'Delhi Canteen', phone: '7982181870', email: 'delhicanteen@gmail.com', address: 'mangal bazar road, swarn park, mundka' };
const initial = { products: [], categories: [GENERAL_CATEGORY], bills: [], heldBills: [], customers: [], nextCustomerNumber: 1, shopProfile: defaultShopProfile };
const isGeneralCategory = category => category?.name?.trim().toLowerCase() === 'general';
const ensureGeneralCategory = savedData => {
  const data = savedData && typeof savedData === 'object' ? savedData : initial;
  const categories = Array.isArray(data.categories) ? data.categories : [];
  const existingGeneral = categories.find(isGeneralCategory);
  const assignedBillNumbers = new Set();
  let nextBillNumber = Math.max(BILL_NUMBER_START, Math.trunc(Number(data.nextBillNumber) || BILL_NUMBER_START));
  const reserveBillNumber = () => {
    while (assignedBillNumbers.has(String(nextBillNumber).padStart(5, '0'))) nextBillNumber += 1;
    const number = String(nextBillNumber).padStart(5, '0');
    assignedBillNumbers.add(number);
    nextBillNumber += 1;
    return number;
  };
  const bills = (Array.isArray(data.bills) ? data.bills : []).map(bill => {
    const number = String(bill.number || '');
    if (isFiveDigitBillNumber(number) && !assignedBillNumbers.has(number)) {
      assignedBillNumbers.add(number);
      nextBillNumber = Math.max(nextBillNumber, Number(number) + 1);
      return { ...bill, number };
    }
    return { ...bill, number: reserveBillNumber() };
  });
  const customers = Array.isArray(data.customers) ? data.customers : [];
  const highestCustomerNumber = customers.reduce((max, customer) => Math.max(max, Number(String(customer.customerId || '').match(/\d+$/)?.[0]) || 0), 0);
  return {
    ...data,
    customers,
    heldBills: Array.isArray(data.heldBills) ? data.heldBills : [],
    nextCustomerNumber: Math.max(highestCustomerNumber + 1, Math.trunc(Number(data.nextCustomerNumber) || 1)),
    shopProfile: { ...defaultShopProfile, ...(data.shopProfile || {}) },
    // Keep every product field while normalising categories. In particular,
    // stock must never be discarded during a data load/migration.
    products: (Array.isArray(data.products) ? data.products : []).map(product => {
      return !product.category || product.category.trim().toLowerCase() === 'general'
        ? { ...product, category: 'General' }
        : product;
    }),
    categories: [
      existingGeneral ? { ...existingGeneral, name: 'General' } : GENERAL_CATEGORY,
      ...categories.filter(category => !isGeneralCategory(category)),
    ],
    bills,
    nextBillNumber,
  };
};
const apiUrl = 'http://localhost:3001/api';
function Provider({ children }) {
  const [data, setStoredData] = useState(() => { try { return ensureGeneralCategory(JSON.parse(localStorage.getItem('dc-data'))); } catch { return initial; } });
  const [databaseReady, setDatabaseReady] = useState(false);
  const [databaseError, setDatabaseError] = useState('');
  const persistenceQueue = useMemo(() => ({ current: Promise.resolve() }), []);
  const skipInitialDatabaseSave = useMemo(() => ({ current: false }), []);
  const setData = update => setStoredData(current => ensureGeneralCategory(typeof update === 'function' ? update(current) : update));
  const [notice, setNotice] = useState(null);
  const queueDatabaseSave = snapshot => {
    persistenceQueue.current = persistenceQueue.current
      .catch(() => {})
      .then(() => fetch(`${apiUrl}/app-state`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(snapshot) }))
      .then(response => { if (!response.ok) throw new Error('SQLite could not save the latest billing data.'); });
    return persistenceQueue.current;
  };
  const loadDatabase = () => {
    setDatabaseError('');
    setDatabaseReady(false);
    fetch(`${apiUrl}/app-state`)
      .then(response => response.ok ? response.json() : Promise.reject(new Error('The local billing database could not be opened.')))
      .then(async stored => {
        if (stored.data) {
          skipInitialDatabaseSave.current = true;
          setStoredData(ensureGeneralCategory(stored.data));
          return;
        }
        // First run after the SQLite migration: preserve every existing local
        // record and prefer products already written by the old product API.
        const migrated = ensureGeneralCategory({ ...data, products: stored.products?.length ? stored.products : data.products });
        await queueDatabaseSave(migrated);
        skipInitialDatabaseSave.current = true;
        setStoredData(migrated);
      })
      .then(() => setDatabaseReady(true))
      .catch(error => setDatabaseError(error.message || 'The local billing database is unavailable.'));
  };
  useEffect(() => { loadDatabase(); }, []);
  useEffect(() => {
    if (!databaseReady) return;
    localStorage.setItem('dc-data', JSON.stringify(data));
    if (skipInitialDatabaseSave.current) {
      skipInitialDatabaseSave.current = false;
      return;
    }
    queueDatabaseSave(data).catch(() => setNotice({ type: 'error', text: 'SQLite save failed. Do not close the app; retry after checking the database.' }));
  }, [data, databaseReady]);
  // The background Windows task cannot read a browser's localStorage. Keep a
  // current server-side snapshot so it can safely create a backup even while
  // the application window is closed.
  useEffect(() => {
    if (!databaseReady) return;
    const timer = setTimeout(() => {
      fetch(`${apiUrl}/backup-snapshot`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(data) }).catch(() => { });
    }, 350);
    return () => clearTimeout(timer);
  }, [data, databaseReady]);
  useEffect(() => { if (notice) { const t = setTimeout(() => setNotice(null), 3000); return () => clearTimeout(t); } }, [notice]);
  if (databaseError) return <div className="database-status"><h1>SQLite database unavailable</h1><p>{databaseError}</p><Button type="button" onClick={loadDatabase}>Retry connection</Button></div>;
  if (!databaseReady) return <div className="database-status"><h1>Loading billing data…</h1><p>Checking and protecting your SQLite billing records.</p></div>;
  return <Store.Provider value={{ data, setData, notice, setNotice }}>{children}</Store.Provider>;
}
const useStore = () => useContext(Store);
function Empty({ children }) { return <div className="empty">{children}</div> }
function Button({ children, variant = 'primary', ...p }) { return <button className={`btn ${variant}`} {...p}>{children}</button> }
function Field({ label, children }) { const visibleLabel = label === 'Discount (₹)' ? 'Discount (rs)' : label; return <label className="field"><span>{visibleLabel}</span>{children}</label> }
function SalesTrendTooltip({ active, payload, label }) { if (!active || !payload?.length) return null; return <div className="sales-trend-tooltip"><span>{label}</span><b>{money(payload[0].value)}</b><small>Sales</small></div> }

function Dashboard({ onOpenOfflineBills, onViewAllPendingBills, theme, onToggleTheme }) {
  const { data, setData, setNotice } = useStore(); const now = new Date(); const today = dateKey(now); const [reportDay, setReportDay] = useState(today); const reportMonth = reportDay.slice(0, 7);
  const clearAllData = async () => {
    if (!window.confirm('Clear all data? All products, categories, bills, dashboard records and the current billing draft will be permanently deleted.')) return;
    try {
      const response = await fetch(`${apiUrl}/products`, { method: 'DELETE' });
      if (!response.ok) throw new Error('Could not clear product data');
      sessionStorage.removeItem('dc-billing-draft');
      setData({ products: [], categories: [GENERAL_CATEGORY], bills: [] });
      setNotice({ type: 'success', text: 'All data has been cleared. The app is ready for a fresh start.' });
    } catch {
      setNotice({ type: 'error', text: 'Could not clear all data. Please try again.' });
    }1
  };
  const reportDate = new Date(`${reportMonth}-01T00:00:00`), reportLabel = reportDate.toLocaleDateString('en-IN', { month: 'long', year: 'numeric' });
  const selectedBills = data.bills.filter(bill => dateKey(bill.createdAt).startsWith(reportMonth)), selectedDayBills = data.bills.filter(bill => dateKey(bill.createdAt) === reportDay), todayBillCount = data.bills.filter(bill => dateKey(bill.createdAt) === today).length;
  const isCurrentMonth = reportMonth === dateKey(now).slice(0, 7);
  const sum = bills => bills.reduce((a, b) => a + b.grandTotal, 0);
  const allSales = sum(data.bills), averageOrder = selectedBills.length ? sum(selectedBills) / selectedBills.length : 0, customerCount = (data.customers || []).length;
  const cards = [
    { title: `Net sales · ${reportLabel}`, value: sum(selectedBills), detail: `All-time sales: ${money(allSales)}` },
    { title: `Sales · ${new Date(`${reportDay}T00:00:00`).toLocaleDateString('en-IN', { day: 'numeric', month: 'short' })}`, value: sum(selectedDayBills), detail: `${selectedDayBills.length} bill${selectedDayBills.length === 1 ? '' : 's'} on selected date` },
    { title: 'Average order value', value: averageOrder, detail: `For ${reportLabel}` },
    { title: 'Total customers', value: customerCount, detail: 'Customers added individually and by bulk upload' },
  ];
  const trendEnd = isCurrentMonth ? now : new Date(reportDate.getFullYear(), reportDate.getMonth() + 1, 0);
  const weekly = Array.from({ length: 7 }, (_, i) => { const d = new Date(trendEnd); d.setDate(d.getDate() - 6 + i); const k = dateKey(d); return { day: d.toLocaleDateString('en-IN', { weekday: 'short' }), sales: sum(selectedBills.filter(b => dateKey(b.createdAt) === k)) } });
  const sold = {}; selectedBills.forEach(b => b.items.forEach(i => sold[i.productName] = (sold[i.productName] || 0) + i.quantity)); const top = Object.entries(sold).sort((a, b) => b[1] - a[1]).slice(0, 10);
  const pendingBills = selectedBills.filter(bill => Number(bill.remainingAmount || 0) > 0);
  const pendingTotal = sum(pendingBills.map(bill => ({ grandTotal: Number(bill.remainingAmount || 0) })));
  // Stock tracking has been retired. These empty values only keep the legacy,
  // hidden dashboard panel harmless until its markup is removed in a layout pass.
  const activeProducts = data.products.length, inactiveProducts = 0, lowStockProducts = 0, outOfStockProducts = 0, stockAttentionProducts = [];
  return <><header><div><p className="eyebrow">OVERVIEW</p><h1>Good day, Canteen team</h1><p>View your shop performance for any month or date.</p></div><div className="dashboard-header-actions"><button type="button" className="theme-toggle" onClick={onToggleTheme} aria-label={`Switch to ${theme === 'day' ? 'night' : 'day'} theme`} title={`Switch to ${theme === 'day' ? 'night' : 'day'} theme`}>{theme === 'day' ? <Moon size={17} /> : <Sun size={17} />}<span>{theme === 'day' ? 'Night' : 'Day'}</span></button><div className="date-chip dashboard-month-picker"><span>Dashboard date</span><div><input type="date" value={reportDay} max={today} onChange={event => { const value = event.target.value; if (value) setReportDay(value); }} aria-label="Select dashboard date" /></div></div></div></header>
    <section className="cards dashboard-cards">{cards.map(card => <div className="card" key={card.title}><div><p>{card.title}</p><strong>{card.title === 'Total customers' ? card.value : money(card.value)}</strong><small>{card.detail}</small></div></div>)}</section>
    <section className="grid two"><div className="panel sales-trend"><div className="sales-trend-head"><div><p className="eyebrow">SALES TREND</p><h2>Last 7 days · {reportLabel}</h2></div><strong>{money(sum(weekly.map(day => ({ grandTotal: day.sales }))))}</strong></div><div className="chart"><ResponsiveContainer><BarChart data={weekly} margin={{ top: 12, right: 3, left: 3, bottom: 0 }}><XAxis dataKey="day" axisLine={false} tickLine={false} tick={{ fill: '#6b7280', fontSize: 11, fontWeight: 700 }} /><Bar dataKey="sales" fill="#d92338" radius={[6, 6, 1, 1]} maxBarSize={56}><LabelList dataKey="sales" position="insideTop" fill="#fff" fontSize={11} fontWeight={800} formatter={value => Number(value) > 0 ? `Rs. ${Math.round(Number(value))}` : ''} /></Bar></BarChart></ResponsiveContainer></div><div className="sales-trend-summary"><div><span>Selected month</span><b>{money(sum(selectedBills))}</b></div><div><span>Today's bills</span><b>{todayBillCount}</b></div><div><span>All time</span><b>{money(allSales)}</b></div></div></div>
      <div className="panel"><div className="panel-head"><div><h2>Top selling items</h2><p>Based on completed bills</p></div></div>{top.length ? <div className="rank-list">{top.map(([n, q], i) => <div key={n}><b>{String(i + 1).padStart(2, '0')}</b><span>{n}</span><strong>{q} sold</strong></div>)}</div> : <Empty>No sales data yet. Create a bill to see insights.</Empty>}</div></section>
    <section className="grid two"><div className="panel pending-collection"><div className="pending-head"><div><p className="eyebrow">PENDENCY</p><h2>Remaining amount collection</h2></div><span><small>Total to collect</small><b>{money(pendingTotal)}</b></span></div>{pendingBills.length ? <div className="pending-table"><table><thead><tr><th>BILL NO.</th><th>CUSTOMER</th><th>DATE</th><th>BILL TOTAL</th><th>PAID</th><th>REMAINING</th></tr></thead><tbody>{pendingBills.slice(0, 5).map(bill => { const amount = Number(bill.grandTotal || 0), paidAmount = Number(bill.paidAmount || 0), remaining = Number(bill.remainingAmount || 0); return <tr key={bill.id}><td><b>#{bill.number}</b></td><td>{bill.customerName || 'Prabhu Shri Ram'}</td><td>{new Date(bill.createdAt).toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' })}</td><td>Rs. {amount.toFixed(0)}</td><td>Rs. {paidAmount.toFixed(0)}</td><td><b>Rs. {remaining.toFixed(0)}</b></td></tr> })}</tbody></table></div> : <Empty>No pending payments.</Empty>}<footer><span>{pendingBills.length} bill{pendingBills.length === 1 ? '' : 's'} with pending collection</span><div><button type="button" onClick={onViewAllPendingBills}>View all</button><button type="button" onClick={onOpenOfflineBills}>Open offline bills</button></div></footer></div><div className="panel stock-health"><div className="stock-health-head"><div><p className="eyebrow">INVENTORY CONTROL</p><h2>Stock health</h2></div><button>Manage stock</button></div><div className="stock-health-stats"><div><b>{activeProducts}</b><span>Active products</span></div><div><b>{inactiveProducts}</b><span>Inactive products</span></div><div><b>{lowStockProducts}</b><span>Low stock</span></div><div><b>{outOfStockProducts}</b><span>Out of stock</span></div></div>{stockAttentionProducts.length ? <div className="stock-attention-list">{stockAttentionProducts.map(product => { const outOfStock = Number(product.stock) <= 0; return <div className="stock-attention-item" key={product.id}><div><b>{product.name}</b><span>{product.category || 'General'}</span></div><div><strong>{outOfStock ? 'Out of stock' : `${neatNumber(product.stock)} left`}</strong><small className={outOfStock ? 'out' : 'low'}>{outOfStock ? 'Restock now' : 'Low stock'}</small></div></div> })}</div> : <p className="stock-health-message">All products are sufficiently stocked.</p>}</div></section><section className="dashboard-danger-zone"><div><p>DATA MANAGEMENT</p><h2>Reset this app</h2><span>Deletes all products, categories, bills and dashboard records permanently.</span></div><Button type="button" className="clear-all-data" onClick={clearAllData}><Trash2 size={16} />Clear all data</Button></section></>;
}

const blankProduct = { barcode: '', name: '', category: 'General', quantity: '1', unit: 'Pcs', mrp: '', purchasePrice: '', discount: '', sellingPrice: '' };
function Products() {
  const { data, setData, setNotice } = useStore(); const [form, setForm] = useState(blankProduct), [editing, setEditing] = useState(null), [query, setQuery] = useState(''), [showForm, setShowForm] = useState(false), [showUpload, setShowUpload] = useState(false), [uploadReport, setUploadReport] = useState(null);
  useEffect(() => { if (!showForm) return; const select = [...document.querySelectorAll('.product-editor select')].find(s => s.closest('.field')?.querySelector('span')?.textContent === 'Unit'); if (select) { select.innerHTML = '<option value="kg">kg</option><option value="g">g</option><option value="dozen">dozen</option><option value="Pcs">Pcs</option><option value="pack">pack</option><option value="L">L</option>'; select.value = ['pcs', 'piece'].includes(form.unit) ? 'Pcs' : form.unit || 'Pcs' } }, [showForm, form.unit]);
  const change = (k, v) => setForm(f => { let x = { ...f, [k]: v }; if (k === 'discount') { x.manual = false; x.sellingPrice = neatNumber(Math.max(0, (+x.mrp || 0) - (+x.discount || 0))) } else if (k === 'mrp' && !f.manual) x.sellingPrice = neatNumber(Math.max(0, (+x.mrp || 0) - (+x.discount || 0))); return x });
  useEffect(() => { if (!showForm || !form.manual || form.mrp === '' || form.sellingPrice === '') return; const discount = neatNumber(Math.max(0, (+form.mrp || 0) - (+form.sellingPrice || 0))); if (String(form.discount) !== discount) setForm(f => ({ ...f, discount })) }, [showForm, form.manual, form.mrp, form.sellingPrice, form.discount]);
  async function save(e) {
    e.preventDefault();
    if (!form.name || form.mrp === '') return setNotice({ type: 'error', text: 'Please complete all required product fields.' });
    const barcodeKey = form.barcode.trim().toLowerCase();
    if (barcodeKey && data.products.some(product => String(product.barcode || '').trim().toLowerCase() === barcodeKey && product.id !== editing)) return setNotice({ type: 'error', text: 'This barcode already belongs to a product.' });
    const product = { ...form, category: form.category || 'General', id: editing || uid(), mrp: +form.mrp, purchasePrice: form.purchasePrice === '' ? '' : +form.purchasePrice, discount: +form.discount || 0, sellingPrice: +form.sellingPrice || +form.mrp, quantity: +form.quantity || 1 };
    const finishSave = saved => { setData(current => ({ ...current, products: editing ? current.products.map(item => item.id === editing ? saved : item) : [saved, ...current.products] })); setNotice({ type: 'success', placement: editing ? 'bottom' : '', text: editing ? 'Product updated successfully.' : 'Product added successfully.' }); setForm(blankProduct); setEditing(null); setShowForm(false); };
    // Close the editor immediately; a slow or unavailable API must not keep this page open.
    finishSave(product);
    fetch(`${apiUrl}/products`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(product) }).catch(() => {
      // Provider has already saved the update locally for offline/browser use.
    });
  }
  async function validateUpload(file) {
    try {
      const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const [header = [], ...sheetRows] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
      const columns = header.map(value => String(value).trim().toLowerCase());
      const required = ['product name'];
      const missingHeaders = required.filter(name => !columns.includes(name));
      if (missingHeaders.length) return setUploadReport({ fileName: file.name, headerError: `Missing required column${missingHeaders.length > 1 ? 's' : ''}: ${missingHeaders.join(', ')}.`, rows: [], valid: [], invalid: [] });
      const valueAt = (row, name) => String(row[columns.indexOf(name)] ?? '').trim();
      const discountColumn = columns.findIndex(name => name.startsWith('discount'));
      const existingByBarcode = new Map(data.products.filter(product => String(product.barcode || '').trim()).map(product => [String(product.barcode).trim().toLowerCase(), product]));
      const existingByName = new Map(data.products.map(product => [String(product.name || '').trim().toLowerCase(), product]));
      const seenProductKeys = new Set();
      const rows = sheetRows.reduce((results, row, index) => {
        if (row.every(cell => String(cell ?? '').trim() === '')) return results;
        const item = { barcode: valueAt(row, 'barcode'), name: valueAt(row, 'product name'), category: valueAt(row, 'category') || 'General', quantity: valueAt(row, 'quantity') || '1', unit: valueAt(row, 'unit') || 'pcs', mrp: valueAt(row, 'mrp'), purchasePrice: valueAt(row, 'purchasing price') || valueAt(row, 'purchase price'), discount: discountColumn >= 0 ? String(row[discountColumn] ?? '').trim() : '', sellingPrice: valueAt(row, 'selling price') };
        const errors = [];
        ['name'].forEach(field => { if (!item[field]) errors.push('Product name is required'); });
        const numberFields = [['mrp', 'MRP'], ['purchasePrice', 'Purchasing price'], ['quantity', 'Quantity'], ['discount', 'Discount'], ['sellingPrice', 'Selling price']];
        numberFields.forEach(([field, label]) => { if (item[field] !== '' && (!Number.isFinite(Number(item[field])) || Number(item[field]) < 0)) errors.push(`${label} must be a positive number`); });
        const hasMrp = item.mrp !== '', hasDiscount = item.discount !== '', hasPrice = item.sellingPrice !== '';
        if (hasMrp && hasPrice) item.discount = String(Number(item.mrp) - Number(item.sellingPrice));
        else if (hasMrp && hasDiscount) item.sellingPrice = String(Number(item.mrp) - Number(item.discount));
        else if (hasDiscount && hasPrice) item.mrp = String(Number(item.discount) + Number(item.sellingPrice));
        else errors.push('Provide any two values: MRP, Discount, or Selling price');
        if (Number(item.discount) > Number(item.mrp)) errors.push('Discount cannot be greater than MRP');
        if (Number(item.sellingPrice) > Number(item.mrp)) errors.push('Selling price cannot be greater than MRP');
        const barcodeKey = item.barcode.toLowerCase(), nameKey = item.name.toLowerCase();
        const existingProduct = (barcodeKey && existingByBarcode.get(barcodeKey)) || existingByName.get(nameKey);
        const productKey = barcodeKey ? `barcode:${barcodeKey}` : `name:${nameKey}`;
        if (seenProductKeys.has(productKey)) errors.push('This product appears more than once in this file');
        seenProductKeys.add(productKey);
        const product = { ...item, id: existingProduct?.id || uid(), quantity: Number(item.quantity), mrp: Number(item.mrp), purchasePrice: item.purchasePrice === '' ? '' : Number(item.purchasePrice), discount: Number(item.discount), sellingPrice: Number(item.sellingPrice) };
        results.push({ rowNumber: index + 2, product: item.name || 'Unnamed product', barcode: item.barcode || '—', errors, operation: existingProduct ? 'update' : 'add', productData: product });
        return results;
      }, []);
      const valid = rows.filter(row => !row.errors.length);
      setUploadReport({ fileName: file.name, rows, valid, invalid: rows.filter(row => row.errors.length), updates: valid.filter(row => row.operation === 'update'), additions: valid.filter(row => row.operation === 'add') });
    } catch { setUploadReport({ fileName: file.name, headerError: 'We could not read this file. Please upload a valid CSV, XLS, or XLSX file.', rows: [], valid: [], invalid: [] }); }
  }
  async function importValidRows() {
    if (!uploadReport?.valid.length) return;
    const imported = uploadReport.valid.map(row => row.productData);
    const updateIds = new Set(uploadReport.updates.map(row => row.productData.id));
    // Update the local catalogue immediately, then persist every imported row so
    // products are still present after the application is reopened.
    setData(current => ({ ...current, products: [...imported.filter(product => !updateIds.has(product.id)), ...current.products.map(product => imported.find(entry => entry.id === product.id) || product)] }));
    setNotice({ type: 'success', text: `${uploadReport.additions.length} added, ${uploadReport.updates.length} product${uploadReport.updates.length === 1 ? '' : 's'} updated successfully.` });
    setUploadReport(null);
    setShowUpload(false);
    await Promise.allSettled(imported.map(product => fetch(`${apiUrl}/products`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(product) })));
  }
  const list = data.products.filter(p => !query || `${p.name} ${p.barcode} ${p.category}`.toLowerCase().includes(query.toLowerCase()));
  function exportProducts() { const rows = [['Product', 'Category', 'Selling price', 'Purchasing price', 'MRP', 'Unit', 'Barcode'], ...data.products.map(p => [p.name, p.category, p.sellingPrice, p.purchasePrice ?? '', p.mrp, p.unit, p.barcode])]; const csv = rows.map(r => r.map(v => `"${String(v ?? '').replaceAll('"', '""')}"`).join(',')).join('\n'); const link = document.createElement('a'); link.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv' })); link.download = 'products.csv'; link.click(); URL.revokeObjectURL(link.href) }
  const closeForm = () => { setForm(blankProduct); setEditing(null); setShowForm(false) };
  const deleteProduct = async product => {
    try {
      const response = await fetch(`${apiUrl}/products/${encodeURIComponent(product.id)}`, { method: 'DELETE' });
      if (!response.ok) throw new Error('Could not delete product');
      setData(current => ({ ...current, products: current.products.filter(item => item.id !== product.id) }));
      setNotice({ type: 'success', text: `${product.name} deleted successfully.` });
    } catch {
      setNotice({ type: 'error', text: `Could not delete ${product.name}. Please try again.` });
    }
  };
  return <><header><div><p className="eyebrow">INVENTORY</p><h1>Products</h1><p>Search, add and manage products in your inventory.</p></div></header>{showUpload && <BulkUploadModal onClose={() => { setShowUpload(false); setUploadReport(null) }} onFile={validateUpload} report={uploadReport} onImport={importValidRows} />}{showForm && <div className="panel product-editor"><div className="panel-head"><div><h2>{editing ? 'Edit product' : 'Add product'}</h2><p>Fields marked with * are required.</p></div><button className="icon-btn" onClick={closeForm}><X size={18} /></button></div><form className="form-grid" onSubmit={save}><Field label="Barcode *"><input value={form.barcode} onChange={e => change('barcode', e.target.value)} /></Field><Field label="Product name *"><input value={form.name} onChange={e => change('name', e.target.value)} /></Field><Field label="Category *"><select value={form.category} onChange={e => change('category', e.target.value)}><option value="">Select category</option>{data.categories.map(c => <option key={c.id}>{c.name}</option>)}</select></Field><Field label="Stock *"><input type="number" min="0" value={form.stock} onChange={e => change('stock', e.target.value)} /></Field><Field label="Unit"><select value={form.unit} onChange={e => change('unit', e.target.value)}><option>pcs</option><option>kg</option><option>g</option><option>ltr</option><option>pack</option></select></Field><Field label="MRP (₹) *"><input type="number" min="0" step=".01" value={form.mrp} onChange={e => change('mrp', e.target.value)} /></Field><Field label="Discount (₹)"><input type="number" min="0" step=".01" value={form.discount} onChange={e => change('discount', e.target.value)} /></Field><Field label="Selling price (₹)"><input type="number" min="0" step=".01" value={form.sellingPrice} onChange={e => setForm(f => ({ ...f, sellingPrice: e.target.value, manual: true }))} /></Field><Field label="Purchase price (₹)"><input type="number" min="0" step=".01" value={form.purchasePrice} onChange={e => change('purchasePrice', e.target.value)} /></Field><div className="actions"><Button><Plus size={16} />{editing ? 'Update product' : 'Add product'}</Button><Button type="button" variant="secondary" onClick={closeForm}>Cancel</Button></div></form></div>}<div className="catalogue"><div className="product-toolbar"><div className="product-search"><Search size={18} /><input placeholder="Search for products, categories..." value={query} onChange={e => setQuery(e.target.value)} /><Button type="button" onClick={() => { }}>Search</Button></div><div className="catalogue-actions"><Button type="button" variant="secondary" onClick={exportProducts}>Export Excel</Button><Button type="button" variant="secondary" onClick={() => setShowUpload(true)}>Bulk upload</Button><Button type="button" onClick={() => { setForm(blankProduct); setEditing(null); setShowForm(true) }}><Plus size={16} />Add product</Button></div></div><div className="product-list"><ProductTable products={list} onEdit={p => { setForm({ ...p }); setEditing(p.id); setShowForm(true); window.scrollTo(0, 0) }} onDelete={deleteProduct} /></div></div></>
}
function ProductTable({ products, onEdit, onDelete }) { return products.length ? <div className="table-wrap product-table"><table><thead><tr><th>Product</th><th>Category</th><th>Price</th><th>Stock</th><th>Actions</th></tr></thead><tbody>{products.map(p => <tr key={p.id}><td><div className="product-cell"><span><b>{p.name}</b><small>{p.quantity || 1} {p.unit}</small></span></div></td><td>{p.category || 'Uncategorized'}</td><td>Rs. {Number(p.sellingPrice).toFixed(0)}<small>MRP Rs. {Number(p.mrp).toFixed(0)}</small></td><td>{p.stock} units</td><td className="text-actions"><button onClick={() => onEdit(p)}>Edit</button><button onClick={() => onDelete(p)}>Delete</button></td></tr>)}</tbody></table></div> : <Empty>No products found. Click “Add product” to add your first item.</Empty> }

function Categories() {
  const { data, setData, setNotice } = useStore(); const [editingId, setEditingId] = useState(null), [showDialog, setShowDialog] = useState(false), [viewingCategory, setViewingCategory] = useState(null);
  const editingCategory = data.categories.find(category => category.id === editingId);
  const close = () => { setShowDialog(false); setEditingId(null); };
  return <div className="category-page"><div className="category-toolbar"><p><b>{data.categories.length}</b> active categories</p><Button onClick={() => setShowDialog(true)}><Plus size={16} />Add category</Button></div>{showDialog && <CategoryEditorDialog category={editingCategory} data={data} setData={setData} setNotice={setNotice} onClose={close} />}{viewingCategory && <CategoryProductsDialog category={viewingCategory} products={data.products.filter(product => product.category === viewingCategory.name)} onClose={() => setViewingCategory(null)} />}<div className="category-list">{data.categories.length ? <div className="table-wrap"><table><thead><tr><th>Category name</th><th>Products</th><th>View products</th><th>Actions</th></tr></thead><tbody>{data.categories.map(category => { const count = data.products.filter(product => product.category === category.name).length; const protectedCategory = isGeneralCategory(category); return <tr key={category.id}><td><b>{category.name}</b></td><td>{count}</td><td><button className="view-category" onClick={() => setViewingCategory(category)}>View products</button></td><td className="text-actions"><button onClick={() => { setEditingId(category.id); setShowDialog(true); }}>Edit</button><button disabled={protectedCategory} title={protectedCategory ? 'General is the default category and cannot be deleted.' : undefined} onClick={() => { if (protectedCategory) return; setData(current => ({ ...current, categories: current.categories.filter(item => item.id !== category.id), products: current.products.map(product => product.category === category.name ? { ...product, category: 'General' } : product) })); setNotice({ type: 'success', text: 'Category deleted.' }); }}>Delete</button></td></tr> })}</tbody></table></div> : <Empty>No categories yet. Click “Add category” to create your first one.</Empty>}</div></div>
}

function CategoryProductsDialog({ category, products, onClose }) {
  const [query, setQuery] = useState('');
  const matchingProducts = products.filter(product => `${product.name} ${product.barcode || ''}`.toLowerCase().includes(query.toLowerCase()));
  return <div className="category-overlay category-products-overlay" role="dialog" aria-modal="true" aria-label={`${category.name} products`}><div className="category-editor category-products-dialog"><button className="category-close" type="button" onClick={onClose} aria-label="Close"><X size={18} /></button><h2>{category.name}</h2><div className="category-product-search"><input autoFocus placeholder="Search products in this category" value={query} onChange={event => setQuery(event.target.value)} /><span><Search size={18} /></span></div><div className="category-products-results">{matchingProducts.length ? matchingProducts.map(product => <article className="category-product-card" key={product.id}><div className="category-product-main"><b>{product.name}</b><span>{product.quantity || 1} {product.unit || 'Pcs'} · Rs. {Number(product.sellingPrice || 0).toFixed(0)}</span></div><strong className={Number(product.stock) > 0 ? 'in-stock' : 'out-of-stock'}>{Number(product.stock || 0)} in stock</strong></article>) : <Empty>{products.length ? 'No products match your search.' : 'No products have been added to this category yet.'}</Empty>}</div><div className="category-products-actions"><Button type="button" onClick={onClose}>Done</Button></div></div></div>
}

function CategoryEditorDialog({ category, data, setData, setNotice, onClose }) {
  const isEdit = Boolean(category); const [name, setName] = useState(category?.name || ''), [addQuery, setAddQuery] = useState(''), [removeQuery, setRemoveQuery] = useState('');
  const inCategory = data.products.filter(product => product.category === category?.name);
  const available = data.products.filter(product => product.category !== category?.name && product.name.toLowerCase().includes(addQuery.toLowerCase()));
  const assigned = inCategory.filter(product => product.name.toLowerCase().includes(removeQuery.toLowerCase()));
  const save = e => { e.preventDefault(); const value = name.trim(); if (!value) return; if (data.categories.some(item => item.name.toLowerCase() === value.toLowerCase() && item.id !== category?.id)) return setNotice({ type: 'error', text: 'A category with this name already exists.' }); setData(current => ({ ...current, categories: isEdit ? current.categories.map(item => item.id === category.id ? { ...item, name: value } : item) : [...current.categories, { id: uid(), name: value }], products: isEdit ? current.products.map(product => product.category === category.name ? { ...product, category: value } : product) : current.products })); setNotice({ type: 'success', text: isEdit ? 'Category updated successfully.' : 'Category saved successfully.' }); onClose(); };
  const assign = product => setData(current => ({ ...current, products: current.products.map(item => item.id === product.id ? { ...item, category: category.name } : item) }));
  const remove = product => setData(current => ({ ...current, products: current.products.map(item => item.id === product.id ? { ...item, category: 'General' } : item) }));
  return <div className={`category-overlay ${isEdit ? 'category-edit-overlay' : ''}`} role="dialog" aria-modal="true"><div className={`category-editor ${isEdit ? 'category-editor-edit' : ''}`}><button className="category-close" type="button" onClick={onClose} aria-label="Close"><X size={18} /></button><h2>{isEdit ? 'Edit category' : 'Add category'}</h2><form onSubmit={save}><label className="category-name-field"><span>Category name</span><input autoFocus placeholder="e.g. Frozen foods" value={name} onChange={e => setName(e.target.value)} /></label><div className="category-form-actions"><Button type="button" variant="secondary" onClick={onClose}>Cancel</Button><Button>{isEdit ? 'Save changes' : 'Save category'}</Button></div></form>{isEdit && <div className="category-product-manager"><section><h3>Add products to this category</h3><input placeholder="Search products to add" value={addQuery} onChange={e => setAddQuery(e.target.value)} />{available.map(product => <div className="category-product-row" key={product.id}><span><b>{product.name}</b><small>{product.quantity || 1} {product.unit} · Rs. {product.sellingPrice}</small></span><Button type="button" onClick={() => assign(product)}>Add</Button></div>)}</section><section><h3>Remove products from this category</h3><input placeholder="Search products in this category" value={removeQuery} onChange={e => setRemoveQuery(e.target.value)} />{assigned.map(product => <div className="category-product-row" key={product.id}><span><b>{product.name}</b><small>{product.quantity || 1} {product.unit} · Rs. {product.sellingPrice}</small></span><button type="button" className="remove-product" onClick={() => remove(product)}>Remove</button></div>)}</section></div>}</div></div>
}

function Billing() { const { data, setData, setNotice } = useStore(); const [cart, setCart] = useState([]), [q, setQ] = useState(''), [customer, setCustomer] = useState(''), [mobile, setMobile] = useState(''), [payment, setPayment] = useState('Cash'), [saved, setSaved] = useState(null); const searchQuery = q.trim().toLowerCase(); const results = data.products.filter(p => searchQuery && [p.name, p.barcode].some(value => String(value ?? '').toLowerCase().includes(searchQuery))).slice(0, 6); const subtotal = cart.reduce((s, x) => s + x.mrp * x.qty, 0), discount = cart.reduce((s, x) => s + x.discount * x.qty, 0), total = cart.reduce((s, x) => s + x.sellingPrice * x.qty, 0); function add(p) { if (p.stock < 1) return setNotice({ type: 'error', text: 'This product is out of stock.' }); let found = cart.find(x => x.id === p.id); if (found && found.qty >= p.stock) return setNotice({ type: 'error', text: 'Cannot exceed available stock.' }); setCart(c => found ? c.map(x => x.id === p.id ? { ...x, qty: x.qty + 1 } : x) : [...c, { ...p, qty: 1 }]); setQ('') } function qty(id, delta) { setCart(c => c.map(x => x.id === id ? { ...x, qty: Math.max(1, Math.min(x.stock, x.qty + delta)) } : x)) } function newBill() { setCart([]); setCustomer(''); setMobile(''); setSaved(null) } function save() { if (!cart.length) return setNotice({ type: 'error', text: 'Add at least one item to the bill.' }); let bill = { id: uid(), number: billNumber(data), createdAt: new Date().toISOString(), customerName: customer, mobile, payment, items: cart.map(x => ({ productId: x.id, productName: x.name, barcode: x.barcode, quantity: x.qty, unit: x.unit, mrp: x.mrp, discount: x.discount, sellingPrice: x.sellingPrice, total: x.sellingPrice * x.qty })), subtotal, totalDiscount: discount, grandTotal: total }; setData(d => ({ ...d, bills: [bill, ...d.bills], products: d.products.map(p => { let x = cart.find(i => i.id === p.id); return x ? { ...p, stock: p.stock - x.qty } : p }) })); setSaved(bill); setCart([]); setNotice({ type: 'success', text: `Bill ${bill.number} saved and stock updated.` }) } function receipt() { let b = saved; if (!b) return setNotice({ type: 'error', text: 'Save the bill before printing or downloading.' }); let rows = b.items.map(i => `<tr><td>${i.productName}<br><small>${i.barcode}</small></td><td>${i.quantity}</td><td>â‚¹${i.sellingPrice.toFixed(2)}</td><td>â‚¹${i.total.toFixed(2)}</td></tr>`).join(''); return `<html><title>${b.number}</title><style>body{font-family:Arial;max-width:700px;margin:35px auto;color:#111}h1{text-align:center;margin:0}p{text-align:center;color:#555}table{width:100%;border-collapse:collapse;margin-top:20px}th,td{padding:10px;border-bottom:1px solid #ddd;text-align:left}.total{text-align:right;font-size:20px;font-weight:bold;margin-top:16px}</style><h1>Delhi Canteen</h1><p>Shop address & contact details<br>Bill ${b.number} Â· ${new Date(b.createdAt).toLocaleString('en-IN')}</p><p>${b.customerName ? `Customer: ${b.customerName} ${b.mobile ? `(${b.mobile})` : ''}` : 'Prabhu Shri Ram'} Â· ${b.payment}</p><table><thead><tr><th>Item</th><th>Qty</th><th>Price</th><th>Total</th></tr></thead><tbody>${rows}</tbody></table><p class="total">Grand total: â‚¹${b.grandTotal.toFixed(2)}</p><p>Thank you for shopping with Delhi Canteen!</p></html>` } function print() { let w = window.open('', '_blank'); w.document.write(receipt()); w.document.close(); w.print() } function download() { let a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([receipt()], { type: 'text/html' })); a.download = `${saved.number}-receipt.html`; a.click(); URL.revokeObjectURL(a.href) } return <><header><div><p className="eyebrow">POINT OF SALE</p><h1>Create bill</h1><p>Search by product name or scan a barcode to add items.</p></div><div className="header-actions"><Button variant="secondary" onClick={newBill}><Plus size={16} />New bill</Button><Button onClick={save}><FileText size={16} />Save bill</Button></div></header><div className="billing-grid"><div><div className="panel"><div className="search large"><Search size={19} /><input autoFocus placeholder="Search product or scan barcodeâ€¦" value={q} onChange={e => setQ(e.target.value)} /></div>{q && <div className="suggestions">{results.map(p => <button key={p.id} onClick={() => add(p)}><span><b>{p.name}</b><small>{p.barcode} Â· {p.stock} available</small></span><b>{money(p.sellingPrice)}</b></button>)}{!results.length && <Empty>No matching products.</Empty>}</div>}</div><div className="panel bill-table"><h2>Current bill</h2>{cart.length ? <div className="table-wrap"><table><thead><tr><th>Product</th><th>Qty</th><th>Price</th><th>Total</th><th></th></tr></thead><tbody>{cart.map(x => <tr key={x.id}><td><b>{x.name}</b><small>{x.barcode} Â· {x.stock} in stock</small></td><td><div className="qty"><button onClick={() => qty(x.id, -1)}><Minus size={13} /></button><b>{x.qty}</b><button onClick={() => qty(x.id, 1)}><Plus size={13} /></button></div></td><td>{money(x.sellingPrice)}<small>{money(x.discount)} off</small></td><td><b>{money(x.sellingPrice * x.qty)}</b></td><td><button className="icon-btn danger" onClick={() => setCart(c => c.filter(i => i.id !== x.id))}><X size={16} /></button></td></tr>)}</tbody></table></div> : <Empty><ShoppingCart size={28} /><br />Your bill is empty.</Empty>}</div></div><aside className="panel checkout"><h2>Customer details</h2><Field label="Customer name (optional)"><input value={customer} onChange={e => setCustomer('') } /></Field></aside></div></> }

function Customers() {
  const { data, setData, setNotice } = useStore();
  const [query, setQuery] = useState(''), [editing, setEditing] = useState(null), [showForm, setShowForm] = useState(false), [showUpload, setShowUpload] = useState(false);
  const customers = data.customers || [];
  useEffect(() => { const highest = customers.reduce((max, customer) => Math.max(max, Number(String(customer.customerId || '').match(/\d+$/)?.[0]) || 0), 0); if (Number(data.nextCustomerNumber || 1) <= highest) setData(current => ({ ...current, nextCustomerNumber: highest + 1 })); }, [customers, data.nextCustomerNumber, setData]);
  let generatedCustomerNumber = Math.max(Number(data.nextCustomerNumber || 1) - 1, customers.reduce((max, customer) => Math.max(max, Number(String(customer.customerId || '').match(/\d+$/)?.[0]) || 0), 0));
  const nextId = () => `DC-${String(++generatedCustomerNumber).padStart(5, '0')}`;
  const filtered = customers.filter(customer => `${customer.customerId} ${customer.name} ${customer.address} ${customer.phone}`.toLowerCase().includes(query.trim().toLowerCase()));
  const save = customer => { setData(current => { const customerNumber = Number(String(customer.customerId || '').match(/\d+$/)?.[0]) || 0; return { ...current, customers: editing ? current.customers.map(item => item.id === editing.id ? customer : item) : [customer, ...current.customers], nextCustomerNumber: editing ? current.nextCustomerNumber : Math.max(Number(current.nextCustomerNumber || 1), customerNumber + 1) }; }); setShowForm(false); setEditing(null); setNotice({ type: 'success', text: editing ? 'Customer updated successfully.' : 'Customer added successfully.' }); };
  const remove = customer => { setData(current => ({ ...current, customers: current.customers.filter(item => item.id !== customer.id) })); setNotice({ type: 'success', text: 'Customer deleted.' }); };
  const exportCustomers = () => { const rows = [['Customer ID', 'Customer name', 'Address', 'Phone no.'], ...customers.map(customer => [customer.customerId, customer.name, customer.address, customer.phone])]; const sheet = XLSX.utils.aoa_to_sheet(rows), workbook = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(workbook, sheet, 'Customers'); XLSX.writeFile(workbook, 'customers.xlsx'); };
  const downloadCustomerTemplate = () => { const sheet = XLSX.utils.aoa_to_sheet([['Customer name', 'Phone no.', 'Address']]), workbook = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(workbook, sheet, 'Customers'); XLSX.writeFile(workbook, 'customer-upload-template.xlsx'); };
  const importCustomers = async event => { const file = event.target.files?.[0]; event.target.value = ''; if (!file) return; try { const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array' }), sheet = workbook.Sheets[workbook.SheetNames[0]], rows = XLSX.utils.sheet_to_json(sheet, { defval: '' }); const imported = rows.map(row => ({ id: uid(), customerId: String(row['Customer ID'] || row['customer id'] || '').trim() || nextId(), name: String(row['Customer name'] || row['customer name'] || '').trim(), address: String(row.Address || row.address || '').trim(), phone: String(row['Phone no.'] || row['phone no.'] || row.Phone || row.phone || '').trim() })).filter(customer => customer.name); setData(current => ({ ...current, customers: [...imported, ...current.customers] })); setNotice({ type: 'success', text: `${imported.length} customer${imported.length === 1 ? '' : 's'} imported.` }); } catch { setNotice({ type: 'error', text: 'Could not read this customer file.' }); } };
  return <><header className="customers-top"><div><p className="eyebrow">CUSTOMERS</p><h1>Customers</h1><p>Every customer receives a unique DC ID.</p></div><div className="header-actions"><Button type="button" variant="secondary" onClick={exportCustomers}>Export Excel</Button><Button type="button" variant="secondary" onClick={() => setShowUpload(true)}>Bulk upload</Button><Button type="button" onClick={() => { setEditing(null); setShowForm(true); }}><Plus size={16} />Add customer</Button></div></header><section className="panel customers-panel"><div className="customers-toolbar"><h2>Customer accounts</h2><div className="customer-search"><input placeholder="Search ID, name, address or phone" value={query} onChange={event => setQuery(event.target.value)} /><Search size={19} /></div></div>{filtered.length ? <div className="customers-table-wrap"><table><thead><tr><th>CUSTOMER ID</th><th>CUSTOMER NAME</th><th>ADDRESS</th><th>PHONE NO.</th><th>ACTIONS</th></tr></thead><tbody>{filtered.map(customer => <tr key={customer.id}><td><b>{customer.customerId}</b></td><td>{customer.name}</td><td>{customer.address || '—'}</td><td>{customer.phone || '—'}</td><td className="text-actions"><button onClick={() => { setEditing(customer); setShowForm(true); }}>Edit</button><button onClick={() => remove(customer)}>Delete</button></td></tr>)}</tbody></table></div> : <Empty>No customers found. Add a customer or upload an Excel file.</Empty>}</section>{showForm && <CustomerForm customer={editing} customerId={editing?.customerId || nextId()} onClose={() => { setShowForm(false); setEditing(null); }} onSave={save} />}{showUpload && <CustomerUploadModal onClose={() => setShowUpload(false)} onDownloadTemplate={downloadCustomerTemplate} onFile={event => { setShowUpload(false); importCustomers(event); }} />}</>;
}

function CustomerForm({ customer, customerId, onClose, onSave }) {
  const [form, setForm] = useState(() => ({ customerId, name: customer?.name || '', address: customer?.address || '', phone: customer?.phone || '' }));
  const change = (key, value) => setForm(current => ({ ...current, [key]: value }));
  const save = event => { event.preventDefault(); if (!form.name.trim()) return; onSave({ ...customer, ...form, id: customer?.id || uid(), name: form.name.trim(), address: form.address.trim(), phone: form.phone.trim() }); };
  return <div className="customer-modal-overlay" role="dialog" aria-modal="true"><form className="customer-modal" onSubmit={save}><button type="button" className="bill-modal-close" onClick={onClose} aria-label="Close customer form">×</button><p className="eyebrow">{customer ? 'EDIT CUSTOMER' : 'NEW CUSTOMER'}</p><h2>{customer ? 'Edit customer' : 'Add customer'}</h2><div className="customer-form-grid"><Field label="Customer ID"><input value={form.customerId} readOnly /></Field><Field label="Customer name *"><input autoFocus required value={form.name} onChange={event => change('name', event.target.value)} /></Field><Field label="Address"><input value={form.address} onChange={event => change('address', event.target.value)} /></Field><Field label="Phone number"><input inputMode="tel" value={form.phone} onChange={event => change('phone', event.target.value)} /></Field></div><footer><Button type="button" variant="secondary" onClick={onClose}>Cancel</Button><Button type="submit">{customer ? 'Save changes' : 'Add customer'}</Button></footer></form></div>;
}

function CustomerUploadModal({ onClose, onDownloadTemplate, onFile }) {
  return <div className="customer-upload-overlay" role="dialog" aria-modal="true" aria-label="Bulk customer upload" onMouseDown={event => { if (event.target === event.currentTarget) onClose(); }}><section className="customer-upload-modal"><button type="button" className="customer-upload-close" onClick={onClose} aria-label="Close bulk customer upload">×</button><h2>Bulk customer upload</h2><p>Upload an Excel file with Customer Name and Phone No.</p><div className="customer-upload-actions"><Button type="button" variant="secondary" onClick={onDownloadTemplate}>Download Excel template</Button><label className="btn customer-file-button">Choose Excel file<input type="file" accept=".csv,.xlsx,.xls" onChange={onFile} /></label></div></section></div>;
}

function ShopProfile() {
  const { data, setData, setNotice } = useStore();
  const [profile, setProfile] = useState(() => ({ ...defaultShopProfile, ...(data.shopProfile || {}) }));
  const [backupFile, setBackupFile] = useState(null);
  const [restoringBackup, setRestoringBackup] = useState(false);
  const [creatingBackup, setCreatingBackup] = useState(false);
  const [automaticBackup, setAutomaticBackup] = useState(null);
  useEffect(() => setProfile({ ...defaultShopProfile, ...(data.shopProfile || {}) }), [data.shopProfile]);
  const loadAutomaticBackupStatus = () => fetch(`${apiUrl}/backups/status`).then(response => response.ok ? response.json() : Promise.reject()).then(setAutomaticBackup).catch(() => setAutomaticBackup({ unavailable: true }));
  useEffect(() => { loadAutomaticBackupStatus(); const timer = window.setInterval(loadAutomaticBackupStatus, 60_000); return () => window.clearInterval(timer); }, []);
  const change = (key, value) => setProfile(current => ({ ...current, [key]: value }));
  const save = event => { event.preventDefault(); setData(current => ({ ...current, shopProfile: { ...profile } })); setNotice({ type: 'success', text: 'Shop profile saved. New receipts will use these details.' }); };
  const createManualBackup = async () => { const now = new Date(); const timestamp = `${dateKey(now)}-${String(now.getHours()).padStart(2, '0')}-${String(now.getMinutes()).padStart(2, '0')}-${String(now.getSeconds()).padStart(2, '0')}`; setCreatingBackup(true); try { const response = await fetch(`${apiUrl}/backups`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ name: `delhi-canteen-manual-backup-${timestamp}`, data }) }); const result = await response.json(); if (!response.ok) throw new Error(result.error); setNotice({ type: 'success', text: `Manual backup created: ${result.filename}` }); } catch (error) { setNotice({ type: 'error', text: error.message || 'Could not create the manual backup.' }); } finally { setCreatingBackup(false); } };
  const restoreBackup = async event => { event.preventDefault(); if (!backupFile) return setNotice({ type: 'error', text: 'Choose a backup file first.' }); if (!window.confirm('Restore this backup? Current products, bills, customers and settings will be replaced.')) return; setRestoringBackup(true); try { const backup = JSON.parse(await backupFile.text()); const restoredData = backup?.data; if (!restoredData || typeof restoredData !== 'object' || !Array.isArray(restoredData.products)) throw new Error('Please choose a valid Delhi Canteen backup JSON file.'); const response = await fetch(`${apiUrl}/backups/restore`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(restoredData) }); const result = await response.json(); if (!response.ok) throw new Error(result.error); setData(restoredData); setBackupFile(null); setNotice({ type: 'success', text: 'Backup restored successfully. All saved data is now active.' }); } catch (error) { setNotice({ type: 'error', text: error.message || 'Could not restore this backup file.' }); } finally { setRestoringBackup(false); } };
  const automaticBackupStatus = automaticBackup?.unavailable ? 'Status is unavailable. Keep the app open once so its backup service can start.' : automaticBackup?.exists ? `Last updated: ${new Date(automaticBackup.updatedAt).toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' })}` : 'The first hourly backup will appear after the next clock hour.';
  return <><header><div><p className="eyebrow">SHOP SETTINGS</p><h1>Shop profile</h1><p>These details are shown on every bill receipt.</p></div></header><form className="panel shop-profile" onSubmit={save}><div className="shop-profile-grid"><Field label="Owner name"><input required value={profile.ownerName} onChange={event => change('ownerName', event.target.value)} /></Field><Field label="Phone number"><input required inputMode="tel" value={profile.phone} onChange={event => change('phone', event.target.value)} /></Field><Field label="Email ID"><input type="email" value={profile.email} onChange={event => change('email', event.target.value)} /></Field><Field label="Shop address"><textarea required rows="4" value={profile.address} onChange={event => change('address', event.target.value)} /></Field></div><div className="shop-profile-actions"><Button type="submit">Save profile</Button></div></form><section className="panel backup-settings"><div><p className="eyebrow">AUTOMATIC BACKUP</p><h2>Hourly backup is enabled</h2><p>Every hour, today’s backup file is updated with the latest products, bills, customers and settings.</p></div><div className="backup-status"><span className={automaticBackup?.exists ? 'backup-status-dot ready' : 'backup-status-dot'}></span><div><b>{automaticBackup?.filename || 'Preparing today’s backup file…'}</b><small>{automaticBackupStatus}</small></div></div><small>Saved in C:\Delhi Canteen Daily Backup. One file is kept for each date; it is updated every hour instead of creating duplicate files.</small></section><section className="panel backup-settings"><div><p className="eyebrow">MANUAL BACKUP</p><h2>Create backup now</h2><p>Save an additional timestamped copy immediately before making major changes.</p></div><div className="backup-controls"><small>Saved in C:\Delhi Canteen Daily Backup</small><Button type="button" onClick={createManualBackup} disabled={creatingBackup}>{creatingBackup ? 'Creating backup…' : 'Create manual backup'}</Button></div></section><form className="panel backup-settings" onSubmit={restoreBackup}><div><p className="eyebrow">DATA RESTORE</p><h2>Restore backup</h2><p>Choose a Delhi Canteen backup JSON file to restore products, categories, bills, customers, dashboard data and shop settings.</p></div><div className="backup-controls"><label className="backup-file-picker"><span>{backupFile ? backupFile.name : 'No backup file chosen'}</span><input type="file" accept=".json,application/json" onChange={event => setBackupFile(event.target.files?.[0] || null)} /><b>Choose backup file</b></label><Button type="submit" disabled={!backupFile || restoringBackup}>{restoringBackup ? 'Restoring…' : 'Restore backup'}</Button></div><small>Restoring replaces the current data. Files are saved in C:\Delhi Canteen Daily Backup.</small></form></>;
}

function PendingBillsPage({ onBack, onOpenOfflineBills }) {
  const { data, setData, setNotice } = useStore();
  const [updatingBillId, setUpdatingBillId] = useState(null);
  const pendingBills = (data.bills || []).filter(bill => Number(bill.remainingAmount || 0) > 0);
  const totalPending = pendingBills.reduce((sum, bill) => sum + Number(bill.remainingAmount || 0), 0);
  const updatingBill = (data.bills || []).find(bill => bill.id === updatingBillId);
  const updatePayment = (bill, paidAmount) => {
    const total = Number(bill.grandTotal || 0);
    const paid = Math.min(total, Math.max(0, Number(paidAmount) || 0));
    setData(current => ({ ...current, bills: current.bills.map(entry => entry.id === bill.id ? { ...entry, paidAmount: paid, remainingAmount: Math.max(0, total - paid) } : entry) }));
    setUpdatingBillId(null);
    setNotice({ type: 'success', text: paid >= total ? `Bill ${bill.number} marked as paid.` : `Payment for bill ${bill.number} updated.` });
  };
  return <><header><div><p className="eyebrow">PENDENCY</p><h1>All pending bills</h1><p>View every bill with a remaining payment balance.</p></div><div className="header-actions"><Button type="button" variant="secondary" onClick={onBack}>Back to dashboard</Button><Button type="button" onClick={onOpenOfflineBills}>Open offline bills</Button></div></header><section className="panel pending-bills-page"><div className="pending-bills-page-head"><div><h2>Pending payment details</h2><span>{pendingBills.length} pending bill{pendingBills.length === 1 ? '' : 's'}</span></div><strong>Total to collect: {money(totalPending)}</strong></div>{pendingBills.length ? <div className="billing-table-wrap"><table><thead><tr><th>BILL NO.</th><th>DATE &amp; TIME</th><th>CUSTOMER</th><th>PHONE</th><th>ADDRESS</th><th>PAYMENT</th><th>BILL TOTAL</th><th>PAID</th><th>REMAINING</th><th>UPDATE</th></tr></thead><tbody>{pendingBills.map(bill => <tr key={bill.id}><td><b>#{bill.number}</b></td><td>{new Date(bill.createdAt).toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' })}</td><td>{bill.customerName || 'Prabhu Shri Ram'}</td><td>{bill.mobile || '—'}</td><td>{bill.address || '—'}</td><td>{bill.payment || 'Cash'}</td><td>Rs. {Number(bill.grandTotal || 0).toFixed(2)}</td><td>Rs. {Number(bill.paidAmount || 0).toFixed(2)}</td><td><b>Rs. {Number(bill.remainingAmount || 0).toFixed(2)}</b></td><td><div className="pending-payment-actions"><button type="button" className="bill-edit" onClick={() => setUpdatingBillId(bill.id)}>Update</button><button type="button" className="mark-paid" onClick={() => updatePayment(bill, bill.grandTotal)}>Paid</button></div></td></tr>)}</tbody></table></div> : <Empty>No pending payments.</Empty>}</section>{updatingBill && <PaymentUpdateModal bill={updatingBill} onClose={() => setUpdatingBillId(null)} onSave={paidAmount => updatePayment(updatingBill, paidAmount)} />}</>;
}

function App() {
  const [page, setPage] = useState('Dashboard');
  const [billingTab, setBillingTab] = useState('new');
  const [open, setOpen] = useState(false);
  const [theme, setTheme] = useState(() => localStorage.getItem('dc-theme') === 'night' ? 'night' : 'day');
  const { notice } = useStore();
  const nav = [['Dashboard', LayoutDashboard], ['Add Products', PackagePlus], ['Categories', ClipboardList], ['Billing', ShoppingCart], ['Customers', ClipboardList], ['Shop Profile', StoreIcon]];
  useEffect(() => { document.body.classList.toggle('theme-night', theme === 'night'); localStorage.setItem('dc-theme', theme); }, [theme]);
  useEffect(() => {
    const handleShortcut = event => {
      // Billing tab switching needs to work even while a field is focused;
      // otherwise Ctrl+N / Ctrl+O would trigger the browser's New/Open action.
      if (event.ctrlKey && !event.altKey && !event.metaKey) {
        const key = event.key.toLowerCase();
        if (key === 'n' || key === 'o') {
          event.preventDefault();
          setBillingTab(key === 'n' ? 'new' : 'view');
          setPage('Billing');
          return;
        }
      }

      // Keep normal typing and shortcuts inside form fields untouched.
      const target = event.target;
      const isEditing = target instanceof HTMLElement && (target.isContentEditable || ['INPUT', 'TEXTAREA', 'SELECT'].includes(target.tagName));
      if (isEditing || !event.ctrlKey || event.altKey || event.metaKey) return;

      if (event.key.toLowerCase() === 'd') {
        event.preventDefault();
        setPage('Dashboard');
      }
      if (event.key.toLowerCase() === 'l') {
        event.preventDefault();
        setBillingTab('new');
        setPage('Billing');
      }
    };
    window.addEventListener('keydown', handleShortcut);
    return () => window.removeEventListener('keydown', handleShortcut);
  }, []);
  useEffect(() => {
    const closeTopDialogOnEscape = event => {
      if (event.key !== 'Escape') return;
      const dialogs = [...document.querySelectorAll('[role="dialog"], .product-quick-edit-overlay')];
      if (!dialogs.length) return;
      const topDialog = dialogs.reduce((top, dialog) => {
        const topLayer = Number.parseInt(window.getComputedStyle(top).zIndex, 10) || 0;
        const dialogLayer = Number.parseInt(window.getComputedStyle(dialog).zIndex, 10) || 0;
        return dialogLayer >= topLayer ? dialog : top;
      });
      const closeButton = topDialog.querySelector('.category-close, .bill-modal-close, .customer-upload-close, .upload-close, .product-quick-close, button[aria-label*="Close" i]');
      if (!closeButton) return;
      event.preventDefault();
      closeButton.click();
    };
    window.addEventListener('keydown', closeTopDialogOnEscape);
    return () => window.removeEventListener('keydown', closeTopDialogOnEscape);
  }, []);
  const toggleTheme = () => setTheme(current => current === 'day' ? 'night' : 'day');
  const Content = page === 'Dashboard' ? Dashboard : page === 'Pending Bills' ? PendingBillsPage : page === 'Add Products' ? Products : page === 'Categories' ? Categories : page === 'Customers' ? Customers : page === 'Shop Profile' ? ShopProfile : BillingWorkspace;
  const openOfflineBills = () => { setBillingTab('view'); setPage('Billing'); };
  const contentProps = page === 'Dashboard' ? { theme, onToggleTheme: toggleTheme, onOpenOfflineBills: openOfflineBills, onViewAllPendingBills: () => setPage('Pending Bills') } : page === 'Pending Bills' ? { onBack: () => setPage('Dashboard'), onOpenOfflineBills: openOfflineBills } : page === 'Billing' ? { initialTab: billingTab } : {};
  return <div className="app"><aside className={`sidebar ${open ? 'open' : ''}`}><div className="brand"><span>DC</span><div>Delhi Canteen<small>Billing software</small></div></div><nav>{nav.map(([n, I]) => <button key={n} className={page === n ? 'active' : ''} onClick={() => { setPage(n); if (n === 'Billing') setBillingTab('new'); setOpen(false) }}><I size={19} />{n}</button>)}</nav><div className="offline"><span></span> Made by Happy</div></aside><main><button className="mobile-menu" onClick={() => setOpen(!open)}><Menu /></button><Content {...contentProps} /></main>{notice && <div className={`toast ${notice.type} ${notice.placement || ''}`} role="status"><span className="toast-icon">{notice.type === 'error' ? <AlertTriangle size={16} /> : <CheckCircle2 size={16} />}</span><span>{notice.text}</span></div>}</div>
}
createRoot(document.getElementById('root')).render(<Provider><App /></Provider>);

function BillingPage() {
  const { data, setData, setNotice } = useStore();
  const empty = { barcode: '', name: '', category: '', quantity: '1', unit: '', mrp: '', discount: '', price: '' };
  const [cart, setCart] = useState([]), [customerId, setCustomerId] = useState(''), [customer, setCustomer] = useState(''), [mobile, setMobile] = useState(''), [payment, setPayment] = useState('Cash'), [paid, setPaid] = useState(''), [line, setLine] = useState(empty);
  const subtotal = cart.reduce((sum, item) => sum + item.mrp * item.qty, 0), discount = cart.reduce((sum, item) => sum + item.discount * item.qty, 0), rawTotal = cart.reduce((sum, item) => sum + item.sellingPrice * item.qty, 0), roundedTotal = roundBillTotal(rawTotal), total = roundedTotal;
  const change = (key, value) => setLine(current => ({ ...current, [key]: value }));
  const findProduct = value => { const key = value.trim().toLowerCase(); return data.products.find(item => String(item.barcode || '').trim().toLowerCase() === key || String(item.name || '').trim().toLowerCase() === key); };
  const chooseProduct = value => { const item = findProduct(value); if (!item) { if (!value.trim()) setLine(empty); return; } setLine({ barcode: item.barcode || '', name: item.name, category: item.category || 'General', quantity: '1', unit: item.unit || '', mrp: String(item.mrp || ''), discount: String(item.discount || 0), price: String(item.sellingPrice || item.mrp || '') }); };
  function addProduct() { const product = findProduct(line.barcode) || findProduct(line.name); if (!product) return setNotice({ type: 'error', text: 'Select a product using its barcode or exact name.' }); if (product.stock < 1) return setNotice({ type: 'error', text: 'This product is out of stock.' }); const quantity = Math.max(1, Number(line.quantity) || 1), current = cart.find(item => item.id === product.id); if ((current?.qty || 0) + quantity > product.stock) return setNotice({ type: 'error', text: 'Cannot exceed available stock.' }); const item = { ...product, category: line.category || product.category || 'General', unit: line.unit || product.unit, mrp: Number(line.mrp) || 0, discount: Number(line.discount) || 0, sellingPrice: Number(line.price) || Number(product.sellingPrice), qty: quantity }; setCart(items => current ? items.map(entry => entry.id === product.id ? { ...entry, qty: entry.qty + quantity } : entry) : [...items, item]); setLine(empty); }
  function clearBill() { setCart([]); setCustomerId(''); setCustomer(''); setMobile(''); setPaid(''); setLine(empty); }
  function saveBill() { if (!cart.length) return setNotice({ type: 'error', text: 'Add at least one item to the bill.' }); const bill = { id: uid(), number: billNumber(data), createdAt: new Date().toISOString(), customerId, customerName: customer, mobile, payment, items: cart.map(item => ({ productId: item.id, productName: item.name, barcode: item.barcode, quantity: item.qty, unit: item.unit, mrp: item.mrp, discount: item.discount, sellingPrice: item.sellingPrice, total: item.sellingPrice * item.qty })), subtotal, totalDiscount: discount, grandTotal: total }; setData(current => ({ ...current, bills: [bill, ...current.bills], products: current.products.map(product => { const item = cart.find(entry => entry.id === product.id); return item ? { ...product, stock: product.stock - item.qty } : product; }) })); setCart([]); setPaid(''); setNotice({ type: 'success', text: `Bill ${bill.number} saved and stock updated.` }); }
  return <div className="billing-page"><section className="panel billing-customer"><div className="billing-section-head"><h2>Customer details</h2><span>Customer ID is optional</span></div><div className="billing-customer-grid"><Field label="Customer ID"><input placeholder="Enter existing customer ID" value={customerId} onChange={e => setCustomerId(e.target.value)} /></Field><Field label="Customer name"><input placeholder="Enter Customer name." value={customer} onChange={e => setCustomer(e.target.value)} /></Field><Field label="Phone number"><input placeholder="Phone number" value={mobile} onChange={e => setMobile(e.target.value)} /></Field></div></section><section className="panel billing-add-product"><div className="billing-section-head"><h2>Add product</h2><span>Scan a barcode or search by name</span></div><div className="billing-product-grid"><Field label="Barcode"><input placeholder="Scan or enter barcode" value={line.barcode} onChange={e => { change('barcode', e.target.value); chooseProduct(e.target.value); }} /></Field><Field label="Product name"><input placeholder="Search product" value={line.name} onChange={e => { change('name', e.target.value); chooseProduct(e.target.value); }} /></Field><Field label="Category"><input placeholder="Enter category" value={line.category} onChange={e => change('category', e.target.value)} /></Field><Field label="Quantity"><input type="number" min="1" value={line.quantity} onChange={e => change('quantity', e.target.value)} /></Field><Field label="Unit"><input placeholder="Enter unit" value={line.unit} onChange={e => change('unit', e.target.value)} /></Field><Field label="MRP"><input type="number" min="0" value={line.mrp} onChange={e => change('mrp', e.target.value)} /></Field><Field label="Discount (₹)"><input type="number" min="0" value={line.discount} onChange={e => change('discount', e.target.value)} /></Field><Field label="Price"><input type="number" min="0" value={line.price} onChange={e => change('price', e.target.value)} /></Field><div className="billing-amount"><span>Amount</span><b>Rs. {((Number(line.price) || 0) * (Number(line.quantity) || 1)).toFixed(0)}</b></div><Button type="button" onClick={addProduct}>Add product</Button></div></section><section className="panel billing-items"><div className="billing-section-head"><h2>Bill items</h2><span>{cart.length} {cart.length === 1 ? 'item' : 'items'}</span></div><div className="billing-table-wrap"><table><thead><tr><th>S.NO.</th><th>PRODUCT NAME</th><th>QTY</th><th>UNIT</th><th>MRP</th><th>DISCOUNT</th><th>PRICE</th><th>AMOUNT</th><th>ACTION</th></tr></thead><tbody>{cart.map((item, index) => <tr key={item.id}><td>{index + 1}</td><td><b>{item.name}</b><small>{item.barcode}</small></td><td><input aria-label={`Quantity for ${item.name}`} type="number" min="1" max={item.stock} value={item.qty} onChange={e => setCart(items => items.map(entry => entry.id === item.id ? { ...entry, qty: Math.max(1, Math.min(entry.stock, Number(e.target.value) || 1)) } : entry))} /></td><td>{item.unit}</td><td>Rs. {item.mrp.toFixed(0)}</td><td>Rs. {item.discount.toFixed(0)}</td><td>Rs. {item.sellingPrice.toFixed(0)}</td><td><b>Rs. {(item.sellingPrice * item.qty).toFixed(0)}</b></td><td><button className="bill-remove" onClick={() => setCart(items => items.filter(entry => entry.id !== item.id))}>Remove</button></td></tr>)}</tbody></table>{!cart.length && <Empty>Add products to begin this bill.</Empty>}</div></section><section className="panel billing-summary"><div className="billing-summary-head"><h2>Bill summary</h2><Field label="Payment method"><select value={payment} onChange={e => setPayment(e.target.value)}><option>Cash</option><option>UPI</option><option>Card</option></select></Field></div><div className="summary-rows"><p><span>Total MRP</span><b>Rs. {subtotal.toFixed(2)}</b></p><p><span>Total discount</span><b>Rs. {discount.toFixed(2)}</b></p><p><span>Item total</span><b>Rs. {total.toFixed(2)}</b></p><hr /><p className="bill-total"><span>Total bill value</span><strong>Rs. {total.toFixed(2)}</strong></p><p><span>Remaining amount</span><b>Rs. {Math.max(0, total - (Number(paid) || 0)).toFixed(2)}</b></p></div><div className="paid-row"><Field label="Paid amount"><input type="number" min="0" value={paid} onChange={e => setPaid(e.target.value)} /></Field><Button type="button" variant="secondary" onClick={() => setPaid(total.toFixed(2))}>Paid</Button></div><div className="billing-actions"><Button type="button" variant="ghost" onClick={clearBill}>Clear bill</Button><Button type="button" onClick={saveBill}>Save &amp; print</Button></div></section></div>;
}

function BillingWorkspace({ initialTab = 'new' }) {
  const [tab, setTab] = useState(initialTab);
  useEffect(() => setTab(initialTab), [initialTab]);
  return <div className="billing-workspace"><div className="billing-tabs" role="tablist"><button type="button" title="Shortcut: Ctrl + N" aria-keyshortcuts="Control+N" className={tab === 'new' ? 'active' : ''} onClick={() => setTab('new')}>New bill</button><button type="button" title="Shortcut: Ctrl + O" aria-keyshortcuts="Control+O" className={tab === 'view' ? 'active' : ''} onClick={() => setTab('view')}>View bills</button></div>{tab === 'new' ? <BillingPageFixed /> : <BillsView />}</div>;
}

function BillingPageFixed() {
  const { data, setData, setNotice } = useStore();
  const blank = { barcode: '', name: '', category: '', quantity: '1', unit: '', mrp: '', purchasePrice: '', discount: '', price: '' };
  const [savedDraft] = useState(() => { try { return JSON.parse(sessionStorage.getItem('dc-billing-draft')) || {}; } catch { return {}; } });
  const [cart, setCart] = useState(savedDraft.cart || []), [customerId, setCustomerId] = useState(savedDraft.customerId || ''), [customer, setCustomer] = useState(savedDraft.customer || ''), [mobile, setMobile] = useState(savedDraft.mobile || ''), [address, setAddress] = useState(savedDraft.address || ''), [payment, setPayment] = useState(savedDraft.payment || 'Cash'), [paid, setPaid] = useState(savedDraft.paid || ''), [additional, setAdditional] = useState(savedDraft.additional || ''), [line, setLine] = useState(savedDraft.line || blank), [selectedProductId, setSelectedProductId] = useState(null), [addRequested, setAddRequested] = useState(false), [hideSuggestions, setHideSuggestions] = useState(false), [editingId, setEditingId] = useState(null), [barcodeDetectedAt, setBarcodeDetectedAt] = useState(0), [activeSuggestionIndex, setActiveSuggestionIndex] = useState(-1), [showHeldBills, setShowHeldBills] = useState(false);
  const [showShortcutHelp, setShowShortcutHelp] = useState(false);
  // Quantity is intentionally independent from whole-unit stock: items such
  // as weight-based products can be billed in fractions (for example, 0.5).
  useEffect(() => {
    const lineInput = document.querySelector('.billing-add-product input[placeholder="Quantity"]');
    // Fractional quantities remain available when typed manually, but the
    // spinner and keyboard arrows move in complete units for faster billing.
    if (lineInput) { lineInput.min = '0'; lineInput.step = '1'; }
    const inputs = [...document.querySelectorAll('.billing-items input[type="number"][min="1"]')];
    const cleanups = inputs.map((input, index) => {
      input.min = '0'; input.step = '1';
      const keepDecimalQuantity = () => {
        const quantity = Number(input.value);
        if (quantity <= 0 || quantity >= 1) return;
        // The pre-existing row handler clamps to whole units. Apply the
        // fractional value immediately after it has finished.
        window.setTimeout(() => setCart(items => items.map((item, itemIndex) => itemIndex === index ? { ...item, qty: quantity } : item)), 0);
      };
      input.addEventListener('input', keepDecimalQuantity, true);
      return () => input.removeEventListener('input', keepDecimalQuantity, true);
    });
    return () => cleanups.forEach(cleanup => cleanup());
  }, [cart, line.quantity]);
  useEffect(() => {
    const adjustQuantityByWholeUnit = event => {
      if (!event.target.matches('.billing-add-product input[placeholder="Quantity"]') || !['ArrowUp', 'ArrowDown'].includes(event.key)) return;
      event.preventDefault();
      const current = Number(event.target.value) || 0;
      const next = event.key === 'ArrowUp' ? Math.floor(current) + 1 : Math.max(0.01, Math.ceil(current) - 1);
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
      setter.call(event.target, String(next));
      event.target.dispatchEvent(new Event('input', { bubbles: true }));
    };
    document.addEventListener('keydown', adjustQuantityByWholeUnit, true);
    return () => document.removeEventListener('keydown', adjustQuantityByWholeUnit, true);
  }, []);
  useEffect(() => {
    const summary = document.querySelector('.billing-summary');
    const head = summary?.querySelector('.billing-summary-head');
    if (!summary || !head) return;
    let field = summary.querySelector('.additional-bill-input');
    if (!field) {
      field = document.createElement('label');
      field.className = 'field additional-bill-input';
      field.innerHTML = '<span>Additional amount</span><input type="number" step="0.01" placeholder="0.00" aria-label="Additional amount" />';
      head.after(field);
    }
    const input = field.querySelector('input');
    input.value = additional;
    input.oninput = event => setAdditional(event.target.value);
    input.onfocus = event => event.target.select();
    return () => { input.oninput = null; input.onfocus = null; };
  }, [additional]);
  useEffect(() => { const isEmptyDraft = !cart.length && !customerId && !customer && !mobile && !address && payment === 'Cash' && !paid && !additional && Object.values(line).every(value => value === '' || value === '1'); if (isEmptyDraft) { sessionStorage.removeItem('dc-billing-draft'); return; } sessionStorage.setItem('dc-billing-draft', JSON.stringify({ cart, customerId, customer, mobile, address, payment, paid, additional, line })); }, [cart, customerId, customer, mobile, address, payment, paid, additional, line]);
  // Keep every matching product available in the dropdown. The menu itself
  // scrolls, so a large catalogue does not hide results after the first few.
  const matches = line.name.trim() && !hideSuggestions ? data.products.filter(product => `${product.name} ${product.barcode || ''}`.toLowerCase().includes(line.name.trim().toLowerCase())) : [];
  const subtotal = cart.reduce((sum, item) => sum + item.mrp * item.qty, 0), discount = cart.reduce((sum, item) => sum + item.discount * item.qty, 0), itemTotal = cart.reduce((sum, item) => sum + item.sellingPrice * item.qty, 0), additionalAmount = Number(additional) || 0, rawTotal = itemTotal + additionalAmount, roundedTotal = roundBillTotal(rawTotal), total = roundedTotal;
  const change = (key, value) => {
    if (key === 'name') { setHideSuggestions(false); setActiveSuggestionIndex(-1); }
    if (key === 'name' || key === 'barcode') setSelectedProductId(null);
    // A previous selection populates both fields. Clear its paired value when
    // starting a new search, otherwise addProduct can resolve the stale barcode
    // (for example, Mahan ghee) instead of the newly entered product name.
    setLine(current => {
      // Removing either product identifier starts a fresh line, so none of the
      // details from the previously selected product can be billed by mistake.
      if ((key === 'name' || key === 'barcode') && !value.trim()) {
        return { ...blank, [key]: value };
      }
      const next = {
        ...current,
        [key]: value,
        ...(key === 'name' && value !== current.name ? { barcode: '' } : {}),
        ...(key === 'barcode' && value !== current.barcode ? { name: '' } : {}),
      };
      // Keep incomplete decimal values (such as "12." and ".5") in the
      // field. Converting them while the user is typing removes the decimal
      // point and makes fractional prices impossible to enter naturally.
      const completeAmount = /^\d*\.?\d+$/;
      if (!['mrp', 'discount', 'price'].includes(key) || !completeAmount.test(value)) return next;
      if (key === 'mrp' && current.discount !== '') {
        const mrp = Math.max(0, Number(value));
        next.price = neatNumber(Math.max(0, mrp - Number(current.discount)));
      }
      if (key === 'discount' && current.mrp !== '') {
        const discount = Math.min(Math.max(0, Number(value)), Number(current.mrp) || 0);
        next.discount = discount === Number(value) ? value : neatNumber(discount);
        next.price = neatNumber(Math.max(0, Number(current.mrp) - discount));
      }
      if (key === 'price' && current.mrp !== '') {
        const price = Math.min(Math.max(0, Number(value)), Number(current.mrp) || 0);
        next.price = price === Number(value) ? value : neatNumber(price);
        next.discount = neatNumber(Math.max(0, Number(current.mrp) - price));
      }
      return next;
    });
  };
  const focusLineQuantity = () => requestAnimationFrame(() => {
    const quantity = document.querySelector('.billing-add-product input[placeholder="Quantity"]');
    quantity?.focus();
    quantity?.select();
  });
  const selectProduct = (product, addNow = false, focusQuantity = true) => { setSelectedProductId(product.id); setLine({ barcode: product.barcode || '', name: product.name, category: product.category || 'General', quantity: '1', unit: product.unit || '', mrp: String(product.mrp || ''), purchasePrice: String(product.purchasePrice ?? ''), discount: String(product.discount || 0), price: String(product.sellingPrice || product.mrp || '') }); setHideSuggestions(true); if (addNow) setAddRequested(true); else if (focusQuantity) focusLineQuantity(); };
  const handleProductNameKeyDown = event => {
    if (!matches.length) return;
    if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
      event.preventDefault();
      event.stopPropagation();
      setActiveSuggestionIndex(index => event.key === 'ArrowDown' ? (index + 1) % matches.length : index < 0 ? matches.length - 1 : (index - 1 + matches.length) % matches.length);
    } else if (event.key === 'Enter' && activeSuggestionIndex >= 0) {
      event.preventDefault();
      // Do not let the grid's Enter handler add the product to the bill.
      // This Enter only confirms the highlighted autocomplete suggestion.
      event.stopPropagation();
      selectProduct(matches[activeSuggestionIndex]);
    } else if (event.key === 'Escape') {
      event.stopPropagation();
      setHideSuggestions(true);
      setActiveSuggestionIndex(-1);
    }
  };
  const handleBarcodeChange = value => {
    const barcode = value.trim();
    // An empty barcode must reset the line, never match a product whose
    // barcode is also blank.
    const product = barcode && data.products.find(item => String(item.barcode || '').trim().toLowerCase() === barcode.toLowerCase());
    // A scanner fills the line only. Its immediate trailing Enter is ignored so
    // the administrator can review or edit the details before adding.
    if (product) { setBarcodeDetectedAt(performance.now()); selectProduct(product, false, false); }
    else change('barcode', value);
  };
  useEffect(() => { const addSuggestions = (input, values, detail) => { if (!input) return () => {}; const field = input.closest('.field'); field.style.position = 'relative'; let menu; const close = () => menu?.remove(); const render = () => { close(); const query = input.value.trim().toLowerCase(); const options = values.filter(value => !query || value.toLowerCase().includes(query)); if (!options.length) return; menu = document.createElement('div'); menu.className = 'billing-product-suggestions billing-field-suggestions'; menu.setAttribute('role', 'listbox'); options.forEach(value => { const option = document.createElement('button'); option.type = 'button'; option.innerHTML = `<b>${value}</b><small>${detail}</small>`; option.addEventListener('mousedown', event => event.preventDefault()); option.addEventListener('click', () => { const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set; setter.call(input, value); input.dispatchEvent(new Event('input', { bubbles: true })); close(); input.focus(); }); menu.appendChild(option); }); field.appendChild(menu); }; const blur = () => setTimeout(close, 120); input.addEventListener('focus', render); input.addEventListener('input', render); input.addEventListener('blur', blur); return () => { close(); input.removeEventListener('focus', render); input.removeEventListener('input', render); input.removeEventListener('blur', blur); }; }; const category = addSuggestions(document.querySelector('.billing-add-product input[placeholder="Enter category"]'), data.categories.map(item => item.name), 'Category'); const unit = addSuggestions(document.querySelector('.billing-add-product input[placeholder="Enter unit"]'), ['piece', 'pcs', 'kg', 'g', 'litre', 'ml', 'pack', 'dozen'], 'Unit'); return () => { category(); unit(); }; }, [data.categories]);
  useEffect(() => { const units = new Set(data.products.map(product => product.unit).filter(Boolean)); const filterUnits = event => { if (!event.target.matches('.billing-add-product input[placeholder="Enter unit"]')) return; setTimeout(() => document.querySelectorAll('.billing-field-suggestions button').forEach(option => { if (!units.has(option.querySelector('b')?.textContent)) option.remove(); }), 0); }; document.addEventListener('focusin', filterUnits); document.addEventListener('input', filterUnits); return () => { document.removeEventListener('focusin', filterUnits); document.removeEventListener('input', filterUnits); }; }, [data.products]);
  useEffect(() => { const input = document.querySelector('.billing-add-product input[placeholder="Enter unit"]'); if (!input) return; const select = document.createElement('select'); select.className = 'billing-unit-select'; ['Pcs', 'kg', 'g', 'dozen', 'pack', 'L'].forEach(unit => { const option = document.createElement('option'); option.value = unit; option.textContent = unit; select.appendChild(option); }); select.value = ['pcs', 'piece'].includes(String(line.unit).toLowerCase()) ? 'Pcs' : line.unit || 'Pcs'; const updateUnit = () => { const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set; setter.call(input, select.value); input.dispatchEvent(new Event('input', { bubbles: true })); }; select.addEventListener('change', updateUnit); input.style.display = 'none'; input.after(select); return () => { select.removeEventListener('change', updateUnit); select.remove(); input.style.display = ''; }; }, [line.unit]);
  useEffect(() => {
    const amount = document.querySelector('.billing-add-product .billing-amount');
    if (!amount) return;
    let purchasePrice = amount.querySelector('.purchase-price');
    if (!purchasePrice) { purchasePrice = document.createElement('small'); purchasePrice.className = 'purchase-price'; amount.appendChild(purchasePrice); }
    const value = Number(line.purchasePrice);
    purchasePrice.textContent = line.purchasePrice !== '' && Number.isFinite(value) ? `Purchase price: Rs. ${value.toFixed(2)}` : '';
  }, [line.purchasePrice]);
  useEffect(() => { const onFieldKeys = event => { const input = event.target.matches('.billing-add-product input[placeholder="Enter category"], .billing-add-product input[placeholder="Enter unit"]') ? event.target : event.target.closest('.billing-field-suggestions')?.previousElementSibling; const menu = input?.closest('.field')?.querySelector('.billing-field-suggestions'); const options = menu ? [...menu.querySelectorAll('button')] : []; const index = options.indexOf(document.activeElement); if (!options.length || (document.activeElement !== input && index < 0)) return; if (event.key === 'ArrowDown' || event.key === 'ArrowUp') { event.preventDefault(); const next = event.key === 'ArrowDown' ? (index + 1) % options.length : index < 0 ? options.length - 1 : (index - 1 + options.length) % options.length; options[next].focus(); } else if (event.key === 'Enter' && index >= 0) { event.preventDefault(); options[index].click(); } else if (event.key === 'Escape') { event.preventDefault(); menu.remove(); input.focus(); } }; document.addEventListener('keydown', onFieldKeys); return () => document.removeEventListener('keydown', onFieldKeys); }, []);
  useEffect(() => { const selectFocusedOption = event => { const option = event.target.closest('.billing-field-suggestions button'); if (event.key === 'Enter' && option) { event.preventDefault(); event.stopPropagation(); option.click(); } }; document.addEventListener('keydown', selectFocusedOption, true); return () => document.removeEventListener('keydown', selectFocusedOption, true); }, []);
  useEffect(() => {
    const input = document.querySelector('.billing-add-product input[placeholder="Search product"]');
    if (!input) return;
    input.addEventListener('keydown', handleProductNameKeyDown);
    return () => input.removeEventListener('keydown', handleProductNameKeyDown);
  }, [matches, activeSuggestionIndex, editingId]);
  useEffect(() => {
    const options = [...document.querySelectorAll('.billing-add-product .billing-product-suggestions button')];
    options.forEach((option, index) => option.classList.toggle('active', index === activeSuggestionIndex));
    // Keep arrow-key (or hover) selection inside the suggestion list instead
    // of allowing the highlighted product to move below the visible area.
    // Key repeat can advance faster than a smooth animation. Use an immediate
    // nearest-row scroll so the visible list always stays aligned with the
    // highlighted product.
    options[activeSuggestionIndex]?.scrollIntoView({ block: 'nearest', behavior: 'auto' });
  }, [activeSuggestionIndex, matches]);
  useEffect(() => { const moveBetweenFields = event => { if (event.key !== 'Tab') return; const product = document.querySelector('.billing-add-product input[placeholder="Search product"]'); const barcode = document.querySelector('.billing-add-product input[placeholder="Scan or enter barcode"]'); const category = document.querySelector('.billing-add-product input[placeholder="Enter category"]'); const quantity = document.querySelector('.billing-add-product input[type="number"]'); const focusQuantity = () => { quantity?.focus(); requestAnimationFrame(() => quantity?.select()); }; if (event.shiftKey) { if (event.target === product) { event.preventDefault(); barcode?.focus(); } else if (event.target === category) { event.preventDefault(); product?.focus(); } else if (event.target === quantity) { event.preventDefault(); category?.focus(); } } else if (event.target === product) { event.preventDefault(); category?.focus(); } else if (event.target === category) { event.preventDefault(); focusQuantity(); } }; document.addEventListener('keydown', moveBetweenFields, true); return () => document.removeEventListener('keydown', moveBetweenFields, true); }, []);
  useEffect(() => {
    const focusLineField = event => {
      if (!event.ctrlKey || event.altKey || event.metaKey) return;
      const active = document.activeElement;
      const isEditing = active instanceof HTMLInputElement
        || active instanceof HTMLTextAreaElement
        || active instanceof HTMLSelectElement
        || active?.isContentEditable;
      if (isEditing) return;

      const shortcuts = {
        b: '.billing-add-product input[placeholder="Scan or enter barcode"]',
        p: '.billing-add-product input[placeholder="Search product"]',
        q: '.billing-add-product input[placeholder="Quantity"]',
        m: '.billing-add-product input[placeholder="MRP"]',
        c: '.billing-customer input[placeholder="Enter Customer name"]',
      };
      const field = document.querySelector(shortcuts[event.key.toLowerCase()]);
      if (!field) return;
      event.preventDefault();
      field.focus();
      field.select();
    };
    document.addEventListener('keydown', focusLineField);
    return () => document.removeEventListener('keydown', focusLineField);
  }, []);
  useEffect(() => { const closeSuggestions = event => { if (event.target.matches('.billing-add-product input[placeholder="Search product"]')) setHideSuggestions(false); else if (!event.target.closest('.billing-add-product .field:nth-child(2)')) setHideSuggestions(true); }; document.addEventListener('pointerdown', closeSuggestions); return () => document.removeEventListener('pointerdown', closeSuggestions); }, []);
  useEffect(() => { if (!addRequested) return; setAddRequested(false); addProduct(); }, [addRequested, line]);
  useEffect(() => { const edits = []; document.querySelectorAll('.billing-items .bill-remove').forEach((remove, index) => { const item = cart[index]; const edit = document.createElement('button'); edit.type = 'button'; edit.className = 'bill-edit'; edit.textContent = 'Edit'; edit.addEventListener('click', () => { if (!item) return; setLine({ barcode: item.barcode || '', name: item.name, category: item.category || 'General', quantity: String(item.qty), unit: item.unit || '', mrp: String(item.mrp), purchasePrice: String(item.purchasePrice ?? ''), discount: String(item.discount), price: String(item.sellingPrice) }); setEditingId(item.id); document.querySelector('.billing-add-product')?.scrollIntoView({ behavior: 'smooth', block: 'center' }); }); remove.parentElement?.insertBefore(edit, remove); edits.push(edit); }); return () => edits.forEach(edit => edit.remove()); }, [cart]);
  function updateEditedProduct() { const product = data.products.find(item => item.id === editingId); if (!product) return; const quantity = Math.max(0.01, Number(line.quantity) || 1); setCart(items => items.map(item => item.id === editingId ? { ...item, category: line.category || item.category || 'General', qty: quantity, unit: line.unit || item.unit, mrp: Number(line.mrp) || 0, purchasePrice: line.purchasePrice === '' ? item.purchasePrice : Number(line.purchasePrice), discount: Number(line.discount) || 0, sellingPrice: Number(line.price) || 0 } : item)); setEditingId(null); setLine(blank); setNotice({ type: 'success', text: 'Bill item updated.' }); }
  useEffect(() => {
    const button = document.querySelector('.billing-add-product .billing-product-grid > .btn');
    if (!button) return;
    button.textContent = editingId ? 'Update product' : 'Add product';
    const update = event => {
      if (!editingId) return;
      event.preventDefault();
      event.stopImmediatePropagation();
      updateEditedProduct();
    };
    button.addEventListener('click', update);
    return () => button.removeEventListener('click', update);
  }, [editingId, line, cart]);
  useEffect(() => {
    const header = document.querySelector('.billing-add-product .billing-section-head');
    if (!header || header.querySelector('.product-record-actions')) return;
    const actions = document.createElement('div');
    actions.className = 'product-record-actions';
    const createButton = document.createElement('button');
    createButton.type = 'button'; createButton.className = 'btn'; createButton.textContent = 'Add product';
    const editButton = document.createElement('button');
    editButton.type = 'button'; editButton.className = 'btn secondary edit-product-record'; editButton.textContent = 'Edit product';
    createButton.addEventListener('click', openNewProductRecord);
    editButton.addEventListener('click', updateProductRecord);
    actions.append(createButton, editButton); header.append(actions);
    return () => {
      createButton.removeEventListener('click', openNewProductRecord);
      editButton.removeEventListener('click', updateProductRecord);
      actions.remove();
    };
  }, [selectedProductId, line, data.products]);
  function addProduct() {
    const normalizedName = line.name.trim().toLowerCase();
    const product = data.products.find(item => item.id === selectedProductId)
      || data.products.find(item => line.barcode && String(item.barcode || '') === line.barcode);
    if (!normalizedName) return setNotice({ type: 'error', text: 'Enter a product name before adding the item.' });
    const quantity = Math.max(0.01, Number(line.quantity) || 1), inCart = product ? cart.find(item => item.id === product.id) : null;
    const item = product
      ? { ...product, name: line.name.trim(), category: line.category || product.category || 'General', unit: line.unit || product.unit, mrp: Number(line.mrp) || 0, discount: Number(line.discount) || 0, sellingPrice: Number(line.price) || Number(product.sellingPrice), qty: quantity }
      : { id: uid(), barcode: line.barcode.trim(), name: line.name.trim(), category: line.category || 'General', unit: line.unit || 'Pcs', mrp: Number(line.mrp) || 0, discount: Number(line.discount) || 0, sellingPrice: Number(line.price) || 0, qty: quantity };
    setCart(items => inCart ? items.map(entry => entry.id === product.id ? { ...entry, qty: entry.qty + quantity } : entry) : [item, ...items]);
    setSelectedProductId(null);
    setLine(blank);
    // Keep barcode entry ready for the next scan as soon as the current item
    // has been added to the bill list.
    requestAnimationFrame(() => document.querySelector('.billing-add-product input[placeholder="Scan or enter barcode"]')?.focus());
  }
  function openNewProductRecord() {
    const overlay = document.createElement('div');
    overlay.className = 'product-quick-edit-overlay';
    overlay.innerHTML = '<section class="product-quick-edit"><button class="product-quick-close" type="button">×</button><p>PRODUCT DETAILS</p><div class="product-quick-grid"></div><footer><button type="button" class="btn product-quick-save">Add product</button><button type="button" class="btn secondary product-quick-cancel">Cancel</button></footer></section>';
    const editor = overlay.querySelector('.product-quick-edit'), grid = editor.querySelector('.product-quick-grid');
    const values = { barcode: '', name: '', category: 'General', unit: 'Pcs', mrp: '', discount: '', sellingPrice: '', purchasePrice: '' };
    Object.entries({ barcode: 'Barcode *', name: 'Product name *', category: 'Category', unit: 'Unit', mrp: 'MRP', discount: 'Discount (Rs)', sellingPrice: 'Selling price', purchasePrice: 'Purchase price (₹)' }).forEach(([key, label]) => { const field = document.createElement('label'); field.textContent = label; const input = document.createElement(key === 'category' || key === 'unit' ? 'select' : 'input'); input.name = key; if (['mrp', 'discount', 'sellingPrice', 'purchasePrice'].includes(key)) { input.type = 'number'; input.min = '0'; input.step = '0.01'; } input.value = values[key]; field.append(input); grid.append(field); });
    const fields = Object.fromEntries([...grid.querySelectorAll('input,select')].map(input => [input.name, input]));
    data.categories.forEach(category => fields.category.add(new Option(category.name, category.name)));
    ['Pcs', 'kg', 'g', 'dozen', 'pack', 'L'].forEach(unit => fields.unit.add(new Option(unit, unit)));
    const syncPrices = changedField => {
      const mrp = Math.max(0, Number(fields.mrp.value) || 0);
      if (changedField === 'mrp') { const discount = Math.min(mrp, Math.max(0, Number(fields.discount.value) || 0)); fields.discount.value = neatNumber(discount); fields.sellingPrice.value = neatNumber(mrp - discount); }
      else if (changedField === 'discount') { if (!/^\d*\.?\d+$/.test(fields.discount.value)) return; const discount = Math.min(mrp, Math.max(0, Number(fields.discount.value) || 0)); if (discount !== Number(fields.discount.value)) fields.discount.value = neatNumber(discount); fields.sellingPrice.value = neatNumber(mrp - discount); }
      else { const sellingPrice = Math.min(mrp, Math.max(0, Number(fields.sellingPrice.value) || 0)); fields.sellingPrice.value = neatNumber(sellingPrice); fields.discount.value = neatNumber(mrp - sellingPrice); }
    };
    fields.mrp.addEventListener('input', () => syncPrices('mrp'));
    fields.discount.addEventListener('input', () => syncPrices('discount'));
    fields.sellingPrice.addEventListener('input', () => syncPrices('sellingPrice'));
    const close = () => overlay.remove();
    editor.querySelector('.product-quick-close').onclick = close; editor.querySelector('.product-quick-cancel').onclick = close;
    editor.querySelector('.product-quick-save').onclick = () => {
      const barcode = fields.barcode.value.trim(), name = fields.name.value.trim();
      if (!barcode || !name || fields.mrp.value === '') return setNotice({ type: 'error', text: 'Barcode, product name and MRP are required.' });
      const barcodeExists = data.products.some(item => String(item.barcode || '').trim().toLowerCase() === barcode.toLowerCase());
      if (barcodeExists) return setNotice({ type: 'error', text: 'This barcode already belongs to a product.' });
      syncPrices('sellingPrice');
      const product = { id: uid(), barcode, name, category: fields.category.value || 'General', unit: fields.unit.value || 'Pcs', quantity: 1, stock: 0, mrp: Number(fields.mrp.value), discount: Number(fields.discount.value) || 0, sellingPrice: Number(fields.sellingPrice.value) || 0, purchasePrice: fields.purchasePrice.value === '' ? '' : Number(fields.purchasePrice.value) };
      setData(current => ({ ...current, products: [product, ...current.products] }));
      fetch(`${apiUrl}/products`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(product) }).catch(() => {});
      selectProduct(product, false, false);
      setNotice({ type: 'success', text: 'Product added successfully. It is ready to add to this bill.' });
      close();
    };
    overlay.onclick = event => { if (event.target === overlay) close(); };
    document.body.append(overlay);
    requestAnimationFrame(() => fields.barcode.focus());
  }
  function updateProductRecord() {
    const product = data.products.find(item => item.id === selectedProductId)
      || data.products.find(item => line.barcode && String(item.barcode || '') === line.barcode)
      || data.products.find(item => String(item.name || '').trim().toLowerCase() === line.name.trim().toLowerCase());
    if (!product) return setNotice({ type: 'error', text: 'Select a product before editing it.' });
    const overlay = document.createElement('div');
    overlay.className = 'product-quick-edit-overlay';
    overlay.innerHTML = '<section class="product-quick-edit"><button class="product-quick-close" type="button">×</button><p>PRODUCT DETAILS</p><div class="product-quick-grid"></div><footer><button type="button" class="btn product-quick-save">Update product</button><button type="button" class="btn secondary product-quick-cancel">Cancel</button></footer></section>';
    const editor = overlay.querySelector('.product-quick-edit'), grid = editor.querySelector('.product-quick-grid');
    const values = { barcode: product.barcode || '', name: product.name || '', category: product.category || 'General', unit: product.unit || 'Pcs', mrp: product.mrp ?? '', discount: product.discount ?? 0, sellingPrice: product.sellingPrice ?? '', purchasePrice: product.purchasePrice ?? '' };
    Object.entries({ barcode: 'Barcode *', name: 'Product name *', category: 'Category', unit: 'Unit', mrp: 'MRP', discount: 'Discount (Rs)', sellingPrice: 'Selling price', purchasePrice: 'Purchase price (₹)' }).forEach(([key, label]) => { const field = document.createElement('label'); field.textContent = label; const input = document.createElement(key === 'category' || key === 'unit' ? 'select' : 'input'); input.name = key; if (['mrp', 'discount', 'sellingPrice', 'purchasePrice'].includes(key)) { input.type = 'number'; input.min = '0'; input.step = '0.01'; } input.value = values[key]; field.append(input); grid.append(field); });
    const fields = Object.fromEntries([...grid.querySelectorAll('input,select')].map(input => [input.name, input]));
    data.categories.forEach(category => fields.category.add(new Option(category.name, category.name)));
    fields.category.value = values.category;
    ['Pcs', 'kg', 'g', 'dozen', 'pack', 'L'].forEach(unit => fields.unit.add(new Option(unit, unit)));
    fields.unit.value = values.unit;
    // These three prices always describe the same relationship:
    // selling price = MRP - discount. Keep the product editor in sync no
    // matter which of the three fields the user changes.
    const syncPrices = changedField => {
      const mrp = Math.max(0, Number(fields.mrp.value) || 0);
      if (changedField === 'mrp') {
        const discount = Math.min(mrp, Math.max(0, Number(fields.discount.value) || 0));
        fields.discount.value = neatNumber(discount);
        fields.sellingPrice.value = neatNumber(mrp - discount);
      } else if (changedField === 'discount') {
        // A number input briefly holds "." / "0." while a decimal is being
        // entered. Do not turn that temporary value into 0 and block ".5".
        if (!/^\d*\.?\d+$/.test(fields.discount.value)) return;
        const discount = Math.min(mrp, Math.max(0, Number(fields.discount.value) || 0));
        if (discount !== Number(fields.discount.value)) fields.discount.value = neatNumber(discount);
        fields.sellingPrice.value = neatNumber(mrp - discount);
      } else {
        const sellingPrice = Math.min(mrp, Math.max(0, Number(fields.sellingPrice.value) || 0));
        fields.sellingPrice.value = neatNumber(sellingPrice);
        fields.discount.value = neatNumber(mrp - sellingPrice);
      }
    };
    fields.mrp.addEventListener('input', () => syncPrices('mrp'));
    fields.discount.addEventListener('input', () => syncPrices('discount'));
    fields.sellingPrice.addEventListener('input', () => syncPrices('sellingPrice'));
    const close = () => overlay.remove();
    editor.querySelector('.product-quick-close').onclick = close; editor.querySelector('.product-quick-cancel').onclick = close;
    editor.querySelector('.product-quick-save').onclick = () => {
      if (!fields.name.value.trim() || fields.mrp.value === '') return setNotice({ type: 'error', text: 'Product name and MRP are required.' });
      syncPrices('sellingPrice');
      const updatedLine = { barcode: fields.barcode.value.trim(), name: fields.name.value.trim(), category: fields.category.value, unit: fields.unit.value, mrp: String(Number(fields.mrp.value)), discount: String(Number(fields.discount.value) || 0), price: String(Number(fields.sellingPrice.value) || 0), purchasePrice: fields.purchasePrice.value === '' ? '' : String(Number(fields.purchasePrice.value)) };
      const updatedProduct = { ...product, barcode: updatedLine.barcode, name: updatedLine.name, category: updatedLine.category, unit: updatedLine.unit, mrp: Number(updatedLine.mrp), discount: Number(updatedLine.discount), sellingPrice: Number(updatedLine.price), purchasePrice: updatedLine.purchasePrice === '' ? '' : Number(updatedLine.purchasePrice) };
      setData(current => ({ ...current, products: current.products.map(item => item.id === product.id ? updatedProduct : item) }));
      // This editor is also available from the Billing view. Save through the
      // API so the edit is not replaced by the old SQLite record on restart.
      fetch(`${apiUrl}/products`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(updatedProduct) }).catch(() => {});
      setLine(current => ({ ...current, ...updatedLine }));
      setSelectedProductId(product.id);
      setNotice({ type: 'success', text: 'Product updated successfully.' });
      close();
    };
    overlay.onclick = event => { if (event.target === overlay) close(); }; document.body.append(overlay);
  }
  function openPrintPreview(bill) {
    const safe = value => String(value ?? '').replace(/[&<>"']/g, character => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[character]);
    const shop = { ...defaultShopProfile, ...(data.shopProfile || {}) };
    const date = new Date(bill.createdAt).toLocaleString('en-IN', { dateStyle: 'short', timeStyle: 'short' });
    const rows = bill.items.map((item, index) => `<tr><td>${index + 1}</td><td>${safe(item.productName)}</td><td>${item.quantity}</td><td>${neatNumber(item.mrp)}</td><td>${neatNumber(item.discount)}</td><td>${neatNumber(item.sellingPrice)}</td><td>${neatNumber(item.total)}</td></tr>`).join('');
    const receiptWindow = window.open('', '_blank', 'width=460,height=760');
    if (!receiptWindow) return setNotice({ type: 'error', text: 'Allow pop-ups to open the print preview.' });
    receiptWindow.document.write(`<!doctype html><html><head><title>${safe(bill.number)} receipt</title><style>@page{size:80mm auto;margin:0}*{box-sizing:border-box}body{margin:0;background:#eef0f3;font-family:Arial,sans-serif;color:#182231;font-size:11px}.toolbar{position:sticky;top:0;display:flex;justify-content:space-between;align-items:center;padding:12px 16px;background:#182231;color:#fff}.toolbar button{border:0;border-radius:5px;padding:8px 16px;background:#d92338;color:#fff;font-weight:bold;cursor:pointer}.receipt{width:80mm;min-height:130mm;margin:14px auto;background:#fff;border:1px solid #8b929b;padding:10px}.brand{text-align:center;border-bottom:3px solid #df2338;padding-bottom:10px}.brand h1{margin:0;color:#df2338;font-size:20px}.brand p{margin:5px 0 0;line-height:1.45}.meta{line-height:1.6;padding:7px 0}.meta b{margin-right:5px}table{width:100%;border-collapse:collapse;font-size:10px}th{text-align:left;border-bottom:1px solid #77808b;padding:5px 2px;font-weight:700}td{padding:5px 2px;border-bottom:1px solid #d5d9de;vertical-align:top;white-space:nowrap;font-weight:600}td:nth-child(2){white-space:normal;font-weight:700}th:nth-child(n+4),td:nth-child(n+4){text-align:right}.summary{margin-top:8px;border-bottom:3px double #25313e;padding-bottom:6px}.summary div{display:flex;justify-content:space-between;padding:2px 0}.summary .total{font-size:15px;font-weight:700;border-top:3px solid #25313e;margin-top:10px;padding-top:8px}.footer{text-align:center;padding-top:8px;font-size:11px}@media print{body{background:#fff}.toolbar{display:none}.receipt{margin:0;border:0;width:80mm;min-height:0}}</style></head><body><div class="toolbar"><span>Bill saved — print preview</span><button onclick="window.print()">Print receipt</button></div><main class="receipt"><header class="brand"><h1>${safe(shop.ownerName)}</h1><p>${safe(shop.address).replace(/\n/g, '<br>')}<br>${safe(shop.phone)}</p></header><section class="meta"><b>Bill No: ${safe(bill.number)}</b> &nbsp; <b>Date: ${safe(date)}</b><br><b>Customer:</b> ${safe(bill.customerName || 'Prabhu Shri Ram')} &nbsp; <b>Phone:</b> ${safe(bill.mobile || '—')}</section><table><thead><tr><th>#</th><th>Product</th><th>Qty</th><th>MRP</th><th>Discount</th><th>Price</th><th>Amt</th></tr></thead><tbody>${rows}</tbody></table><section class="summary"><div><span>Total MRP</span><b>Rs. ${Number(bill.subtotal).toFixed(0)}</b></div><div><span>Total discount</span><b>Rs. ${Number(bill.totalDiscount).toFixed(2)}</b></div><div><span>Item total</span><b>Rs. ${Number(bill.subtotal - bill.totalDiscount).toFixed(2)}</b></div><div class="total"><span>Total bill value</span><b>Rs. ${Number(bill.grandTotal).toFixed(0)}</b></div><div><span>Paid amount</span><b>Rs. ${Number(bill.paidAmount || 0).toFixed(0)}</b></div><div><span>Remaining amount</span><b>Rs. ${Number(bill.remainingAmount || 0).toFixed(0)}</b></div><div><span><b>Payment method:</b> ${safe(bill.payment)}</span></div></section><footer class="footer">Unsealed Item will not be changed or returned <br> No exchange or return after 3 days <br> Thank you for shopping with us.</footer><script>document.addEventListener('keydown',function(event){if(event.key==='Enter'&&!event.repeat&&!event.ctrlKey&&!event.altKey&&!event.metaKey){event.preventDefault();window.print();}});</script></main></body></html>`);
    const receiptDetailStyle = receiptWindow.document.createElement('style');
    receiptDetailStyle.textContent = 'table{font-size:12px}tbody td{font-weight:700}';
    receiptWindow.document.head.append(receiptDetailStyle);
    // Discount is retained in the savings summary, but not as a per-item
    // receipt-table column.
    receiptWindow.document.querySelectorAll('table tr').forEach(row => row.children[4]?.remove());
    receiptWindow.document.querySelector('.summary div:nth-child(2) span')?.replaceChildren('Savings');
    receiptWindow.document.querySelector('.summary div:nth-child(2) b').textContent = `Rs. ${neatNumber(bill.totalDiscount)}`;
    receiptWindow.document.querySelector('.summary div:nth-child(3) b').textContent = `Rs. ${neatNumber(bill.subtotal - bill.totalDiscount)}`;
    const toolbar = receiptWindow.document.querySelector('.toolbar');
    const printButton = toolbar?.querySelector('button');
    const downloadPdfButton = receiptWindow.document.createElement('button');
    const shareWhatsAppButton = receiptWindow.document.createElement('button');
    const receiptStatus = toolbar?.querySelector('span');
    downloadPdfButton.type = 'button';
    downloadPdfButton.textContent = 'Download PDF';
    downloadPdfButton.style.background = '#fff';
    downloadPdfButton.style.color = '#182231';
    shareWhatsAppButton.type = 'button';
    shareWhatsAppButton.textContent = 'Share on WhatsApp';
    shareWhatsAppButton.style.background = '#25d366';
    shareWhatsAppButton.style.color = '#102b19';
    if (toolbar && printButton) {
      const actions = receiptWindow.document.createElement('div');
      actions.style.display = 'flex';
      actions.style.gap = '8px';
      toolbar.replaceChild(actions, printButton);
      actions.append(shareWhatsAppButton, downloadPdfButton, printButton);
    }
    receiptWindow.document.close();
    const showPdfSuccessPopup = openPdf => {
      const successPopup = receiptWindow.document.createElement('button');
      successPopup.type = 'button';
      successPopup.innerHTML = '<b>PDF downloaded successfully</b><span>Click here to open it</span>';
      successPopup.style.cssText = 'position:fixed;bottom:20px;left:50%;transform:translateX(-50%);z-index:20;display:grid;gap:4px;min-width:220px;padding:13px 16px;border:0;border-radius:8px;background:#df2338;color:#fff;box-shadow:0 10px 28px rgba(15,23,42,.28);font:13px Arial;text-align:left;cursor:pointer';
      successPopup.querySelector('b').style.fontSize = '13px';
      successPopup.querySelector('span').style.fontSize = '11px';
      successPopup.addEventListener('click', async () => {
        successPopup.disabled = true;
        successPopup.querySelector('span').textContent = 'Opening PDF…';
        try {
          await openPdf();
          successPopup.remove();
        } catch (error) {
          successPopup.disabled = false;
          successPopup.querySelector('span').textContent = error.message || 'Could not open the PDF.';
        }
      });
      receiptWindow.document.body.append(successPopup);
      receiptWindow.setTimeout(() => successPopup.remove(), 12_000);
    };
    const createBrowserPdf = () => {
      const pdf = new jsPDF({ unit: 'mm', format: [80, 200] });
      const printDate = new Date(bill.createdAt).toLocaleString('en-IN', { dateStyle: 'short', timeStyle: 'short' });
      pdf.setFont('helvetica', 'bold');
      pdf.setFontSize(16);
      pdf.setTextColor(220, 35, 56);
      pdf.text(shop.ownerName, 40, 10, { align: 'center' });
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(8);
      pdf.setTextColor(24, 34, 49);
      pdf.text(pdf.splitTextToSize(shop.address, 70), 40, 15, { align: 'center' });
      pdf.text(shop.phone, 40, 23, { align: 'center' });
      pdf.setDrawColor(220, 35, 56);
      pdf.setLineWidth(.7);
      pdf.line(4, 26, 76, 26);
      pdf.setFont('helvetica', 'bold');
      pdf.setFontSize(7.5);
      pdf.text(`Bill No: ${bill.number}    Date: ${printDate}`, 4, 31);
      pdf.text(`Customer: ${bill.customerName || 'Prabhu Shri Ram'}`, 4, 35);
      pdf.text(`Phone: ${bill.mobile || '—'}`, 4, 39);
      autoTable(pdf, {
        startY: 42,
        head: [['#', 'Product', 'Qty', 'MRP', 'Price', 'Amt']],
        body: bill.items.map((item, index) => [index + 1, item.productName, neatNumber(item.quantity), neatNumber(item.mrp), neatNumber(item.sellingPrice), neatNumber(item.total)]),
        margin: { left: 4, right: 4 },
        styles: { fontSize: 6.5, cellPadding: 1, textColor: [24, 34, 49] },
        headStyles: { fillColor: [255, 255, 255], textColor: [24, 34, 49], lineWidth: .1, lineColor: [100, 110, 120] },
        columnStyles: { 0: { cellWidth: 4 }, 1: { cellWidth: 25 }, 2: { cellWidth: 7 }, 3: { cellWidth: 9 }, 4: { cellWidth: 10 }, 5: { cellWidth: 10 } },
      });
      let y = pdf.lastAutoTable.finalY + 6;
      if (y > 164) { pdf.addPage([80, 200]); y = 10; }
      pdf.setFontSize(8);
      const summary = [
        ['Total MRP', Number(bill.subtotal).toFixed(2)],
        ['Savings', Number(bill.totalDiscount).toFixed(2)],
        ['Item total', Number(bill.subtotal - bill.totalDiscount).toFixed(2)],
        ['Total bill value', Number(bill.grandTotal).toFixed(2)],
        ['Paid amount', Number(bill.paidAmount || 0).toFixed(2)],
        ['Remaining amount', Number(bill.remainingAmount || 0).toFixed(2)],
      ];
      summary.forEach(([label, value], index) => {
        if (index === 3) {
          pdf.setDrawColor(37, 49, 62);
          pdf.setLineWidth(.4);
          pdf.line(4, y, 76, y);
          y += 8;
          pdf.setFont('helvetica', 'bold');
          pdf.setFontSize(10);
        }
        pdf.text(label, 4, y);
        pdf.text(`Rs. ${value}`, 76, y, { align: 'right' });
        if (index === 3) { pdf.setFont('helvetica', 'normal'); pdf.setFontSize(8); }
        y += 5;
      });
      pdf.setFont('helvetica', 'bold');
      pdf.text(`Payment method: ${bill.payment}`, 4, y + 2);
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(7);
      pdf.text(['Unsealed Item will not be changed or returned', 'No exchange or return after 3 days', 'Thank you for shopping with us.'], 40, y + 12, { align: 'center' });
      return pdf.output('blob');
    };
    const downloadBrowserPdf = () => {
      const blobUrl = URL.createObjectURL(createBrowserPdf());
      const filename = `Bill-${bill.number}-${new Date(bill.createdAt).toISOString().replace(/[:.]/g, '-')}.pdf`;
      const link = receiptWindow.document.createElement('a');
      link.href = blobUrl;
      link.download = filename;
      receiptWindow.document.body.append(link);
      link.click();
      link.remove();
      receiptWindow.setTimeout(() => URL.revokeObjectURL(blobUrl), 300_000);
      return blobUrl;
    };
    downloadPdfButton.addEventListener('click', async () => {
      const saveReceiptPdf = window.canteen?.saveReceiptPdf;
      if (!saveReceiptPdf) {
        const blobUrl = downloadBrowserPdf();
        if (receiptStatus) receiptStatus.textContent = 'PDF download started.';
        showPdfSuccessPopup(() => { receiptWindow.open(blobUrl, '_blank'); });
        return;
      }
      downloadPdfButton.disabled = true;
      downloadPdfButton.textContent = 'Saving PDF…';
      try {
        const result = await saveReceiptPdf({
          billNumber: bill.number,
          createdAt: bill.createdAt,
          html: receiptWindow.document.documentElement.outerHTML,
        });
        downloadPdfButton.textContent = 'PDF saved';
        if (receiptStatus) receiptStatus.textContent = `PDF saved: ${result.path}`;
        showPdfSuccessPopup(() => window.canteen?.openReceiptPdf(result.path));
      } catch (error) {
        try {
          const blobUrl = downloadBrowserPdf();
          downloadPdfButton.disabled = false;
          downloadPdfButton.textContent = 'PDF download started';
          if (receiptStatus) receiptStatus.textContent = 'Desktop save failed, so a browser PDF download was started.';
          showPdfSuccessPopup(() => { receiptWindow.open(blobUrl, '_blank'); });
        } catch (fallbackError) {
          downloadPdfButton.disabled = false;
          downloadPdfButton.textContent = 'Download PDF';
          if (receiptStatus) receiptStatus.textContent = fallbackError.message || error.message || 'Could not save the PDF.';
        }
      }
    });
    shareWhatsAppButton.addEventListener('click', async () => {
      const shareReceiptOnWhatsApp = window.canteen?.shareReceiptOnWhatsApp;
      if (!shareReceiptOnWhatsApp) {
        if (receiptStatus) receiptStatus.textContent = 'Open this bill in the desktop app to share it on WhatsApp.';
        return;
      }
      shareWhatsAppButton.disabled = true;
      shareWhatsAppButton.textContent = 'Opening WhatsApp…';
      try {
        const share = await shareReceiptOnWhatsApp({ phone: bill.mobile, text: whatsappBillText(bill, shop) });
        shareWhatsAppButton.textContent = 'WhatsApp opened';
        if (receiptStatus) receiptStatus.textContent = `${share.hasRecipient ? 'Customer chat opened.' : 'WhatsApp opened.'} Bill text ready hai — WhatsApp mein Send karein.`;
      } catch (error) {
        shareWhatsAppButton.disabled = false;
        shareWhatsAppButton.textContent = 'Share on WhatsApp';
        if (receiptStatus) receiptStatus.textContent = error.message || 'Could not open WhatsApp.';
      }
    });
    receiptWindow.addEventListener('keydown', event => { if (event.key === 'Escape') receiptWindow.close(); });
  }
  const handleCustomerIdChange = value => { setCustomerId(value); const enteredId = value.trim().toLowerCase(); if (!enteredId) { setCustomer(''); setMobile(''); setAddress(''); return; } const enteredNumber = enteredId.match(/\d+$/)?.[0]; const matchedCustomer = (data.customers || []).find(item => { const savedId = String(item.customerId || '').trim().toLowerCase(), savedNumber = savedId.match(/\d+$/)?.[0]; return savedId === enteredId || (enteredNumber && savedNumber && Number(savedNumber) === Number(enteredNumber)); }); if (matchedCustomer) { setCustomer(matchedCustomer.name || ''); setMobile(matchedCustomer.phone || ''); setAddress(matchedCustomer.address || ''); } };
  function resetBill() { sessionStorage.removeItem('dc-billing-draft'); setCart([]); setCustomerId(''); setCustomer(''); setMobile(''); setAddress(''); setPayment('Cash'); setPaid(''); setAdditional(''); setLine(blank); setSelectedProductId(null); setEditingId(null); setHideSuggestions(false); }
  const holdBill = () => {
    if (!cart.length) return setNotice({ type: 'error', text: 'Add at least one item before holding this bill.' });
    const heldBill = { id: uid(), heldAt: new Date().toISOString(), cart, customerId, customer, mobile, address, payment, paid, additional, line };
    setData(current => ({ ...current, heldBills: [heldBill, ...(current.heldBills || [])] }));
    resetBill();
    setNotice({ type: 'success', text: 'Bill placed on hold. You can continue it from View held.' });
  };
  const continueHeldBill = heldBill => {
    if (cart.length && !window.confirm('Replace the current bill with this held bill? Your current work remains saved as a browser draft.')) return;
    setCart(heldBill.cart || []); setCustomerId(heldBill.customerId || ''); setCustomer(heldBill.customer || ''); setMobile(heldBill.mobile || ''); setAddress(heldBill.address || ''); setPayment(heldBill.payment || 'Cash'); setPaid(heldBill.paid || ''); setAdditional(heldBill.additional || ''); setLine(heldBill.line || blank); setSelectedProductId(null); setEditingId(null); setHideSuggestions(false);
    setData(current => ({ ...current, heldBills: (current.heldBills || []).filter(item => item.id !== heldBill.id) }));
    setShowHeldBills(false);
    setNotice({ type: 'success', text: 'Held bill restored. Continue billing when ready.' });
  };
  const removeHeldBill = heldBill => {
    if (!window.confirm('Remove this held bill? This cannot be undone.')) return;
    setData(current => ({ ...current, heldBills: (current.heldBills || []).filter(item => item.id !== heldBill.id) }));
    setNotice({ type: 'success', text: 'Held bill removed.' });
  };
  async function saveBill() {
    if (!cart.length) return setNotice({ type: 'error', text: 'Add at least one item to the bill.' });
    const paidAmount = Math.min(roundedTotal, Math.max(0, Number(paid) || 0)), remainingAmount = Math.max(0, roundedTotal - paidAmount);
    const bill = { id: uid(), number: billNumber(data), createdAt: new Date().toISOString(), customerId, customerName: customer, mobile, address, payment, paidAmount, remainingAmount, items: cart.map(item => ({ productId: item.id, productName: item.name, barcode: item.barcode, quantity: item.qty, unit: item.unit, mrp: item.mrp, discount: item.discount, sellingPrice: item.sellingPrice, total: item.sellingPrice * item.qty })), subtotal, totalDiscount: discount, additionalAmount, grandTotal: roundedTotal };
    setData(current => ({ ...current, bills: [bill, ...current.bills] }));
    openPrintPreview(bill);
    resetBill();
    setNotice({ type: 'success', text: `Bill ${bill.number} saved. You can create the next bill now.` });
  }
  useEffect(() => {
    const handleBillingShortcut = event => {
      // These shortcuts belong only to the main billing workspace. Let an
      // open product/held-bill dialog keep its normal keyboard behaviour.
      if (!event.ctrlKey || event.altKey || event.metaKey || document.querySelector('[role="dialog"]')) return;

      const action = event.key.toLowerCase();
      const isShortcutHelp = event.key === '/';
      const shortcuts = {
        u: openNewProductRecord,
        e: resetBill,
        w: updateProductRecord,
        s: saveBill,
        h: holdBill,
        j: () => setShowHeldBills(true),
        l: () => setPaid(roundedTotal.toFixed(2)),
      };
      const run = isShortcutHelp ? () => setShowShortcutHelp(true) : shortcuts[action];
      if (!run) return;

      // Prevent browser actions such as Save page and Print, so the billing
      // command is consistently executed instead.
      event.preventDefault();
      run();
    };
    document.addEventListener('keydown', handleBillingShortcut);
    return () => document.removeEventListener('keydown', handleBillingShortcut);
  }, [roundedTotal, cart, line, selectedProductId, data, customerId, customer, mobile, address, payment, paid, additional, subtotal, discount, additionalAmount]);
  useEffect(() => {
    // In the packaged app, Electron intercepts Ctrl+W before it can close the
    // window and forwards this action through the preload bridge.
    if (!window.canteen?.onBillingShortcut) return;
    return window.canteen.onBillingShortcut(shortcut => {
      if (shortcut === 'edit-product') updateProductRecord();
    });
  }, [line, selectedProductId, data.products, data.categories]);
  useEffect(() => {
    if (!showShortcutHelp) return;
    const overlay = document.createElement('div');
    overlay.className = 'shortcut-help-overlay';
    overlay.setAttribute('role', 'dialog');
    overlay.setAttribute('aria-modal', 'true');
    overlay.setAttribute('aria-label', 'Billing keyboard shortcuts');
    overlay.innerHTML = `<section class="shortcut-help-modal"><header><div><span>KEYBOARD SHORTCUTS</span><h2>Work faster in billing</h2><p>All shortcuts available in the app, collected in one place.</p></div><button type="button" class="bill-modal-close" aria-label="Close shortcut help">×</button></header><div class="shortcut-help-intro"><span class="shortcut-help-dot"></span><p><b>Tip:</b> Billing actions work on the <b>New bill</b> screen. Field focus shortcuts work when you are not already typing in a field.</p></div><div class="shortcut-help-list"><section><h3>Navigate</h3><div class="shortcut-help-grid"><div><kbd>Ctrl <i>+</i> N</kbd><span>Open a new bill</span></div><div><kbd>Ctrl <i>+</i> O</kbd><span>Open saved bills</span></div><div><kbd>Ctrl <i>+</i> D</kbd><span>Open dashboard</span></div><div><kbd>Ctrl <i>+</i> L</kbd><span>Go to New bill / set full paid amount</span></div></div></section><section><h3>Bill actions</h3><div class="shortcut-help-grid"><div><kbd>Ctrl <i>+</i> U</kbd><span>Add a product record</span></div><div><kbd>Ctrl <i>+</i> W</kbd><span>Edit the selected product</span></div><div><kbd>Ctrl <i>+</i> E</kbd><span>Clear the current bill</span></div><div><kbd>Ctrl <i>+</i> S</kbd><span>Save and print the bill</span></div><div><kbd>Ctrl <i>+</i> H</kbd><span>Hold the current bill</span></div><div><kbd>Ctrl <i>+</i> J</kbd><span>View held bills</span></div></div></section><section><h3>Jump to a field</h3><div class="shortcut-help-grid"><div><kbd>Ctrl <i>+</i> B</kbd><span>Focus barcode</span></div><div><kbd>Ctrl <i>+</i> P</kbd><span>Focus product search</span></div><div><kbd>Ctrl <i>+</i> M</kbd><span>Focus MRP</span></div><div><kbd>Ctrl <i>+</i> C</kbd><span>Focus customer name</span></div></div></section><section><h3>Product search &amp; app controls</h3><div class="shortcut-help-grid"><div><kbd>↑ <i>/</i> ↓</kbd><span>Move through product suggestions</span></div><div><kbd>Enter</kbd><span>Choose a suggestion or add the product</span></div><div><kbd>Esc</kbd><span>Close search suggestions or this help</span></div><div><kbd>Ctrl <i>+</i> + <i>/</i> −</kbd><span>Zoom in or out <small>Desktop app</small></span></div><div><kbd>Ctrl <i>+</i> 0</kbd><span>Reset zoom <small>Desktop app</small></span></div><div><kbd>Ctrl <i>+</i> /</kbd><span>Open this shortcut help</span></div></div></section></div><footer><span>Press <kbd>Esc</kbd> anytime to close</span><button type="button" class="btn">Done</button></footer></section>`;
    overlay.querySelectorAll('.shortcut-help-grid')[2]?.insertAdjacentHTML('beforeend', '<div><kbd>Ctrl <i>+</i> Q</kbd><span>Focus quantity</span></div>');
    const close = () => setShowShortcutHelp(false);
    overlay.querySelector('.bill-modal-close')?.addEventListener('click', close);
    overlay.querySelector('footer button')?.addEventListener('click', close);
    overlay.addEventListener('mousedown', event => { if (event.target === overlay) close(); });
    const onKeyDown = event => { if (event.key === 'Escape') close(); };
    document.addEventListener('keydown', onKeyDown);
    document.body.append(overlay);
    return () => { document.removeEventListener('keydown', onKeyDown); overlay.remove(); };
  }, [showShortcutHelp]);
  const updateOnEnter = event => {
    if (event.key !== 'Enter' || !event.target.matches('input, select')) return;
    event.preventDefault();
    const isProductLookup = event.target.matches('input[placeholder="Scan or enter barcode"], input[placeholder="Search product"]');
    if (isProductLookup) {
      const value = event.target.value.trim().toLowerCase();
      const product = data.products.find(item => String(item.barcode || '').trim().toLowerCase() === value || String(item.name || '').trim().toLowerCase() === value);
      if (product && product.id !== selectedProductId) selectProduct(product);
      else focusLineQuantity();
      return;
    }
    if (editingId) updateEditedProduct();
    else addProduct();
  };
  // Whole amounts stay clean (e.g. 20), while decimal discounts keep their
  // exact value in the calculated price and amount columns (e.g. 19.5).
  useEffect(() => {
    const amount = (value) => neatNumber(value);
    const currentAmount = (Number(line.price) || 0) * (Number(line.quantity) || 1);
    const addAmount = document.querySelector('.billing-add-product .billing-amount b');
    if (addAmount) addAmount.textContent = `Rs. ${amount(currentAmount)}`;
    const activeInputLabel = document.activeElement?.getAttribute('aria-label');
    document.querySelectorAll('.billing-items tbody tr').forEach((row, index) => {
      const item = cart[index], cells = row.querySelectorAll('td');
      if (!item || cells.length < 8) return;
      const makeRowInput = (cell, value, label, update, minimum = 0) => {
        const input = document.createElement('input');
        // Text + inputMode avoids browser number inputs discarding a trailing
        // decimal point before the next digit is entered.
        input.type = 'text'; input.inputMode = 'decimal'; input.value = value;
        input.className = 'bill-row-input'; input.setAttribute('aria-label', `${label} for ${item.name}`);
        // Free typing is committed on blur/change; arrow adjustments below
        // commit immediately. This avoids rebuilding the row mid-typing.
        let lastValidValue = String(value);
        input.addEventListener('input', () => {
          if (/^\d*(?:\.\d{0,2})?$/.test(input.value)) lastValidValue = input.value;
          else input.value = lastValidValue;
        });
        const commit = () => {
          if (!/^\d*\.?\d{1,2}$/.test(input.value)) { input.value = neatNumber(value); return; }
          update(input.value);
        };
        input.addEventListener('change', commit);
        input.addEventListener('keydown', event => {
          if (event.key === 'Enter') { event.preventDefault(); input.blur(); return; }
          if (event.key !== 'ArrowUp' && event.key !== 'ArrowDown') return;
          // Commit keyboard increments immediately, so all linked fields and
          // the bill total refresh just like they do after a mouse-spinner tap.
          event.preventDefault();
          const current = Number(input.value) || 0;
          const next = label === 'Quantity'
            ? (event.key === 'ArrowUp' ? Math.floor(current) + 1 : Math.max(minimum, Math.ceil(current) - 1))
            : Math.max(minimum, current + (event.key === 'ArrowUp' ? 0.01 : -0.01));
          input.value = neatNumber(next);
          input.dispatchEvent(new Event('input', { bubbles: true }));
          input.dispatchEvent(new Event('change', { bubbles: true }));
        });
        input.addEventListener('focus', () => input.select());
        cell.replaceChildren(input);
        if (input.getAttribute('aria-label') === activeInputLabel) requestAnimationFrame(() => input.focus());
        return input;
      };
      const quantityInput = makeRowInput(cells[2], item.qty, 'Quantity', value => setCart(items => items.map(entry => entry.id === item.id ? { ...entry, qty: Math.max(0.01, Number(value) || 1) } : entry)), 0.01);
      const mrpInput = makeRowInput(cells[4], item.mrp, 'MRP', value => setCart(items => items.map(entry => {
        if (entry.id !== item.id) return entry;
        const mrp = Math.max(0, Number(value) || 0), discount = Math.min(Number(entry.discount) || 0, mrp);
        return { ...entry, mrp, discount, sellingPrice: mrp - discount };
      })));
      const discountInput = makeRowInput(cells[5], item.discount, 'Discount', value => setCart(items => items.map(entry => {
        if (entry.id !== item.id) return entry;
        const discount = Math.min(Math.max(0, Number(value) || 0), Number(entry.mrp) || 0);
        return { ...entry, discount, sellingPrice: (Number(entry.mrp) || 0) - discount };
      })));
      const priceInput = makeRowInput(cells[6], item.sellingPrice, 'Price', value => setCart(items => items.map(entry => {
        if (entry.id !== item.id) return entry;
        const sellingPrice = Math.min(Math.max(0, Number(value) || 0), Number(entry.mrp) || 0);
        return { ...entry, sellingPrice, discount: (Number(entry.mrp) || 0) - sellingPrice };
      })));
      const updateLiveAmount = () => { cells[7].textContent = `Rs. ${amount((Number(priceInput.value) || 0) * (Number(quantityInput.value) || 0))}`; };
      quantityInput.addEventListener('input', updateLiveAmount);
      mrpInput.addEventListener('input', () => { const mrp = Math.max(0, Number(mrpInput.value) || 0), discount = Math.min(Math.max(0, Number(discountInput.value) || 0), mrp); discountInput.value = discount; priceInput.value = mrp - discount; updateLiveAmount(); });
       discountInput.addEventListener('input', () => { if (!/^\d*\.?\d+$/.test(discountInput.value)) return updateLiveAmount(); const mrp = Math.max(0, Number(mrpInput.value) || 0), discount = Math.min(Math.max(0, Number(discountInput.value)), mrp); if (discount !== Number(discountInput.value)) discountInput.value = neatNumber(discount); priceInput.value = neatNumber(mrp - discount); updateLiveAmount(); });
       priceInput.addEventListener('input', () => { if (!/^\d*\.?\d+$/.test(priceInput.value)) return updateLiveAmount(); const mrp = Math.max(0, Number(mrpInput.value) || 0), price = Math.min(Math.max(0, Number(priceInput.value)), mrp); if (price !== Number(priceInput.value)) priceInput.value = neatNumber(price); discountInput.value = neatNumber(mrp - price); updateLiveAmount(); });
      updateLiveAmount();
    });
  }, [cart, line.price, line.quantity]);
   return <div className="billing-page"><section className="panel billing-customer"><div className="billing-section-head"><h2>Customer details</h2></div><div className="billing-customer-grid"><Field label="Customer ID"><input placeholder="Enter customer ID" value={customerId} onChange={e => handleCustomerIdChange(e.target.value)} /></Field><Field label="Customer name"><input placeholder="Enter Customer name" value={customer} onChange={e => setCustomer(e.target.value)} /></Field><Field label="Phone number"><input placeholder="Phone number" value={mobile} onChange={e => setMobile(e.target.value)} /></Field><Field label="Address"><input placeholder="Enter address" value={address} onChange={e => setAddress(e.target.value)} /></Field></div></section><section className="panel billing-add-product"><div className="billing-section-head"><h2>Add product</h2><span>Scan a barcode or search by name</span></div><div className="billing-product-grid" onKeyDown={updateOnEnter}><Field label="Barcode"><input placeholder="Scan or enter barcode" value={line.barcode} onChange={e => handleBarcodeChange(e.target.value)} /></Field><Field label="Product name"><input placeholder="Search product" value={line.name} onChange={e => change('name', e.target.value)} aria-autocomplete="list" aria-controls="billing-product-suggestions" aria-activedescendant={activeSuggestionIndex >= 0 ? `billing-product-suggestion-${matches[activeSuggestionIndex]?.id}` : undefined} />{matches.length > 0 && <div id="billing-product-suggestions" className="billing-product-suggestions" role="listbox">{matches.map((product, index) => <button type="button" id={`billing-product-suggestion-${product.id}`} role="option" aria-selected={index === activeSuggestionIndex} className={index === activeSuggestionIndex ? 'active' : ''} key={product.id} onMouseEnter={() => setActiveSuggestionIndex(index)} onClick={() => selectProduct(product)}><span className="billing-suggestion-row"><b>{product.name}</b><small>MRP: Rs. {Number(product.mrp || 0).toFixed(0)}</small></span></button>)}</div>}</Field><Field label="Category"><input placeholder="Enter category" value={line.category} onChange={e => change('category', e.target.value)} /></Field><Field label="Quantity"><input placeholder="Quantity" type="number" min="1" value={line.quantity} onChange={e => change('quantity', e.target.value)} /></Field><Field label="Unit"><input placeholder="Enter unit" value={line.unit} onChange={e => change('unit', e.target.value)} /></Field><Field label="MRP"><DecimalAmountInput label="MRP" value={line.mrp} onChange={value => change('mrp', value)} /></Field><Field label="Discount (₹)"><DecimalAmountInput label="Discount" value={line.discount} onChange={value => change('discount', value)} /></Field><Field label="Price"><DecimalAmountInput label="Price" value={line.price} onChange={value => change('price', value)} /></Field><div className="billing-amount"><span>Amount</span><b>Rs. {((Number(line.price) || 0) * (Number(line.quantity) || 1)).toFixed(0)}</b></div><Button type="button" onClick={addProduct}>Add product</Button></div></section><section className="panel billing-items"><div className="billing-section-head"><h2>Bill items</h2><span>{cart.length} {cart.length === 1 ? 'item' : 'items'}</span></div><div className="billing-table-wrap"><table><thead><tr><th>S.NO.</th><th>PRODUCT NAME</th><th>QTY</th><th>UNIT</th><th>MRP</th><th>DISCOUNT</th><th>PRICE</th><th>AMOUNT</th><th>ACTION</th></tr></thead><tbody>{cart.map((item, index) => <tr key={item.id}><td>{index + 1}</td><td><b>{item.name}</b><small>{item.barcode}</small></td><td><input type="number" min="1" max={item.stock} value={item.qty} onChange={e => setCart(items => items.map(entry => entry.id === item.id ? { ...entry, qty: Math.max(1, Math.min(entry.stock, Number(e.target.value) || 1)) } : entry))} /></td><td>{item.unit}</td><td>Rs. {item.mrp.toFixed(0)}</td><td>Rs. {item.discount.toFixed(0)}</td><td>Rs. {item.sellingPrice.toFixed(0)}</td><td><b>Rs. {(item.sellingPrice * item.qty).toFixed(0)}</b></td><td><button className="bill-remove" onClick={() => setCart(items => items.filter(entry => entry.id !== item.id))}>Remove</button></td></tr>)}</tbody></table>{!cart.length && <Empty>Add products to begin this bill.</Empty>}</div></section><section className="panel billing-summary"><div className="billing-summary-head"><h2>Bill summary</h2><Field label="Payment method"><select value={payment} onChange={e => setPayment(e.target.value)}><option>Cash</option><option>UPI</option></select></Field></div><div className="summary-rows"><p><span>Total MRP</span><b>Rs. {subtotal.toFixed(2)}</b></p><p><span>Total discount</span><b>Rs. {discount.toFixed(2)}</b></p><p><span>Item total</span><b>Rs. {total.toFixed(2)}</b></p><hr /><p className="bill-total"><span>Total bill value</span><strong>Rs. {total.toFixed(2)}</strong></p><p><span>Remaining amount</span><b>Rs. {Math.max(0, total - (Number(paid) || 0)).toFixed(2)}</b></p></div><div className="paid-row"><Field label="Paid amount"><input type="number" min="0" value={paid} onChange={e => setPaid(e.target.value)} /></Field><Button type="button" variant="secondary" onClick={() => setPaid(total.toFixed(2))}>Paid</Button></div><div className="billing-actions"><Button type="button" variant="ghost" onClick={resetBill}>Clear bill</Button><Button type="button" variant="secondary" onClick={holdBill}>Hold bill</Button><Button type="button" variant="secondary" onClick={() => setShowHeldBills(true)}>View held{(data.heldBills || []).length ? ` (${data.heldBills.length})` : ''}</Button><Button type="button" onClick={saveBill}>Save &amp; print</Button></div></section>{showHeldBills && <HeldBillsModal bills={data.heldBills || []} onClose={() => setShowHeldBills(false)} onContinue={continueHeldBill} onRemove={removeHeldBill} />}</div>;
}

function HeldBillsModal({ bills, onClose, onContinue, onRemove }) {
  return <div className="held-bills-overlay" role="dialog" aria-modal="true" aria-label="Held bills"><section className="held-bills-modal"><header><div><span>HELD BILLS</span><h2>Continue a saved bill</h2><p>Held bills are not finalised and do not affect stock.</p></div><button className="bill-modal-close" onClick={onClose} aria-label="Close held bills">×</button></header>{bills.length ? <div className="held-bills-list">{bills.map(bill => { const total = roundBillTotal((bill.cart || []).reduce((sum, item) => sum + Number(item.sellingPrice || 0) * Number(item.qty || 0), 0)); return <article className="held-bill-card" key={bill.id}><div className="held-bill-main"><b>{bill.customer || 'Walk-in customer'}</b><span>{(bill.cart || []).length} item{(bill.cart || []).length === 1 ? '' : 's'} · Held {new Date(bill.heldAt).toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' })}</span></div><strong>Rs. {total.toFixed(0)}</strong><div className="held-bill-actions"><Button type="button" variant="secondary" onClick={() => onContinue(bill)}>Continue</Button><button type="button" className="held-remove" onClick={() => onRemove(bill)}>Remove</button></div></article>; })}</div> : <div className="held-bills-empty"><FileText size={26} /><b>No held bills</b><span>Use Hold bill while creating a bill to save it here.</span></div>}<footer><Button type="button" variant="secondary" onClick={onClose}>Done</Button></footer></section></div>;
}

function BillsView() {
  const { data, setData, setNotice } = useStore();
  const [selected, setSelected] = useState(null), [editingBillId, setEditingBillId] = useState(null), [editingProductId, setEditingProductId] = useState(null), [updatingBillId, setUpdatingBillId] = useState(null), [query, setQuery] = useState(''), [dateFilter, setDateFilter] = useState('');
  const bills = data.bills || [];
  const normalizedQuery = query.trim().toLowerCase();
  const filteredBills = bills.filter(item => {
    const billDetails = `${item.number} ${item.customerName || ''} ${item.mobile || ''}`.toLowerCase();
    const containsProduct = (item.items || []).some(product => String(product.productName || product.name || '').toLowerCase().includes(normalizedQuery));
    const search = !normalizedQuery || billDetails.includes(normalizedQuery) || containsProduct;
    const date = !dateFilter || new Date(item.createdAt).toISOString().slice(0, 10) === dateFilter;
    return search && date;
  });
  const bill = bills.find(item => item.id === selected);
  const editingBill = bills.find(item => item.id === editingBillId);
  const editingProduct = data.products.find(item => item.id === editingProductId);
  const updatingBill = bills.find(item => item.id === updatingBillId);
  const printBill = item => {
    const shop = { ...defaultShopProfile, ...(data.shopProfile || {}) };
    const popup = window.open('', '_blank', 'width=420,height=700');
    if (!popup) return setNotice({ type: 'error', text: 'Allow pop-ups to print this bill.' });
    const rows = item.items.map((product, index) => `<tr><td>${index + 1}</td><td>${product.productName}</td><td>${product.quantity}</td><td>${Number(product.mrp).toFixed(0)}</td><td>${Number(product.sellingPrice).toFixed(0)}</td><td>${Number(product.total).toFixed(0)}</td></tr>`).join('');
    popup.document.write(`<!doctype html><title>${item.number} receipt</title><style>@page{size:80mm auto;margin:0}body{font:11px Arial;margin:0;background:#eef0f3;color:#182231}.toolbar{padding:12px;background:#182231;color:#fff;display:flex;justify-content:space-between}.toolbar button{border:0;border-radius:4px;background:#df2338;color:#fff;font-weight:bold;padding:7px 13px}.toolbar-actions{display:flex;gap:8px}.toolbar #share-whatsapp{background:#25d366;color:#102b19}.toolbar #download-pdf{background:#fff;color:#182231}.receipt{box-sizing:border-box;width:80mm;min-height:130mm;margin:14px auto;background:#fff;border:1px solid #89919a;padding:10px}.brand{text-align:center;border-bottom:3px solid #df2338;padding-bottom:8px}.brand h1{color:#df2338;margin:0;font-size:19px}.brand p{line-height:1.45;margin:4px 0}.meta{line-height:1.6;padding:7px 0}table{width:100%;border-collapse:collapse;font-size:9px}th,td{text-align:left;padding:4px 2px;border-bottom:1px solid #d5d9de}th{border-bottom-color:#7c8791}th:nth-child(n+4),td:nth-child(n+4){text-align:right}.totals{border-top:3px solid #25313e;margin-top:8px;padding-top:5px}.totals div{display:flex;justify-content:space-between;padding:2px 0}.totals .grand{font-size:15px;font-weight:bold;border-top:1px solid #25313e;margin-top:5px;padding-top:6px}.footer{text-align:center;border-top:1px solid #77808b;margin-top:7px;padding-top:7px}@media print{body{background:#fff}.toolbar{display:none}.receipt{margin:0;border:0}}</style><body><div class="toolbar"><span>Print preview</span><div class="toolbar-actions"><button id="share-whatsapp" type="button">Share on WhatsApp</button><button id="download-pdf" type="button">Download PDF</button><button onclick="print()">Print receipt</button></div></div><main class="receipt"><div class="brand"><h1>${shop.ownerName}</h1><p>${shop.address.replace(/\n/g, '<br>')}<br>${shop.phone}</p></div><div class="meta"><b>Bill No: ${item.number}</b> &nbsp; <b>Date: ${new Date(item.createdAt).toLocaleString('en-IN')}</b><br><b>Customer:</b> ${item.customerName || 'Prabhu Shri Ram'} &nbsp; <b>Phone:</b> ${item.mobile || '—'}</div><table><thead><tr><th>#</th><th>Product</th><th>Qty</th><th>MRP</th><th>Price</th><th>Amt</th></tr></thead><tbody>${rows}</tbody></table><div class="totals"><div><span>Total MRP</span><b>Rs. ${Number(item.subtotal).toFixed(0)}</b></div><div><span>Total discount</span><b>Rs. ${Number(item.totalDiscount).toFixed(0)}</b></div>${Number(item.additionalAmount || 0) ? `<div><span>Additional amount</span><b>Rs. ${Number(item.additionalAmount).toFixed(0)}</b></div>` : ''}<div class="grand"><span>Total bill value</span><b>Rs. ${Number(item.grandTotal).toFixed(0)}</b></div><div><span>Remaining amount</span><b>Rs. ${Number(item.remainingAmount || 0).toFixed(0)}</b></div><div><span>Payment method</span><b>${item.payment}</b></div></div><div class="footer">Unsealed Item will not be changed or returned <br> No exchange or return after 3 days <br> Thank you for shopping with us.</div></main></body>`);
    const billDetailStyle = popup.document.createElement('style');
    billDetailStyle.textContent = 'table{font-size:11px}tbody td{font-weight:700}';
    popup.document.head.append(billDetailStyle);
    popup.document.close();
    const downloadPdfButton = popup.document.querySelector('#download-pdf');
    downloadPdfButton?.addEventListener('click', () => {
      // This is deliberately the same narrow receipt composition used by the
      // print preview—not a generic invoice layout—so PDF and print match.
      const pdf = new jsPDF({ unit: 'mm', format: [80, 200] });
      pdf.setFont('helvetica', 'bold');
      pdf.setFontSize(16);
      pdf.setTextColor(220, 35, 56);
      pdf.text(shop.ownerName, 40, 10, { align: 'center' });
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(7);
      pdf.setTextColor(24, 34, 49);
      const addressLines = pdf.splitTextToSize(shop.address, 70);
      pdf.text(addressLines, 40, 16, { align: 'center' });
      const phoneY = 16 + addressLines.length * 3.4 + 1;
      pdf.text(shop.phone, 40, phoneY, { align: 'center' });
      pdf.setDrawColor(220, 35, 56);
      const dividerY = phoneY + 6;
      pdf.setLineWidth(.7);
      pdf.line(4, dividerY, 76, dividerY);
      pdf.setFont('helvetica', 'bold');
      pdf.setFontSize(7.5);
      const detailsY = dividerY + 5;
      pdf.text(`Bill No: ${item.number}`, 4, detailsY);
      pdf.text(`Date: ${new Date(item.createdAt).toLocaleString('en-IN', { dateStyle: 'short', timeStyle: 'short' })}`, 29, detailsY);
      pdf.text(`Customer: ${item.customerName || 'Prabhu Shri Ram'}`, 4, detailsY + 4);
      pdf.text(`Phone: ${item.mobile || '—'}`, 51, detailsY + 4);
      autoTable(pdf, {
        startY: detailsY + 8,
        head: [['#', 'Product', 'Qty', 'MRP', 'Price', 'Amt']],
        body: item.items.map((product, index) => [index + 1, product.productName, neatNumber(product.quantity), neatNumber(product.mrp), neatNumber(product.sellingPrice), neatNumber(product.total)]),
        margin: { left: 4, right: 4 },
        theme: 'plain',
        styles: { fontSize: 7, cellPadding: { top: 1.2, right: .6, bottom: 1.2, left: .6 }, textColor: [24, 34, 49], fontStyle: 'bold' },
        headStyles: { fillColor: [255, 255, 255], textColor: [24, 34, 49], lineWidth: .2, lineColor: [120, 130, 140] },
        columnStyles: { 0: { cellWidth: 4 }, 1: { cellWidth: 34 }, 2: { cellWidth: 8 }, 3: { cellWidth: 8, halign: 'right' }, 4: { cellWidth: 8, halign: 'right' }, 5: { cellWidth: 8, halign: 'right' } },
      });
      let y = pdf.lastAutoTable.finalY + 5;
      if (y > 164) { pdf.addPage([80, 200]); y = 10; }
      const writeTotalRow = (label, value, bold = false) => {
        pdf.setFont('helvetica', bold ? 'bold' : 'normal');
        pdf.setFontSize(bold ? 10 : 7.5);
        pdf.text(label, 4, y);
        pdf.text(`Rs. ${Number(value || 0).toFixed(0)}`, 76, y, { align: 'right' });
        y += bold ? 6 : 4;
      };
      pdf.setFontSize(8);
      writeTotalRow('Total MRP', item.subtotal);
      writeTotalRow('Savings', item.totalDiscount);
      writeTotalRow('Item total', Number(item.subtotal || 0) - Number(item.totalDiscount || 0));
      pdf.setDrawColor(37, 49, 62);
      pdf.setLineWidth(.45);
      pdf.line(4, y, 76, y);
      y += 7;
      writeTotalRow('Total bill value', item.grandTotal, true);
      writeTotalRow('Paid amount', item.paidAmount || 0);
      writeTotalRow('Remaining amount', item.remainingAmount || 0);
      pdf.setFont('helvetica', 'normal');
      pdf.text(`Payment method: ${item.payment || 'Cash'}`, 4, y);
      y += 5;
      pdf.setLineWidth(.2);
      pdf.line(4, y, 76, y);
      pdf.setFontSize(6.5);
      pdf.text(['Unsealed Item will not be changed or returned', 'No exchange or return after 3 days', 'Thank you for shopping with us.'], 40, y + 5, { align: 'center' });
      pdf.save(`Bill-${item.number}.pdf`);
    });
    popup.addEventListener('keydown', event => {
      if (event.key === 'Escape') {
        event.preventDefault();
        popup.close();
      }
    });
    const shareButton = popup.document.querySelector('#share-whatsapp');
    const status = popup.document.querySelector('.toolbar span');
    const shareOnWhatsApp = async () => {
      const shareReceiptOnWhatsApp = window.canteen?.shareReceiptOnWhatsApp;
      if (!shareReceiptOnWhatsApp) {
        if (status) status.textContent = 'Open this bill in the desktop app to share it on WhatsApp.';
        return;
      }
      shareButton.disabled = true;
      shareButton.textContent = 'Opening WhatsApp…';
      try {
        const share = await shareReceiptOnWhatsApp({ phone: item.mobile, text: whatsappBillText(item, shop) });
        shareButton.textContent = 'WhatsApp opened';
        if (status) status.textContent = `${share.hasRecipient ? 'Customer chat opened.' : 'WhatsApp opened.'} Bill text ready hai — WhatsApp mein Send karein.`;
      } catch (error) {
        shareButton.disabled = false;
        shareButton.textContent = 'Share on WhatsApp';
        if (status) status.textContent = error.message || 'Could not open WhatsApp.';
      }
    };
    shareButton?.addEventListener('click', shareOnWhatsApp);
  };
  const deleteBill = item => { setData(current => ({ ...current, bills: current.bills.filter(entry => entry.id !== item.id) })); setSelected(current => current === item.id ? null : current); setNotice({ type: 'success', text: `Bill ${item.number} deleted.` }); };
  const updatePayment = (item, paidAmount) => { const paid = Math.min(Number(item.grandTotal || 0), paidAmount); setData(current => ({ ...current, bills: current.bills.map(entry => entry.id === item.id ? { ...entry, paidAmount: paid, remainingAmount: Math.max(0, Number(entry.grandTotal || 0) - paid) } : entry) })); setNotice({ type: 'success', text: `Payment updated for ${item.number}.` }); };
  const saveEditedProduct = updatedProduct => {
    setData(current => ({ ...current, products: current.products.map(product => product.id === updatedProduct.id ? updatedProduct : product) }));
    setNotice({ type: 'success', text: 'Product updated successfully.' });
    setEditingProductId(null);
    // Products are reloaded from SQLite when the app starts. Persist this
    // quick edit there as well, otherwise a restart restores the old record.
    fetch(`${apiUrl}/products`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(updatedProduct) }).catch(() => {});
  };
  return <><section className="panel saved-bills"><div className="saved-bills-toolbar"><div><h2>Saved offline bills</h2><span>{bills.length} saved bill{bills.length === 1 ? '' : 's'}</span></div><div className="saved-bills-filters"><div className="saved-bill-search"><input placeholder="Search bill, customer or product" value={query} onChange={event => setQuery(event.target.value)} /><button aria-label="Search bills"><Search size={17} /></button></div><input aria-label="Filter bills by date" type="date" value={dateFilter} onChange={event => setDateFilter(event.target.value)} /></div></div>{bills.length ? <div className="billing-table-wrap"><table><thead><tr><th>BILL NO.</th><th>DATE &amp; TIME</th><th>CUSTOMER</th><th>PHONE</th><th>REMAINING AMOUNT</th><th>AMOUNT</th><th>ACTION</th></tr></thead><tbody>{filteredBills.map(item => { const amount = Number(item.grandTotal || 0), remaining = Number(item.remainingAmount ?? 0); return <tr key={item.id}><td><b>{item.number}</b></td><td>{new Date(item.createdAt).toLocaleString('en-IN', { dateStyle: 'short', timeStyle: 'short' })}</td><td>{item.customerName || 'Prabhu Shri Ram'}</td><td>{item.mobile || '—'}</td><td><b>Rs. {roundBillTotal(remaining)}</b></td><td><b>Rs. {roundBillTotal(amount)}</b></td><td className="bill-actions"><button className="bill-remove view-bill-button" onClick={() => setSelected(item.id)}>View</button><button className="bill-edit" onClick={() => setEditingBillId(item.id)}>Edit</button><button className="bill-remove view-bill-button" onClick={() => printBill(item)}>Print</button><button className="bill-remove bill-delete" onClick={() => deleteBill(item)}>Delete</button></td></tr> })}</tbody></table>{!filteredBills.length && <Empty>No bills match the selected filters.</Empty>}</div> : <Empty>No bills saved yet. Create a new bill to see it here.</Empty>}</section>{bill && <BillDetailsModal bill={bill} onClose={() => setSelected(null)} onUpdatePayment={() => setUpdatingBillId(bill.id)} onMarkPaid={() => updatePayment(bill, Number(bill.grandTotal || 0))} />}{editingBill && <BillEditModal bill={editingBill} products={data.products} onPrint={printBill} onEditProduct={product => product ? setEditingProductId(product.id) : setNotice({ type: 'error', text: 'Select a product before editing it.' })} onClose={() => setEditingBillId(null)} onSave={updatedBill => { setData(current => { const previous = current.bills.find(entry => entry.id === updatedBill.id); const quantityByProduct = items => items.reduce((totals, item) => ({ ...totals, [item.productId]: (totals[item.productId] || 0) + Number(item.quantity || 0) }), {}); const oldQuantities = quantityByProduct(previous?.items || []), newQuantities = quantityByProduct(updatedBill.items); return { ...current, bills: current.bills.map(entry => entry.id === updatedBill.id ? updatedBill : entry), products: current.products.map(product => product.id in oldQuantities || product.id in newQuantities ? { ...product, stock: Number(product.stock || 0) + (oldQuantities[product.id] || 0) - (newQuantities[product.id] || 0) } : product) }; }); setEditingBillId(null); setNotice({ type: 'success', text: `Bill ${updatedBill.number} updated.` }); }} />}{editingProduct && <ProductQuickEditModal product={editingProduct} categories={data.categories} onClose={() => setEditingProductId(null)} onSave={saveEditedProduct} />}{updatingBill && <PaymentUpdateModal bill={updatingBill} onClose={() => setUpdatingBillId(null)} onSave={paidAmount => { updatePayment(updatingBill, paidAmount); setUpdatingBillId(null); }} />}</>;
}

function BillDetailsModal({ bill, onClose, onUpdatePayment, onMarkPaid }) {
  const amount = Number(bill.grandTotal || 0);
  const paid = Number(bill.paidAmount ?? (bill.remainingAmount == null ? amount : 0));
  const remaining = Math.max(0, Number(bill.remainingAmount ?? amount - paid));
  return <div className="bill-modal-overlay" role="dialog" aria-modal="true" aria-label={`Bill ${bill.number}`}><section className="bill-modal"><header className="bill-modal-head"><div><span>OFFLINE BILL</span><h2>Bill {bill.number}</h2></div><button className="bill-modal-close" onClick={onClose} aria-label="Close bill">×</button></header><div className="bill-modal-content"><div className="bill-info-grid"><div><span>CUSTOMER</span><b>{bill.customerName || 'Prabhu Shri Ram'}</b></div><div><span>PHONE</span><b>{bill.mobile || '—'}</b></div><div><span>CUSTOMER ID</span><b>{bill.customerId || '—'}</b></div><div><span>PAYMENT METHOD</span><b>{bill.payment || 'Cash'}</b></div><div><span>DATE</span><b>{new Date(bill.createdAt).toLocaleString('en-IN', { dateStyle: 'short', timeStyle: 'medium' })}</b></div></div><div className="billing-table-wrap bill-modal-table"><table><thead><tr><th>#</th><th>PRODUCT</th><th>QTY</th><th>MRP</th><th>DISCOUNT</th><th>PRICE</th><th>AMOUNT</th></tr></thead><tbody>{bill.items.map((item, index) => <tr key={`${item.productId}-${index}`}><td>{index + 1}</td><td><b>{item.productName}</b><small>{item.category || 'Product'} · {item.unit}</small></td><td>{item.quantity}</td><td>Rs. {Number(item.mrp).toFixed(0)}</td><td>Rs. {Number(item.discount).toFixed(0)}</td><td>Rs. {Number(item.sellingPrice).toFixed(0)}</td><td><b>Rs. {Number(item.total).toFixed(0)}</b></td></tr>)}</tbody></table></div><div className="bill-payment-summary">{Number(bill.additionalAmount || 0) ? <p>Additional amount <b>Rs. {Number(bill.additionalAmount).toFixed(2)}</b></p> : null}<p>Total bill value <b>Rs. {amount.toFixed(2)}</b></p><p>Paid amount <b>Rs. {paid.toFixed(2)}</b></p><div className={remaining ? 'remaining due' : 'remaining'}><b>Remaining amount</b><strong>Rs. {remaining.toFixed(2)}</strong>{remaining ? <><button onClick={onUpdatePayment}>Update</button><button className="mark-paid" onClick={onMarkPaid}>Paid</button></> : <em>Paid</em>}</div></div><footer><Button variant="secondary" onClick={onClose}>Done</Button></footer></div></section></div>;
}

function ProductQuickEditModal({ product, categories, onClose, onSave, isNew = false }) {
  const [form, setForm] = useState(() => ({ barcode: product.barcode || '', name: product.name || '', category: product.category || 'General', unit: product.unit || 'Pcs', mrp: String(product.mrp ?? ''), discount: String(product.discount ?? 0), sellingPrice: String(product.sellingPrice ?? ''), purchasePrice: product.purchasePrice === '' || product.purchasePrice == null ? '' : String(product.purchasePrice) }));
  const change = (key, value) => setForm(current => {
    const next = { ...current, [key]: value };
    if (key === 'mrp') { const mrp = Math.max(0, Number(value) || 0); next.discount = String(Math.min(Number(current.discount) || 0, mrp)); next.sellingPrice = String(Math.max(0, mrp - Number(next.discount))); }
    if (key === 'discount') {
      // Keep incomplete values such as "." and "0." intact while typing.
      // Normalising them immediately prevents entering values like .5.
      if (/^\d*\.?\d+$/.test(value)) {
        const discount = Math.min(Math.max(0, Number(value)), Number(current.mrp) || 0);
        next.discount = discount === Number(value) ? value : String(discount);
        next.sellingPrice = String(Math.max(0, (Number(current.mrp) || 0) - discount));
      }
    }
    if (key === 'sellingPrice') { const price = Math.min(Math.max(0, Number(value) || 0), Number(current.mrp) || 0); next.sellingPrice = String(price); next.discount = String(Math.max(0, (Number(current.mrp) || 0) - price)); }
    return next;
  });
  const save = () => { if (!form.name.trim() || form.mrp === '') return; onSave({ ...product, barcode: form.barcode.trim(), name: form.name.trim(), category: form.category || 'General', unit: form.unit || 'Pcs', mrp: Number(form.mrp), discount: Number(form.discount) || 0, sellingPrice: Number(form.sellingPrice) || 0, purchasePrice: form.purchasePrice === '' ? '' : Number(form.purchasePrice) }); };
  return <div className="product-quick-edit-overlay" role="dialog" aria-modal="true" aria-label={`${isNew ? 'Add' : 'Edit'} ${product.name || 'product'}`} onMouseDown={event => { if (event.target === event.currentTarget) onClose(); }}><section className="product-quick-edit"><button className="product-quick-close" type="button" onClick={onClose} aria-label="Close product editor">×</button><p>PRODUCT DETAILS</p><div className="product-quick-grid"><Field label="Barcode"><input value={form.barcode} onChange={event => change('barcode', event.target.value)} /></Field><Field label="Product name *"><input value={form.name} onChange={event => change('name', event.target.value)} /></Field><Field label="Category"><select value={form.category} onChange={event => change('category', event.target.value)}>{categories.map(category => <option key={category.id} value={category.name}>{category.name}</option>)}</select></Field><Field label="Unit"><select value={form.unit} onChange={event => change('unit', event.target.value)}>{['Pcs', 'kg', 'g', 'dozen', 'pack', 'L'].map(unit => <option key={unit}>{unit}</option>)}</select></Field><Field label="MRP"><input type="number" min="0" step="0.01" value={form.mrp} onChange={event => change('mrp', event.target.value)} /></Field><Field label="Discount (rs)"><input type="number" min="0" step="0.01" value={form.discount} onChange={event => change('discount', event.target.value)} /></Field><Field label="Selling price"><input type="number" min="0" step="0.01" value={form.sellingPrice} onChange={event => change('sellingPrice', event.target.value)} /></Field><Field label="Purchase price (₹)"><input type="number" min="0" step="0.01" value={form.purchasePrice} onChange={event => change('purchasePrice', event.target.value)} /></Field></div><footer><Button type="button" onClick={save}>{isNew ? 'Add product' : 'Update product'}</Button><Button type="button" variant="secondary" onClick={onClose}>Cancel</Button></footer></section></div>;
}

function DecimalQuantityInput({ label, value, onChange }) {
  const [displayValue, setDisplayValue] = useState(String(value));
  useEffect(() => setDisplayValue(String(value)), [value]);
  const update = nextValue => {
    // Keep partial decimal values in the field so users can type .5 or 0.5
    // naturally, instead of immediately coercing them back to 1.
    if (!/^\d*\.?\d*$/.test(nextValue)) return;
    setDisplayValue(nextValue);
    const quantity = Number(nextValue);
    if (quantity > 0) onChange(quantity);
  };
  const finish = () => {
    const quantity = Number(displayValue);
    if (quantity > 0) {
      setDisplayValue(String(quantity));
      onChange(quantity);
    } else {
      setDisplayValue(String(value));
    }
  };
  const adjustWithArrow = event => {
    if (event.key !== 'ArrowUp' && event.key !== 'ArrowDown') return;
    event.preventDefault();
    const current = Number(displayValue) || Number(value) || 0.01;
    const quantity = event.key === 'ArrowUp' ? Math.floor(current) + 1 : Math.max(0.01, Math.ceil(current) - 1);
    const nextValue = neatNumber(quantity);
    setDisplayValue(nextValue);
    onChange(quantity);
  };
  return <input aria-label={label} type="text" inputMode="decimal" value={displayValue} onChange={event => update(event.target.value)} onKeyDown={adjustWithArrow} onBlur={finish} />;
}

function DecimalAmountInput({ label, value, onChange }) {
  const [displayValue, setDisplayValue] = useState(String(value));
  useEffect(() => setDisplayValue(String(value)), [value]);
  const update = nextValue => {
    // Currency amounts are limited to paise precision: 0 to 2 digits after
    // the decimal point. This prevents values such as 9.950000000000003.
    if (!/^\d*(?:\.\d{0,2})?$/.test(nextValue)) return;
    setDisplayValue(nextValue);
    if (/^\d*\.?\d{1,2}$/.test(nextValue)) onChange(roundAmount(nextValue));
  };
  const finish = () => {
    if (/^\d*\.?\d{1,2}$/.test(displayValue)) onChange(roundAmount(displayValue));
    else setDisplayValue(String(value));
  };
  return <input aria-label={label} type="text" inputMode="decimal" value={displayValue} onChange={event => update(event.target.value)} onBlur={finish} />;
}

function BillEditModal({ bill, products, onEditProduct, onPrint, onClose, onSave }) {
  const { data, setData, setNotice } = useStore();
  const customers = data.customers || [];
  const [customerId, setCustomerId] = useState(bill.customerId || '');
  const [customerName, setCustomerName] = useState(bill.customerName || '');
  const [mobile, setMobile] = useState(bill.mobile || '');
  const [address, setAddress] = useState(bill.address || '');
  const [payment, setPayment] = useState(bill.payment || 'Cash');
  const [paidAmount, setPaidAmount] = useState(String(bill.paidAmount ?? (bill.remainingAmount == null ? bill.grandTotal || 0 : 0)));
  const [additionalAmount, setAdditionalAmount] = useState(String(bill.additionalAmount ?? ''));
  const [showNewProduct, setShowNewProduct] = useState(false);
  const [items, setItems] = useState(() => bill.items.map(item => ({ ...item, quantity: Math.max(0.01, Number(item.quantity) || 1), mrp: Number(item.mrp) || 0, discount: Number(item.discount) || 0, sellingPrice: Number(item.sellingPrice) || 0 })));
  const [editingItemIndex, setEditingItemIndex] = useState(null);
  const [error, setError] = useState('');
  const subtotal = items.reduce((sum, item) => sum + item.mrp * item.quantity, 0);
  const totalDiscount = items.reduce((sum, item) => sum + item.discount * item.quantity, 0);
  const itemTotal = items.reduce((sum, item) => sum + item.sellingPrice * item.quantity, 0);
  const extraAmount = Number(additionalAmount) || 0;
  const grandTotal = Math.max(0, roundBillTotal(itemTotal + extraAmount));
  const fillCustomer = customer => {
    if (!customer) return;
    setCustomerId(customer.customerId || '');
    setCustomerName(customer.name || '');
    setMobile(customer.phone || '');
    setAddress(customer.address || '');
  };
  const findCustomerById = value => {
    const enteredId = value.trim().toLowerCase();
    const enteredNumber = enteredId.match(/\d+$/)?.[0];
    return customers.find(customer => {
      const savedId = String(customer.customerId || '').trim().toLowerCase();
      const savedNumber = savedId.match(/\d+$/)?.[0];
      return savedId === enteredId || (enteredNumber && savedNumber && Number(enteredNumber) === Number(savedNumber));
    });
  };
  const handleCustomerIdChange = value => { setCustomerId(value); if (value.trim()) fillCustomer(findCustomerById(value)); };
  const handleCustomerNameChange = value => {
    setCustomerName(value);
    const matchedCustomer = customers.find(customer => String(customer.name || '').trim().toLowerCase() === value.trim().toLowerCase());
    if (matchedCustomer) fillCustomer(matchedCustomer);
  };
  const changeItem = (index, key, value) => {
    setError('');
    setItems(current => current.map((item, itemIndex) => {
      if (itemIndex !== index) return item;
      const amount = Math.max(0, Number(value) || 0);
      if (key === 'mrp') {
        const discount = Math.min(item.discount, amount);
        return { ...item, mrp: amount, discount, sellingPrice: amount - discount };
      }
      if (key === 'discount') {
        const discount = Math.min(amount, item.mrp);
        return { ...item, discount, sellingPrice: item.mrp - discount };
      }
      if (key === 'sellingPrice') {
        const sellingPrice = Math.min(amount, item.mrp);
        return { ...item, sellingPrice, discount: item.mrp - sellingPrice };
      }
      return { ...item, [key]: amount };
    }));
  };
  const editItem = index => { setError(''); setEditingItemIndex(index); };
  const removeItem = index => {
    setError('');
    setItems(current => current.filter((_, itemIndex) => itemIndex !== index));
    setEditingItemIndex(current => current === index ? null : current !== null && current > index ? current - 1 : current);
  };
  const addProduct = entry => {
    const product = products.find(item => item.id === entry.productId);
    if (!product) return;
    const quantity = Math.max(0.01, Number(entry.quantity) || 1);
    const nextItem = { productId: product.id, productName: product.name, barcode: product.barcode || '', category: product.category || 'General', unit: product.unit || 'Pcs', quantity, mrp: Number(entry.mrp) || 0, purchasePrice: Number(entry.purchasePrice) || 0, discount: Number(entry.discount) || 0, sellingPrice: Number(entry.price) || 0 };
    setError('');
    setItems(current => {
      if (editingItemIndex !== null) return current.map((item, index) => index === editingItemIndex ? nextItem : item);
      const existingIndex = current.findIndex(item => item.productId === product.id);
      if (existingIndex >= 0) return current.map((item, index) => index === existingIndex ? { ...nextItem, quantity: item.quantity + quantity } : item);
      return [...current, nextItem];
    });
    setEditingItemIndex(null);
  };
  const saveNewProduct = product => {
    const barcode = String(product.barcode || '').trim().toLowerCase();
    if (barcode && products.some(item => String(item.barcode || '').trim().toLowerCase() === barcode)) return setNotice({ type: 'error', text: 'This barcode already belongs to a product.' });
    const saved = { ...product, id: uid(), quantity: 1, stock: 0 };
    setData(current => ({ ...current, products: [saved, ...current.products] }));
    fetch(`${apiUrl}/products`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(saved) }).catch(() => {});
    setItems(current => [...current, { productId: saved.id, productName: saved.name, barcode: saved.barcode || '', category: saved.category || 'General', unit: saved.unit || 'Pcs', quantity: 1, mrp: Number(saved.mrp) || 0, purchasePrice: saved.purchasePrice === '' ? '' : Number(saved.purchasePrice) || 0, discount: Number(saved.discount) || 0, sellingPrice: Number(saved.sellingPrice) || 0 }]);
    setShowNewProduct(false);
    setNotice({ type: 'success', text: 'Product added to the catalogue and this bill.' });
  };
  const editedBill = () => {
    const paid = Math.min(grandTotal, Math.max(0, Number(paidAmount) || 0));
    return { ...bill, customerId, customerName, mobile, address, payment, paidAmount: paid, remainingAmount: Math.max(0, grandTotal - paid), subtotal, totalDiscount, additionalAmount: extraAmount, grandTotal, items: items.map(item => ({ ...item, total: item.sellingPrice * item.quantity })) };
  };
  const save = () => {
    if (!items.length) return;
    onSave(editedBill());
  };
  const print = () => onPrint(editedBill());
  return <>{showNewProduct && <ProductQuickEditModal product={{ barcode: '', name: '', category: 'General', unit: 'Pcs', mrp: '', discount: 0, sellingPrice: '', purchasePrice: '' }} categories={data.categories} isNew onClose={() => setShowNewProduct(false)} onSave={saveNewProduct} />}<div className="bill-modal-overlay" role="dialog" aria-modal="true" aria-label={`Edit bill ${bill.number}`}><section className="bill-modal bill-edit-modal"><header className="bill-modal-head"><div><span>EDIT BILL</span><h2>Bill {bill.number}</h2></div><button className="bill-modal-close" onClick={onClose} aria-label="Close bill editor">×</button></header><div className="bill-modal-content"><div className="bill-edit-fields"><Field label="Customer ID"><input list="bill-edit-customer-ids" placeholder="Enter customer ID" value={customerId} onChange={event => handleCustomerIdChange(event.target.value)} /></Field><Field label="Customer name"><input list="bill-edit-customer-names" placeholder="Enter customer name" value={customerName} onChange={event => handleCustomerNameChange(event.target.value)} /></Field><Field label="Phone number"><input value={mobile} onChange={event => setMobile(event.target.value)} /></Field><Field label="Address"><input value={address} onChange={event => setAddress(event.target.value)} /></Field><Field label="Payment method"><select value={payment} onChange={event => setPayment(event.target.value)}><option>Cash</option><option>UPI</option><option>Card</option></select></Field></div><datalist id="bill-edit-customer-ids">{customers.map(customer => <option key={customer.customerId} value={customer.customerId}>{customer.name}</option>)}</datalist><datalist id="bill-edit-customer-names">{customers.map(customer => <option key={customer.customerId} value={customer.name}>{customer.customerId}</option>)}</datalist><BillEditProductEntry products={products} itemToEdit={editingItemIndex === null ? null : items[editingItemIndex]} isEditingItem={editingItemIndex !== null} onAdd={addProduct} onEditProduct={onEditProduct} onCreateProduct={() => setShowNewProduct(true)} /><div className="billing-table-wrap bill-modal-table"><table><thead><tr><th>PRODUCT</th><th>QTY</th><th>MRP</th><th>DISCOUNT</th><th>PRICE</th><th>AMOUNT</th><th></th></tr></thead><tbody>{items.map((item, index) => <tr key={`${item.productId}-${index}`}><td><b>{item.productName}</b><small>{item.barcode || 'No barcode'} · {item.unit || 'Pcs'}</small></td><td><DecimalQuantityInput label={`Quantity for ${item.productName}`} value={item.quantity} onChange={quantity => changeItem(index, 'quantity', quantity)} /></td><td><input aria-label={`MRP for ${item.productName}`} type="number" min="0" step="0.01" value={item.mrp} onChange={event => changeItem(index, 'mrp', event.target.value)} /></td><td><DecimalAmountInput label={`Discount for ${item.productName}`} value={item.discount} onChange={discount => changeItem(index, 'discount', discount)} /></td><td><input aria-label={`Price for ${item.productName}`} type="number" min="0" step="0.01" value={item.sellingPrice} onChange={event => changeItem(index, 'sellingPrice', event.target.value)} /></td><td><b>Rs. {(item.sellingPrice * item.quantity).toFixed(2)}</b></td><td><button type="button" className="bill-edit" onClick={() => editItem(index)}>Edit</button><button type="button" className="bill-remove" onClick={() => removeItem(index)}>Remove</button></td></tr>)}</tbody></table></div>{error && <p className="bill-edit-empty">{error}</p>}{!items.length && <p className="bill-edit-empty">Add at least one item before saving this bill.</p>}<div className="bill-edit-summary"><p><span>Item total</span><b>Rs. {itemTotal.toFixed(2)}</b></p><p><span>Additional amount</span><b>Rs. {extraAmount.toFixed(2)}</b></p></div><div className="bill-edit-total"><span>Total bill value</span><b>Rs. {grandTotal.toFixed(2)}</b></div><div className="bill-edit-paid"><Field label="Additional amount"><input type="number" min="0" step="0.01" placeholder="0.00" value={additionalAmount} onChange={event => setAdditionalAmount(event.target.value)} /></Field><Field label="Paid amount"><input type="number" min="0" max={grandTotal} value={paidAmount} onChange={event => setPaidAmount(event.target.value)} /></Field><Button type="button" variant="secondary" onClick={() => setPaidAmount(grandTotal.toFixed(2))}>Paid</Button><span>Remaining: <b>Rs. {Math.max(0, grandTotal - (Number(paidAmount) || 0)).toFixed(2)}</b></span></div><footer><Button type="button" variant="secondary" onClick={onClose}>Cancel</Button><Button type="button" variant="secondary" onClick={print}><Printer size={16} />Print</Button><Button type="button" onClick={save} disabled={!items.length}>Save changes</Button></footer></div></section></div></>;
}

function BillEditProductEntry({ products, itemToEdit, isEditingItem, onAdd, onEditProduct, onCreateProduct }) {
  const empty = { barcode: '', name: '', quantity: '1', mrp: '', discount: '', price: '', purchasePrice: '', productId: '' };
  const [entry, setEntry] = useState(empty);
  const [activeSuggestionIndex, setActiveSuggestionIndex] = useState(-1);
  const [isProductInputFocused, setIsProductInputFocused] = useState(false);
  useEffect(() => {
    if (!itemToEdit) return;
    setActiveSuggestionIndex(-1);
    setIsProductInputFocused(false);
    setEntry({ barcode: itemToEdit.barcode || '', name: itemToEdit.productName, quantity: String(itemToEdit.quantity), mrp: String(itemToEdit.mrp), discount: String(itemToEdit.discount), price: String(itemToEdit.sellingPrice), purchasePrice: String(itemToEdit.purchasePrice ?? ''), productId: itemToEdit.productId });
    requestAnimationFrame(() => {
      document.querySelector('.bill-edit-add-product')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
      focusQuantity();
    });
  }, [itemToEdit]);
  useEffect(() => {
    const input = document.querySelector('.bill-edit-add-product input[type="number"][min="1"]');
    if (input) { input.min = '0'; input.step = '1'; }
  }, [entry.quantity]);
  // Keep all matching products available, just like the New bill search. The
  // suggestion menu has its own scrollbar, so longer catalogues remain usable.
  const matches = entry.name.trim()
    ? products.filter(product => `${product.name} ${product.barcode || ''}`.toLowerCase().includes(entry.name.trim().toLowerCase()))
    : [];
  useEffect(() => {
    const options = document.querySelectorAll('.bill-edit-product-name .billing-product-suggestions button');
    options[activeSuggestionIndex]?.scrollIntoView({ block: 'nearest', behavior: 'auto' });
  }, [activeSuggestionIndex, matches]);
  const findProduct = value => {
    const key = String(value ?? '').trim().toLowerCase();
    if (!key) return undefined;
    return products.find(product => String(product.barcode || '').trim().toLowerCase() === key || String(product.name || '').trim().toLowerCase() === key);
  };
  const findProductByBarcode = value => {
    const barcode = String(value ?? '').trim().toLowerCase();
    if (!barcode) return undefined;
    return products.find(product => String(product.barcode || '').trim().toLowerCase() === barcode);
  };
  const focusQuantity = () => requestAnimationFrame(() => {
    const quantityInput = document.querySelector('.bill-edit-add-product input[type="number"]');
    quantityInput?.focus();
    quantityInput?.select();
  });
  const selectProduct = (product, shouldFocusQuantity = true) => {
    setActiveSuggestionIndex(-1);
    setIsProductInputFocused(false);
    setEntry({ barcode: product.barcode || '', name: product.name, quantity: '1', mrp: String(product.mrp || ''), discount: String(product.discount || 0), price: String(product.sellingPrice || product.mrp || ''), purchasePrice: String(product.purchasePrice || ''), productId: product.id });
    if (shouldFocusQuantity) focusQuantity();
  };
  const change = (key, value) => {
    const isProductName = key === 'name';
    if (key === 'name') setActiveSuggestionIndex(-1);
    if (isProductName && !value.trim()) {
      setEntry({ ...empty, [key]: value });
      return;
    }
    const product = isProductName ? findProduct(value) : null;
    setEntry(current => {
      if (product) return { ...current, barcode: product.barcode || '', name: product.name, mrp: String(product.mrp || ''), discount: String(product.discount || 0), price: String(product.sellingPrice || product.mrp || ''), purchasePrice: String(product.purchasePrice ?? ''), productId: product.id };
      if (isProductName) return { ...current, name: value, productId: '' };
      return (() => {
          if (key === 'mrp') {
            const mrp = Math.max(0, Number(value) || 0), discount = Math.min(Math.max(0, Number(current.discount) || 0), mrp);
            return { ...current, mrp: String(mrp), discount: String(discount), price: String(mrp - discount) };
          }
          if (key === 'discount') {
            // Preserve partial decimal typing (for example "." before ".5").
            if (!/^\d*\.?\d+$/.test(value)) return { ...current, discount: value };
            const mrp = Math.max(0, Number(current.mrp) || 0), discount = Math.min(Math.max(0, Number(value)), mrp);
            return { ...current, discount: discount === Number(value) ? value : String(discount), price: String(mrp - discount) };
          }
          if (key === 'price') {
            const mrp = Math.max(0, Number(current.mrp) || 0), price = Math.min(Math.max(0, Number(value) || 0), mrp);
            return { ...current, price: String(price), discount: String(mrp - price) };
          }
          return { ...current, [key]: value };
        })();
    });
  };
  const handleBarcodeChange = value => {
    const product = findProductByBarcode(value);
    if (product) {
      // Fill the product details while keeping the cursor in the barcode field.
      // This lets a user continue from "121" to "121s" without interruption.
      selectProduct(product, false);
      return;
    }
    setEntry(current => ({ ...empty, barcode: value, quantity: current.quantity || '1' }));
  };
  const resolveScannedBarcode = value => {
    const product = findProductByBarcode(value);
    if (product) selectProduct(product);
    else focusQuantity();
  };
  const focusBarcode = () => requestAnimationFrame(() => {
    const barcodeInput = document.querySelector('.bill-edit-add-product input[placeholder="Scan or enter barcode"]');
    barcodeInput?.focus();
    barcodeInput?.select();
  });
  const add = () => { const product = products.find(item => item.id === entry.productId) || findProduct(entry.barcode) || findProduct(entry.name); if (!product) return; onAdd({ ...entry, productId: product.id }); setEntry(empty); focusBarcode(); };
  const handleProductNameKeyDown = event => {
    if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
      if (!matches.length) return;
      event.preventDefault();
      setActiveSuggestionIndex(index => event.key === 'ArrowDown' ? (index + 1) % matches.length : index < 0 ? matches.length - 1 : (index - 1 + matches.length) % matches.length);
    } else if (event.key === 'Enter') {
      event.preventDefault();
      if (activeSuggestionIndex >= 0) selectProduct(matches[activeSuggestionIndex]);
      else {
        const product = findProduct(event.currentTarget.value);
        if (product) selectProduct(product);
        else focusQuantity();
      }
    } else if (event.key === 'Escape') {
      setActiveSuggestionIndex(-1);
    }
  };
  const editProduct = () => onEditProduct(products.find(product => product.id === entry.productId) || findProduct(entry.barcode) || findProduct(entry.name));
  return <section className="bill-edit-add-product"><div className="bill-edit-add-product-head"><h3>Add product</h3><div className="bill-edit-product-actions"><Button type="button" onClick={onCreateProduct}>Add product</Button><Button type="button" variant="secondary" onClick={editProduct}>Edit product</Button></div></div><div className="bill-edit-product-grid"><Field label="Barcode"><input placeholder="Scan or enter barcode" value={entry.barcode} onFocus={event => event.currentTarget.select()} onChange={event => handleBarcodeChange(event.target.value)} onKeyDown={event => { if (event.key === 'Enter') { event.preventDefault(); resolveScannedBarcode(event.currentTarget.value); } }} onBlur={event => { const product = findProductByBarcode(event.currentTarget.value); if (product) selectProduct(product, false); }} /></Field><div className="bill-edit-product-name"><Field label="Product name"><input placeholder="Search product" value={entry.name} onFocus={() => setIsProductInputFocused(true)} onBlur={() => setIsProductInputFocused(false)} onChange={event => change('name', event.target.value)} onKeyDown={handleProductNameKeyDown} aria-autocomplete="list" aria-controls="bill-edit-product-suggestions" aria-activedescendant={activeSuggestionIndex >= 0 ? `bill-edit-product-suggestion-${matches[activeSuggestionIndex]?.id}` : undefined} />{isProductInputFocused && matches.length > 0 && <div id="bill-edit-product-suggestions" className="billing-product-suggestions" role="listbox">{matches.map((product, index) => <button type="button" id={`bill-edit-product-suggestion-${product.id}`} role="option" aria-selected={index === activeSuggestionIndex} className={index === activeSuggestionIndex ? 'active' : ''} key={product.id} onMouseDown={event => event.preventDefault()} onMouseEnter={() => setActiveSuggestionIndex(index)} onClick={() => selectProduct(product)}><span className="billing-suggestion-row"><b>{product.name}</b><small>MRP: Rs. {Number(product.mrp || 0).toFixed(0)}</small></span></button>)}</div>}</Field></div><Field label="Quantity"><input type="number" min="1" value={entry.quantity} onChange={event => change('quantity', event.target.value)} onKeyDown={event => { if (event.key === 'Enter') { event.preventDefault(); add(); } }} /></Field><Field label="MRP"><input type="number" min="0" step="0.01" value={entry.mrp} onChange={event => change('mrp', event.target.value)} /></Field><Field label="Discount"><input type="number" min="0" step="0.01" value={entry.discount} onChange={event => change('discount', event.target.value)} /></Field><Field label="Price"><input type="number" min="0" step="0.01" value={entry.price} onChange={event => change('price', event.target.value)} /></Field><Field label="Amount"><input readOnly value={`Rs. ${((Number(entry.price) || 0) * (Number(entry.quantity) || 1)).toFixed(2)}`} /></Field><Button type="button" onClick={add} disabled={!entry.productId}>{isEditingItem ? 'Update product' : 'Add product'}</Button></div><p className="bill-edit-purchase-price">Purchase price: Rs. {(Number(entry.purchasePrice) || 0).toFixed(2)}</p></section>;
}

function PaymentUpdateModal({ bill, onClose, onSave }) {
  const total = Number(bill.grandTotal || 0);
  const alreadyPaid = Number(bill.paidAmount ?? (bill.remainingAmount == null ? total : 0));
  const [amount, setAmount] = useState(String(alreadyPaid));
  const save = () => { const value = Number(amount); if (Number.isFinite(value) && value >= 0 && value <= total) onSave(value); };
  return <div className="payment-update-overlay" role="dialog" aria-modal="true"><section className="payment-update-modal"><button className="bill-modal-close" onClick={onClose} aria-label="Close payment update">×</button><span>PAYMENT UPDATE</span><h2>Update collected amount</h2><p>Bill <b>#{bill.number}</b> · {bill.customerName || 'Prabhu Shri Ram'}</p><div className="payment-stats"><div><small>Bill total</small><b>Rs. {total.toFixed(0)}</b></div><div><small>Amount already paid</small><b>Rs. {alreadyPaid.toFixed(0)}</b></div></div><label>Paid amount<input autoFocus type="number" min="0" max={total} value={amount} onChange={event => setAmount(event.target.value)} onKeyDown={event => { if (event.key === 'Enter') save(); }} /></label><small className="payment-help">Enter the total amount collected so far. Remaining balance will update automatically.</small><footer><Button variant="secondary" onClick={onClose}>Cancel</Button><Button onClick={save}>Update payment</Button></footer></section></div>;
}

function BulkUploadModal({ onClose, onFile, report, onImport }) {
  const [dragging, setDragging] = useState(false); const inputId = 'bulk-product-file';
  const template = () => { const csv = 'barcode,product name,category,quantity,unit,mrp,purchasing price,discount,selling price'; const link = document.createElement('a'); link.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv' })); link.download = 'product-upload-template.csv'; link.click(); URL.revokeObjectURL(link.href) };
  const selectFile = file => { if (file) onFile(file) };
  return <div className="upload-overlay" role="dialog" aria-modal="true"><div className={`upload-modal ${report ? 'upload-modal-report' : ''}`}><button className="upload-close" onClick={onClose} aria-label="Close"><X size={19} /></button><h2>{report ? 'Upload validation report' : 'Bulk product upload'}</h2>{report ? <div className="upload-report"><div className="report-file"><FileSpreadsheet size={20} /><span><b>{report.fileName}</b><small>Existing products with the same barcode or product name will be updated.</small></span></div>{report.headerError ? <div className="report-header-error"><AlertTriangle size={19} /><span><b>File structure needs attention</b>{report.headerError}</span></div> : <><div className="report-summary"><div className="summary-card total"><span>Total rows</span><b>{report.rows.length}</b></div><div className="summary-card valid"><CheckCircle2 size={18} /><span>New products</span><b>{report.additions.length}</b></div><div className="summary-card updating"><Pencil size={18} /><span>Products to update</span><b>{report.updates.length}</b></div><div className="summary-card invalid"><XCircle size={18} /><span>Needs correction</span><b>{report.invalid.length}</b></div></div><div className="validation-list">{report.rows.map(row => <div className={`validation-row ${row.errors.length ? 'invalid' : row.operation === 'update' ? 'updating' : 'valid'}`} key={row.rowNumber}><div className="row-status">{row.errors.length ? <XCircle size={19} /> : row.operation === 'update' ? <Pencil size={18} /> : <CheckCircle2 size={19} />}</div><div className="row-details"><div><b>Excel row {row.rowNumber}</b><span>{row.product}</span><small>Barcode: {row.barcode}</small></div>{row.errors.length ? <ul>{row.errors.map(error => <li key={error}>{error}</li>)}</ul> : <p>{row.operation === 'update' ? 'Existing product — its details will be updated.' : 'New product — ready to add to inventory.'}</p>}</div><span className="status-pill">{row.errors.length ? 'Fix required' : row.operation === 'update' ? 'Will update' : 'New product'}</span></div>)}</div><div className="report-actions"><label htmlFor={inputId} className="replace-file">Choose another file</label><Button onClick={onImport} disabled={!report.valid.length}><CheckCircle2 size={16} />Import {report.valid.length} valid row{report.valid.length === 1 ? '' : 's'}</Button></div></>}</div> : <><div className={`drop-zone ${dragging ? 'dragging' : ''}`} onDragOver={e => { e.preventDefault(); setDragging(true) }} onDragLeave={() => setDragging(false)} onDrop={e => { e.preventDefault(); setDragging(false); selectFile(e.dataTransfer.files[0]) }}><strong>Drag & drop your Excel file here</strong><span>or</span><label htmlFor={inputId} className="browse-file">Browse Excel</label><small>Barcode and category are optional. Same barcode or product name updates the existing product; otherwise a new product is added.</small></div><button className="template-button" onClick={template}>Download Excel template</button></>}<input id={inputId} className="file-picker" type="file" accept=".csv,.xlsx,.xls" onChange={e => { selectFile(e.target.files[0]); e.target.value = '' }} /></div></div>
}
