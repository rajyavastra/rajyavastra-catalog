// Converts data/products.xlsx (sheet "Products") into public/assets/products.json
// Usage: npm run sync
import fs from 'node:fs';
import path from 'node:path';
import ExcelJS from 'exceljs';

const root = path.resolve(import.meta.dirname, '..');
const XLSX_PATH = process.env.PRODUCTS_XLSX || path.join(root, 'data', 'products.xlsx');
const OUT_PATH = path.join(root, 'public', 'assets', 'products.json');
const IMG_ROOT = path.join(root, 'public', 'assets', 'products');

// Excel cells can hold numbers, text, rich text, formulas or hyperlinks: reduce all to plain values.
const plain = (v) => {
  if (v == null) return '';
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map((t) => t.text).join('');
    if ('result' in v) return plain(v.result);
    if (v.text) return String(v.text);
  }
  return v;
};
const text = (v) => String(plain(v)).trim();

const wb = new ExcelJS.Workbook();
await wb.xlsx.readFile(XLSX_PATH);
const ws = wb.getWorksheet('Products');
if (!ws) throw new Error('Sheet named "Products" not found in ' + XLSX_PATH);

const headers = ws.getRow(1).values.slice(1).map((h) => text(h).toLowerCase());
const required = ['id', 'name', 'price'];
for (const r of required) {
  if (!headers.includes(r)) throw new Error(`Missing required column "${r}"`);
}

const products = [];
const seen = new Set();
const warnings = [];

ws.eachRow((row, n) => {
  if (n === 1) return;
  const rec = Object.fromEntries(headers.map((h, i) => [h, plain(row.getCell(i + 1).value)]));
  const id = text(rec.id);
  if (!id && !text(rec.name)) return; // blank row
  if (!id || !text(rec.name)) {
    warnings.push(`Row ${n}: skipped, id and name are both required`);
    return;
  }
  if (seen.has(id)) {
    warnings.push(`Row ${n}: skipped, duplicate id "${id}"`);
    return;
  }
  const price = Number(String(rec.price).replace(/[^0-9.]/g, ''));
  if (!Number.isFinite(price) || rec.price === '') {
    warnings.push(`Row ${n} (${id}): price is not a number`);
  }

  const images = text(rec.images)
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean)
    .filter((file) => {
      const ok = fs.existsSync(path.join(IMG_ROOT, id, file));
      if (!ok) warnings.push(`Row ${n} (${id}): image not found: products/${id}/${file}`);
      return ok;
    })
    .map((file) => `assets/products/${id}/${file}`);

  seen.add(id);
  products.push({
    id,
    name: text(rec.name),
    category: text(rec.category) || 'Other',
    fabric: text(rec.fabric),
    price: Number.isFinite(price) ? price : 0,
    description: text(rec.description),
    inStock: !['no', 'false', '0', 'n'].includes(text(rec.in_stock).toLowerCase()),
    images,
  });
});

if (products.length === 0) throw new Error('No products found. Refusing to write an empty catalog.');

fs.mkdirSync(path.dirname(OUT_PATH), { recursive: true });
fs.writeFileSync(OUT_PATH, JSON.stringify(products, null, 2) + '\n');
console.log(`products.json written: ${products.length} products`);
warnings.forEach((w) => console.warn('  warning: ' + w));
