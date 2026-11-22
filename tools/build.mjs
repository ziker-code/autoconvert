// Build: salin SheetJS ke ekstensi, buat ikon, buat template Excel, zip ekstensi.
// Jalankan: npm run build
import fs from "node:fs";
import path from "node:path";
import zlib from "node:zlib";
import { fileURLToPath } from "node:url";
import ExcelJS from "exceljs";
import { ZipArchive } from "archiver";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const EXT = path.join(ROOT, "extension");
const ASSETS = path.join(ROOT, "website", "assets");
fs.mkdirSync(path.join(EXT, "vendor"), { recursive: true });
fs.mkdirSync(path.join(EXT, "icons"), { recursive: true });
fs.mkdirSync(ASSETS, { recursive: true });

// 1. SheetJS ------------------------------------------------------------
fs.copyFileSync(
  path.join(ROOT, "node_modules", "xlsx", "dist", "xlsx.full.min.js"),
  path.join(EXT, "vendor", "xlsx.full.min.js")
);
console.log("✓ vendor/xlsx.full.min.js");

// Lisensi ikut didistribusikan bersama ekstensi: MIT untuk proyek ini, Apache-2.0 untuk SheetJS
fs.copyFileSync(path.join(ROOT, "LICENSE"), path.join(EXT, "LICENSE"));
fs.copyFileSync(path.join(ROOT, "node_modules", "xlsx", "LICENSE"), path.join(EXT, "vendor", "LICENSE-sheetjs.txt"));
console.log("✓ lisensi");

// 2. Ikon (kotak biru membulat + petir putih) ---------------------------
function crc32(buf) {
  let c, crc = 0xffffffff;
  for (let n = 0; n < buf.length; n++) {
    c = (crc ^ buf[n]) & 0xff;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    crc = (crc >>> 8) ^ c;
  }
  return (crc ^ 0xffffffff) >>> 0;
}
function chunk(type, data) {
  const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
  const td = Buffer.concat([Buffer.from(type), data]);
  const crc = Buffer.alloc(4); crc.writeUInt32BE(crc32(td));
  return Buffer.concat([len, td, crc]);
}
function pointInPoly(x, y, poly) {
  let inside = false;
  for (let i = 0, j = poly.length - 1; i < poly.length; j = i++) {
    const [xi, yi] = poly[i], [xj, yj] = poly[j];
    if (yi > y !== yj > y && x < ((xj - xi) * (y - yi)) / (yj - yi) + xi) inside = !inside;
  }
  return inside;
}
function makeIcon(size) {
  const bolt = [[0.58, 0.12], [0.26, 0.56], [0.48, 0.56], [0.40, 0.88], [0.74, 0.42], [0.52, 0.42], [0.62, 0.12]];
  const r = 0.22, SS = 4;
  const raw = Buffer.alloc(size * (size * 4 + 1));
  for (let y = 0; y < size; y++) {
    raw[y * (size * 4 + 1)] = 0;
    for (let x = 0; x < size; x++) {
      let bg = 0, fg = 0;
      for (let sy = 0; sy < SS; sy++) for (let sx = 0; sx < SS; sx++) {
        const u = (x + (sx + 0.5) / SS) / size, v = (y + (sy + 0.5) / SS) / size;
        const dx = Math.max(r - u, 0, u - (1 - r)), dy = Math.max(r - v, 0, v - (1 - r));
        if (dx * dx + dy * dy <= r * r) { bg++; if (pointInPoly(u, v, bolt)) fg++; }
      }
      const n = SS * SS, a = bg / n, f = bg ? fg / bg : 0;
      const o = y * (size * 4 + 1) + 1 + x * 4;
      raw[o] = Math.round(11 + (255 - 11) * f); // biru tua #0B2F5B -> putih
      raw[o + 1] = Math.round(47 + (255 - 47) * f);
      raw[o + 2] = Math.round(91 + (255 - 91) * f);
      raw[o + 3] = Math.round(a * 255);
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0); ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; ihdr[9] = 6; ihdr[10] = 0; ihdr[11] = 0; ihdr[12] = 0;
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", ihdr), chunk("IDAT", zlib.deflateSync(raw)), chunk("IEND", Buffer.alloc(0)),
  ]);
}
for (const s of [16, 32, 48, 128]) fs.writeFileSync(path.join(EXT, "icons", `icon${s}.png`), makeIcon(s));
fs.copyFileSync(path.join(EXT, "icons", "icon128.png"), path.join(ROOT, "website", "assets", "icon.png"));
fs.copyFileSync(path.join(EXT, "icons", "icon32.png"), path.join(ROOT, "website", "assets", "favicon.png"));
console.log("✓ ikon");

// 3. Template Excel -----------------------------------------------------
const JENIS = ["Berita Acara Kegiatan", "Berita Acara Pembimbingan", "Berita Acara Ujian"];
const TIPE = ["Offline", "Online", "Hybrid"];
const KOLOM = [
  { header: "No", key: "no", width: 6 },
  { header: "Tanggal", key: "tanggal", width: 14, note: "Format DD/MM/YYYY, contoh 07/07/2025" },
  { header: "Jam Mulai", key: "mulai", width: 11, note: "Format HH:MM (24 jam), contoh 08:00" },
  { header: "Jam Selesai", key: "selesai", width: 12, note: "Format HH:MM (24 jam), harus setelah Jam Mulai" },
  { header: "Jenis Kegiatan", key: "jenis", width: 28, note: "Pilih dari daftar. Kosong = Berita Acara Kegiatan" },
  { header: "Tipe Penyelenggaraan", key: "tipe", width: 21, note: "Pilih: Offline / Online / Hybrid" },
  { header: "Lokasi", key: "lokasi", width: 30, note: "Contoh: Kampus IPB Dramaga, Bogor" },
  { header: "Keterangan", key: "keterangan", width: 60, note: "Uraian singkat kegiatan yang dilakukan" },
  { header: "Nama File Bukti", key: "bukti", width: 22, note: "Nama file foto/PDF bukti, contoh hari1.jpg" },
];
const CONTOH = [
  [1, "07/07/2025", "07:00", "12:00", JENIS[0], TIPE[0], "Graha Widya Wisuda IPB, Dramaga", "Pembukaan MPKMB dan pengenalan IPB University", "hari1.jpg"],
  [2, "08/07/2025", "08:00", "11:30", JENIS[0], TIPE[1], "Zoom Meeting", "Materi kepemimpinan dan softskill mahasiswa baru", "hari2.jpg"],
  [3, "09/07/2025", "13:00", "16:00", JENIS[1], TIPE[2], "Kampus IPB Dramaga, Bogor", "Mentoring kelompok bersama pendamping", "hari3.png"],
];

const wb = new ExcelJS.Workbook();
wb.creator = "AutoConvert Logbook IPB";
const BIRU = "FF0B2F5B";

function styleHeader(ws) {
  const row = ws.getRow(1);
  row.height = 24;
  row.eachCell((c) => {
    c.font = { bold: true, color: { argb: "FFFFFFFF" } };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: BIRU } };
    c.alignment = { vertical: "middle", horizontal: "center" };
  });
  ws.views = [{ state: "frozen", ySplit: 1 }];
}

// Sheet Logbook (diisi pengguna)
const ws = wb.addWorksheet("Logbook", { properties: { tabColor: { argb: BIRU } } });
ws.columns = KOLOM.map(({ header, key, width }) => ({ header, key, width }));
styleHeader(ws);
KOLOM.forEach((k, i) => { if (k.note) ws.getCell(1, i + 1).note = k.note; });
const BARIS = 500;
for (let r = 2; r <= BARIS + 1; r++) {
  // tanggal & jam sebagai teks agar tidak diubah Excel
  for (const c of [2, 3, 4, 9]) ws.getCell(r, c).numFmt = "@";
  ws.getCell(r, 8).alignment = { wrapText: true, vertical: "top" };
  ws.getCell(r, 5).dataValidation = {
    type: "list", allowBlank: true, formulae: ['"' + JENIS.join(",") + '"'],
    showErrorMessage: true, errorTitle: "Pilihan tidak valid", error: "Pilih jenis kegiatan dari daftar.",
  };
  ws.getCell(r, 6).dataValidation = {
    type: "list", allowBlank: true, formulae: ['"' + TIPE.join(",") + '"'],
    showErrorMessage: true, errorTitle: "Pilihan tidak valid", error: "Pilih Offline, Online, atau Hybrid.",
  };
  ws.getCell(r, 2).dataValidation = {
    type: "custom", allowBlank: true,
    formulae: [`OR(ISNUMBER(B${r}),AND(LEN(B${r})>=8,ISNUMBER(DATEVALUE(B${r}))),AND(LEN(B${r})=10,MID(B${r},3,1)="/",MID(B${r},6,1)="/"))`],
    showErrorMessage: true, errorStyle: "warning", errorTitle: "Cek format tanggal", error: "Gunakan format DD/MM/YYYY, contoh 07/07/2025.",
  };
}

// Sheet Contoh
const wc = wb.addWorksheet("Contoh");
wc.columns = KOLOM.map(({ header, key, width }) => ({ header, key, width }));
styleHeader(wc);
CONTOH.forEach((r) => wc.addRow(r));
wc.eachRow((row, i) => { if (i > 1) row.getCell(8).alignment = { wrapText: true, vertical: "top" }; });

// Sheet Petunjuk
const wp = wb.addWorksheet("Petunjuk");
wp.getColumn(1).width = 4;
wp.getColumn(2).width = 100;
const petunjuk = [
  ["", "PETUNJUK PENGISIAN — AutoConvert Logbook IPB"],
  ["", ""],
  ["1", 'Isi data di sheet "Logbook" mulai baris ke-2. Satu baris = satu log aktivitas. Jangan ubah judul kolom.'],
  ["2", "Tanggal: format DD/MM/YYYY (contoh 07/07/2025). Harus berada di dalam periode aktivitas."],
  ["3", "Jam Mulai & Jam Selesai: format 24 jam HH:MM (contoh 08:00 dan 12:30). Durasi dihitung otomatis oleh portal."],
  ["4", "Jenis Kegiatan: pilih dari dropdown. Untuk ospek/MPKMB biasanya \"Berita Acara Kegiatan\"."],
  ["5", "Tipe Penyelenggaraan: Offline, Online, atau Hybrid."],
  ["6", "Lokasi dan Keterangan wajib diisi. Tulis keterangan yang jelas dan sesuai kegiatan yang benar-benar diikuti."],
  ["7", "Nama File Bukti: tulis nama file foto/PDF persis seperti di laptop (contoh hari1.jpg). Simpan semua foto di satu folder."],
  ["8", "Format bukti yang diterima: JPG, JPEG, PNG, PDF. Foto besar akan dikompres otomatis; PDF maksimal 10 MB."],
  ["9", 'Lihat sheet "Contoh" untuk contoh pengisian. Sheet "Contoh" TIDAK ikut dikirim.'],
  ["", ""],
  ["", "Setelah selesai: buka halaman Log Aktivitas di Student Portal → klik ⚡ AutoConvert Logbook → pilih Excel & foto → cek → Kirim."],
];
petunjuk.forEach((r) => wp.addRow(r));
wp.getCell("B1").font = { bold: true, size: 14, color: { argb: BIRU } };
wp.eachRow((row) => { row.getCell(2).alignment = { wrapText: true, vertical: "top" }; row.getCell(1).alignment = { vertical: "top" }; });

const templatePath = path.join(ASSETS, "template-logbook-autoconvert.xlsx");
await wb.xlsx.writeFile(templatePath);
console.log("✓ template:", path.relative(ROOT, templatePath));

// 4. Zip ekstensi --------------------------------------------------------
const zipPath = path.join(ASSETS, "autoconvert-extension.zip");
await new Promise((resolve, reject) => {
  const out = fs.createWriteStream(zipPath);
  const zip = new ZipArchive({ zlib: { level: 9 } });
  out.on("close", resolve);
  zip.on("error", reject);
  zip.pipe(out);
  zip.directory(EXT, "autoconvert-extension");
  zip.finalize();
});
console.log("✓ zip:", path.relative(ROOT, zipPath), `(${(fs.statSync(zipPath).size / 1024).toFixed(0)} KB)`);

// 5. Userscript untuk iPhone (Safari + aplikasi Userscripts / Tampermonkey) ----------
// Isinya kode yang sama dengan ekstensi; hanya akses Google Drive yang lewat GM.xmlHttpRequest.
// Alamat website untuk @updateURL (pembaruan otomatis di iPhone): SITE_URL, atau otomatis
// dari domain produksi Vercel saat di-build di Vercel.
const versi = JSON.parse(fs.readFileSync(path.join(EXT, "manifest.json"), "utf8")).version;
const siteUrl = (
  process.env.SITE_URL ||
  (process.env.VERCEL_PROJECT_PRODUCTION_URL ? `https://${process.env.VERCEL_PROJECT_PRODUCTION_URL}` : "")
).replace(/\/+$/, "");
const meta = [
  "// ==UserScript==",
  "// @name         AutoConvert Logbook IPB",
  "// @namespace    autoconvert-logbook-ipb",
  `// @version      ${versi}`,
  "// @description  Isi Log Aktivitas Student Portal IPB otomatis dari Excel + foto (lokal / Google Drive).",
  "// @match        https://studentportal.ipb.ac.id/*",
  "// @exclude      https://studentportal.ipb.ac.id/Account/*",
  "// @run-at       document-idle",
  "// @grant        GM.xmlHttpRequest",
  "// @grant        GM_xmlhttpRequest",
  "// @connect      drive.google.com",
  "// @connect      drive.usercontent.google.com",
  "// @connect      googleusercontent.com",
  "// @noframes",
  "// @license      MIT (AutoConvert); berisi SheetJS CE, Apache-2.0 - https://sheetjs.com",
  "// @homepageURL  https://github.com/ziker-code/autoconvert",
  ...(siteUrl
    ? [`// @downloadURL  ${siteUrl}/assets/autoconvert.user.js`, `// @updateURL    ${siteUrl}/assets/autoconvert.user.js`]
    : []),
  "// ==/UserScript==",
  "",
].join("\n");
const bagian = [
  ["SheetJS (xlsx.mini)", path.join(ROOT, "node_modules", "xlsx", "dist", "xlsx.mini.min.js")],
  ["parser.js", path.join(EXT, "parser.js")],
  ["drive.js", path.join(EXT, "drive.js")],
  ["host.js", path.join(ROOT, "userscript", "host.js")],
  ["content.js", path.join(EXT, "content.js")],
];
const isi = bagian
  .map(([nama, p]) => `\n/* ===== ${nama} ===== */\n` + fs.readFileSync(p, "utf8").replace("__VERSI__", versi))
  .join("\n");
const userscriptPath = path.join(ASSETS, "autoconvert.user.js");
fs.writeFileSync(userscriptPath, meta + isi);
console.log("✓ userscript:", path.relative(ROOT, userscriptPath), `(${(fs.statSync(userscriptPath).size / 1024).toFixed(0)} KB)`);
