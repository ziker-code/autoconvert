// Tes parser: node tools/test-parser.mjs
import fs from "node:fs";
import path from "node:path";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";
import ExcelJS from "exceljs";
import * as XLSX from "xlsx";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
new Function(fs.readFileSync(path.join(ROOT, "extension", "parser.js"), "utf8"))();
const P = globalThis.AutoKonvertParser;

let lulus = 0;
const t = (nama, fn) => { fn(); lulus++; console.log("  ✓", nama); };

console.log("Fungsi dasar");
t("parseTanggal teks DD/MM/YYYY", () => assert.deepEqual(P.parseTanggal("07/07/2025"), { d: 7, m: 7, y: 2025 }));
t("parseTanggal D-M-YY", () => assert.deepEqual(P.parseTanggal("7-8-25"), { d: 7, m: 8, y: 2025 }));
t("parseTanggal ISO", () => assert.deepEqual(P.parseTanggal("2025-08-31"), { d: 31, m: 8, y: 2025 }));
t("parseTanggal nama bulan", () => assert.deepEqual(P.parseTanggal("1 Agustus 2025"), { d: 1, m: 8, y: 2025 }));
t("parseTanggal serial Excel", () => assert.deepEqual(P.parseTanggal(45845), { d: 7, m: 7, y: 2025 }));
t("parseTanggal tidak valid", () => { assert.equal(P.parseTanggal("31/02/2025"), null); assert.equal(P.parseTanggal("besok"), null); });
t("parseJam teks", () => { assert.equal(P.parseJam("08:00"), 480); assert.equal(P.parseJam("8.30"), 510); assert.equal(P.parseJam("13:05:00"), 785); });
t("parseJam desimal Excel", () => { assert.equal(P.parseJam(0.5), 720); assert.equal(P.parseJam(45845.75), 1080); });
t("parseJam tidak valid", () => { assert.equal(P.parseJam("25:00"), null); assert.equal(P.parseJam("pagi"), null); });
t("parseJenis", () => {
  assert.equal(P.parseJenis("Berita Acara Kegiatan").item.id, "3");
  assert.equal(P.parseJenis("Berita Acara Pembimbingan (Konsultasi/Mentoring/Coaching)").item.id, "1");
  assert.equal(P.parseJenis("berita acara ujian").item.id, "2");
  assert.equal(P.parseJenis("").defaulted, true);
  assert.equal(P.parseJenis("rapat").item, null);
});
t("parseTipe", () => {
  assert.equal(P.parseTipe("Offline").value, "true");
  assert.equal(P.parseTipe("online").value, "false");
  assert.equal(P.parseTipe("Hybrid").value, "");
  assert.equal(P.parseTipe("Daring").value, "false");
  assert.equal(P.parseTipe("x"), null);
});
t("parsePeriode lintas bulan", () => {
  const p = P.parsePeriode("7 Jul - 31 Agu 2025");
  assert.equal(p.mulai, Date.UTC(2025, 6, 7)); assert.equal(p.selesai, Date.UTC(2025, 7, 31));
});
t("parsePeriode satu bulan", () => {
  const p = P.parsePeriode("1 - 21 Nov 2025");
  assert.equal(p.mulai, Date.UTC(2025, 10, 1)); assert.equal(p.selesai, Date.UTC(2025, 10, 21));
});
t("parsePeriode lintas tahun", () => {
  const p = P.parsePeriode("20 Des - 5 Jan 2026");
  assert.equal(p.mulai, Date.UTC(2025, 11, 20));
});

console.log("Template resmi");
const template = path.join(ROOT, "website", "assets", "template-logbook-autokonvert.xlsx");
t("template kosong terbaca tanpa error dan 0 baris", () => {
  const { rows, sheet } = P.bacaExcel(XLSX, fs.readFileSync(template));
  assert.equal(sheet, "Logbook");
  assert.equal(rows.length, 0);
});

// isi template dengan variasi format seperti yang mungkin diketik pengguna
const wb = new ExcelJS.Workbook();
await wb.xlsx.readFile(template);
const ws = wb.getWorksheet("Logbook");
ws.getRow(2).values = [1, "07/07/2025", "07:00", "12:00", "Berita Acara Kegiatan", "Offline", "GWW IPB", "Pembukaan MPKMB", "hari1.jpg"];
const r3 = ws.getRow(3);
r3.values = [2, null, null, null, "", "Online", "Zoom", "Materi softskill", "hari2"];
r3.getCell(2).value = new Date(Date.UTC(2025, 6, 8)); r3.getCell(2).numFmt = "dd/mm/yyyy";
r3.getCell(3).value = 8 / 24; r3.getCell(3).numFmt = "hh:mm";
r3.getCell(4).value = 11.5 / 24; r3.getCell(4).numFmt = "hh:mm";
ws.getRow(4).values = [3, "salah", "09:00", "08:00", "rapat", "", "", "", ""];
const buf = Buffer.from(await wb.xlsx.writeBuffer());

t("baris terisi terbaca (kosong dilewati)", () => {
  const { rows } = P.bacaExcel(XLSX, buf);
  assert.equal(rows.length, 3);
  assert.deepEqual(rows.map((r) => r.baris), [2, 3, 4]);
});
t("tanggal/jam teks & tanggal/jam asli Excel sama-sama benar", () => {
  const { rows } = P.bacaExcel(XLSX, buf);
  assert.equal(P.fmtTanggal(P.parseTanggal(rows[0].raw.tanggal)), "07/07/2025");
  assert.equal(P.fmtJam(P.parseJam(rows[0].raw.mulai)), "07:00");
  assert.equal(P.fmtTanggal(P.parseTanggal(rows[1].raw.tanggal)), "08/07/2025");
  assert.equal(P.fmtJam(P.parseJam(rows[1].raw.mulai)), "08:00");
  assert.equal(P.fmtJam(P.parseJam(rows[1].raw.selesai)), "11:30");
  assert.equal(rows[1].raw.bukti, "hari2");
});
t("baris salah tetap terbaca untuk ditandai error", () => {
  const { rows } = P.bacaExcel(XLSX, buf);
  assert.equal(P.parseTanggal(rows[2].raw.tanggal), null);
  assert.equal(P.parseJenis(rows[2].raw.jenis).item, null);
});
t("Excel tanpa judul kolom ditolak dengan pesan jelas", () => {
  const w = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(w, XLSX.utils.aoa_to_sheet([["a", "b"], [1, 2]]), "Sheet1");
  const b = XLSX.write(w, { type: "buffer", bookType: "xlsx" });
  assert.throws(() => P.bacaExcel(XLSX, b), /judul kolom/);
});
t("judul kolom alternatif (CSV) tetap dikenali", () => {
  const csv = "Tgl,Mulai,Selesai,Media,Tempat,Deskripsi,Foto\n10/07/2025,08:00,10:00,Hybrid,Kampus,Kegiatan X,a.png\n";
  const { rows } = P.bacaExcel(XLSX, Buffer.from(csv));
  assert.equal(rows.length, 1);
  assert.equal(rows[0].raw.keterangan, "Kegiatan X");
  assert.equal(rows[0].raw.bukti, "a.png");
});

console.log("Format bebas (spreadsheet buatan sendiri)");
t("parseJam AM/PM & tanda baca nyasar", () => {
  assert.equal(P.parseJam("8:00:00 PM,"), 1200);
  assert.equal(P.parseJam("12:30 AM"), 30);
  assert.equal(P.parseJam("7 pm"), 1140);
  assert.equal(P.parseJam("."), null);
  assert.equal(P.parseJam("8"), null);
});
t("tipeLokasi kolom gabungan", () => {
  assert.deepEqual(P.tipeLokasi("", "Online, Zoom Meeting"), { tipe: P.TIPE.online, lokasi: "Zoom Meeting" });
  assert.deepEqual(P.tipeLokasi("", "Hybrid, IPB University"), { tipe: P.TIPE.hybrid, lokasi: "IPB University" });
  assert.deepEqual(P.tipeLokasi("", "Offline - IPB Dramaga"), { tipe: P.TIPE.offline, lokasi: "IPB Dramaga" });
  assert.deepEqual(P.tipeLokasi("Offline", "Graha Widya Wisuda"), { tipe: P.TIPE.offline, lokasi: "Graha Widya Wisuda" });
  assert.equal(P.tipeLokasi("", "Kampus Dramaga").tipe, null);
});
t("judul kolom: kata utuh", () => {
  assert.equal(P.cocokKolom("Deskripsi Kegiatan"), "keterangan");
  assert.equal(P.cocokKolom("Tempat Pelaksanaan "), "lokasi");
  assert.equal(P.cocokKolom("Dokumentasi"), "bukti");
  assert.equal(P.cocokKolom("Jenis Kegiatan "), "jenis");
  assert.equal(P.cocokKolom("Durasi"), null);
  assert.equal(P.cocokKolom("Nama Pendamping"), null);
});

{
  const H = ["Tanggal", "Jam Mulai", "Jam Selesai", "Durasi", "Jenis Kegiatan ", "Deskripsi Kegiatan", "Tempat Pelaksanaan ", "Dokumentasi"];
  const s1 = XLSX.utils.aoa_to_sheet([
    ["OMB\nKELOMPOK X", "", "", "", "", "", "", ""],
    H,
    ["19/07/2025", 0.5, 2 / 3, 1 / 6, "Berita Acara Ujian ", "Talent mapping", "Online, Website Talent Mapping", ""],
    [45845, 13 / 24, 2 / 3, 0.125, "Berita Acara Kegiatan", "First meet", "Online, Zoom Meeting", "[07-08] INTERNALISASI 1"],
    [45870, 13 / 24, "8:00:00 PM,", "#VALUE!", "Berita Acara Kegiatan", "Penugasan", "Offline, IPB Dramaga", "D.1. HARMONI 1 (2)"],
    ["TOTAL", "", "", "#VALUE!", "", "", "", ""],
  ]);
  XLSX.utils.sheet_add_aoa(s1, [["", "."]], { origin: "A12" }); // titik nyasar di baris 12
  s1["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 7 } }, { s: { r: 5, c: 0 }, e: { r: 5, c: 2 } }];
  const s2 = XLSX.utils.aoa_to_sheet([
    ["7 HABITS", "", "", "", "", "", "", ""],
    H,
    ["14/07/2025", 0.375, 13 / 24, "", "Berita Acara Ujian", "Refleksi diri", "Hybrid, IPB University", "Dokumentasi masing-masing"],
    ["", 0.375, 13 / 24, "", "Berita Acara Ujian", "Challenge 2", "Online, IPB Class", ""],
    ["", 0.375, 13 / 24, "", "Berita Acara Pembimbingan (Konsultasi/Mentoring/Coaching)", "Challenge 3", "Online, IPB Class", ""],
    ["TOTAL LEARNING HOURS", "", "", 1.875, "", "", "", ""],
  ]);
  s2["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 7 } }, { s: { r: 2, c: 0 }, e: { r: 4, c: 0 } }];
  const w = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(w, s1, "omb");
  XLSX.utils.book_append_sheet(w, s2, "7 habits");
  const b = XLSX.write(w, { type: "buffer", bookType: "xlsx" });

  t("dua sheet logbook terdeteksi, default sheet pertama", () => {
    const res = P.bacaExcel(XLSX, b);
    assert.deepEqual(res.sheets, ["omb", "7 habits"]);
    assert.equal(res.sheet, "omb");
  });
  t("judul merge, baris TOTAL & titik nyasar dilewati; nomor baris sesuai Excel", () => {
    const { rows } = P.bacaExcel(XLSX, b, "omb");
    assert.deepEqual(rows.map((r) => r.baris), [3, 4, 5]);
    assert.equal(P.fmtJam(P.parseJam(rows[2].raw.selesai)), "20:00");
    assert.equal(rows[1].raw.bukti, "[07-08] INTERNALISASI 1");
    assert.equal(rows[2].raw.bukti, "D.1. HARMONI 1 (2)");
    assert.equal(P.tipeLokasi(rows[0].raw.tipe, rows[0].raw.lokasi).lokasi, "Website Talent Mapping");
  });
  t("tanggal yang di-merge untuk beberapa baris ikut terisi", () => {
    const { rows } = P.bacaExcel(XLSX, b, "7 habits");
    assert.equal(rows.length, 3);
    assert.deepEqual(rows.map((r) => P.fmtTanggal(P.parseTanggal(r.raw.tanggal))), ["14/07/2025", "14/07/2025", "14/07/2025"]);
    assert.equal(P.parseJenis(rows[2].raw.jenis).item.id, "1");
  });
  t("sheet Contoh di template tidak dihitung sebagai logbook", () => {
    assert.deepEqual(P.bacaExcel(XLSX, fs.readFileSync(template)).sheets, ["Logbook"]);
  });
}

console.log("Google Drive");
t("parseDriveLink", () => {
  assert.deepEqual(P.parseDriveLink("https://drive.google.com/drive/folders/1ContohFolderDrive_AbCdEfGhIjK?usp=share_link"), { kind: "folder", id: "1ContohFolderDrive_AbCdEfGhIjK" });
  assert.deepEqual(P.parseDriveLink("https://drive.google.com/drive/u/1/folders/1AbcdEfghIjk_-"), { kind: "folder", id: "1AbcdEfghIjk_-" });
  assert.deepEqual(P.parseDriveLink("https://drive.google.com/file/d/1ZKeSpqXwGuBNMwd5X/view?usp=sharing"), { kind: "file", id: "1ZKeSpqXwGuBNMwd5X" });
  assert.deepEqual(P.parseDriveLink("https://drive.google.com/open?id=1ZKeSpqXwGuBNMwd5X"), { kind: "file", id: "1ZKeSpqXwGuBNMwd5X" });
  assert.equal(P.parseDriveLink("https://evil.example.com/drive/folders/1ContohFolderDrive"), null);
  assert.equal(P.parseDriveLink(""), null);
});
t("parseDriveFolder: nama, jenis, urutan", () => {
  const entri = (id, href, tipe, nama) =>
    `<div class="flip-entry" id="entry-${id}" tabindex="0" role="link"><div class="flip-entry-info"><a href="${href}" target="_blank"><div class="flip-entry-list-icon"><img src="https://drive-thirdparty.googleusercontent.com/16/type/${tipe}" alt=""/></div><div class="flip-entry-title">${nama}</div></a></div></div>`;
  const html =
    entri("1bbbbbbbbbbbb", "https://drive.google.com/file/d/1bbbbbbbbbbbb/view", "image/png", "foto 10.png") +
    entri("1aaaaaaaaaaaa", "https://drive.google.com/file/d/1aaaaaaaaaaaa/view", "image/jpeg", "foto 2\\.jpg") +
    entri("1cccccccccccc", "https://drive.google.com/drive/folders/1cccccccccccc", "application/vnd.google-apps.folder", "Sub") +
    entri("1dddddddddddd", "https://drive.google.com/file/d/1dddddddddddd/view", "application/pdf", "Surat &amp; LoA.pdf");
  const f = P.parseDriveFolder(html);
  assert.deepEqual(f.map((x) => [x.name, x.jenis]), [["foto 2.jpg", "gambar"], ["foto 10.png", "gambar"], ["Sub", "folder"], ["Surat & LoA.pdf", "pdf"]]);
});
t("hyperlink di sel kolom bukti terbaca", () => {
  const ws = XLSX.utils.aoa_to_sheet([
    ["Tanggal", "Jam Mulai", "Jam Selesai", "Deskripsi Kegiatan", "Tempat Pelaksanaan", "Dokumentasi"],
    ["01/08/2025", "08:00", "10:00", "A", "Offline, IPB", "DAY 1"],
    ["02/08/2025", "08:00", "10:00", "B", "Offline, IPB", "https://drive.google.com/file/d/1ZKeSpqXwGuBNMwd5X/view"],
    ["03/08/2025", "08:00", "10:00", "C", "Offline, IPB", "catatan saja"],
  ]);
  ws.F2.l = { Target: "https://drive.google.com/drive/folders/1ContohFolderDrive_AbCdEfGhIjK" };
  const w = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(w, ws, "x");
  const { rows } = P.bacaExcel(XLSX, XLSX.write(w, { type: "buffer", bookType: "xlsx" }));
  assert.equal(rows[0].raw.bukti, "DAY 1");
  assert.equal(P.parseDriveLink(rows[0].raw.buktiLink).kind, "folder");
  assert.equal(P.parseDriveLink(rows[1].raw.buktiLink).kind, "file");
  assert.equal(rows[2].raw.buktiLink, "");
});

console.log(`\n${lulus} tes lulus`);
