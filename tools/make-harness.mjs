// Membuat halaman uji tiruan Log Aktivitas (tanpa portal asli) untuk melihat panel.
// Hasil: tools/harness/index.html. Drive & portal dipalsukan, tidak ada data dikirim ke mana pun.
import fs from "node:fs";
import * as XLSX from "xlsx";

const H = ["Tanggal", "Jam Mulai", "Jam Selesai", "Durasi", "Jenis Kegiatan ", "Deskripsi Kegiatan", "Tempat Pelaksanaan ", "Dokumentasi"];
const ws = XLSX.utils.aoa_to_sheet([
  ["OMB\nKELOMPOK CONTOH"], H,
  ["19/07/2025", 0.5, 2 / 3, "", "Berita Acara Ujian", "Talent Mapping: pengisian asesmen awal", "Online, Website Talent Mapping", ""],
  [45845, 13 / 24, 2 / 3, "", "Berita Acara Kegiatan", "First meet online kelompok besar", "Online, Zoom Meeting", "[07-08] INTERNALISASI 1"],
  [45846, 19 / 24, 21 / 24, "", "Berita Acara Kegiatan", "Internalisasi 1: perkenalan kelompok sedang", "Online, Zoom Meeting", "[07-09] MUBES"],
  [45870, 13 / 24, "8:00:00 PM,", "", "Berita Acara Kegiatan", "Pembuatan penugasan kelompok", "Offline, IPB Dramaga", "Lampirkan hasil tugasnya"],
  [45847, 19 / 24, 21 / 24, "", "Berita Acara Kegiatan", "Musyawarah besar ketua kelompok", "Online, Zoom Meeting", "[07-09] MUBES BESAR"],
  ["TOTAL", "", "", "", "", "", "", ""],
]);
ws.H4.l = { Target: "https://drive.google.com/drive/folders/1PUBLIKpubliKPUBLIK" };
ws.H5.l = { Target: "https://drive.google.com/drive/folders/1PRIVATprivatPRIVAT" };
ws.H7.l = { Target: "https://drive.google.com/drive/folders/1PUBLIKpubliKPUBLIK" };
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, "omb");
const xlsxB64 = XLSX.write(wb, { type: "base64", bookType: "xlsx" });

const entri = (id, nama) => `<div class="flip-entry" id="entry-${id}"><a href="https://drive.google.com/file/d/${id}/view"><div class="flip-entry-list-icon"><img src="x/type/image/png" alt=""/></div><div class="flip-entry-title">${nama}</div></a></div>`;
const folderHtml = Array.from({ length: 6 }, (_, i) => entri(`1FOTOfotoFOTO${i}xx`, `foto ${i + 1}.png`)).join("");
// PNG 1x1 biru
const png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk4P9fDwAChwGA60e6kgAAAABJRU5ErkJggg==";

const modal = `<form id="main_form" action="/Kegiatan/LogAktivitasKampusMerdeka/Tambah" method="post">
<input name="Id" type="hidden" value=""><input name="AktivitasId" type="hidden" value="abc">
<input name="Waktu"><input name="Tmw"><input name="Tsw">
<select name="JenisLogbookKegiatanKampusMerdekaId"><option value="">--</option><option value="1">1</option><option value="2">2</option><option value="3">3</option></select>
<input name="ListDosenPembimbing[0].Key.PembimbingId" type="hidden" value="389192">
<div class="form-check"><input type="checkbox" id="ListDosenPembimbing_0__Value" name="ListDosenPembimbing[0].Value" value="true"><label for="ListDosenPembimbing_0__Value">Dr. Contoh Dosen, S.Kom., M.Kom. - 1980</label></div>
<input type="radio" name="IsLuring" value=""><input type="radio" name="IsLuring" value="true"><input type="radio" name="IsLuring" value="false">
<input name="Lokasi"><textarea name="Keterangan"></textarea><input type="file" name="File">
<input name="__RequestVerificationToken" type="hidden" value="tok"></form>`;

const shim = `
if (!location.hash.includes("bantuan")) { try { history.replaceState(null, "", "/Kegiatan/LogAktivitasKampusMerdeka/Index/abc"); } catch (e) { console.log("replaceState gagal", e.message); } }
window.__modal = ${JSON.stringify(modal)};
const _fetch = window.fetch;
window.fetch = async (url, opt) => {
  url = String(url);
  if (url.includes("LogAktivitasKampusMerdeka/Tambah?")) return new Response(window.__modal, { status: 200, headers: { "content-type": "text/html" } });
  if (url.includes("LogAktivitasKampusMerdeka/Tambah")) { const f = opt.body.get("File"); window.__posts.push({ Waktu: opt.body.get("Waktu"), Tmw: opt.body.get("Tmw"), Tsw: opt.body.get("Tsw"), IsLuring: opt.body.get("IsLuring"), Lokasi: opt.body.get("Lokasi"), Jenis: opt.body.get("JenisLogbookKegiatanKampusMerdekaId"), Dosen: opt.body.getAll("ListDosenPembimbing[0].Value").join(","), token: opt.body.get("__RequestVerificationToken"), File: f && f.name ? f.name + " (" + f.type + ", " + f.size + "B)" : "(kosong)" }); await new Promise(r => setTimeout(r, 400)); return new Response(JSON.stringify({ status: true, message: "Data berhasil disimpan (UJI)", url: null }), { status: 200 }); }
  return _fetch(url, opt);
};
const PNG = Uint8Array.from(atob(${JSON.stringify(png)}), c => c.charCodeAt(0)).buffer;
window.GM = { xmlHttpRequest(o) { setTimeout(() => {
  if (o.url.includes("embeddedfolderview")) {
    if (o.url.includes("PRIVAT")) return o.onload({ status: 401, finalUrl: o.url, responseHeaders: "content-type: text/html", responseText: "no", response: "no" });
    return o.onload({ status: 200, finalUrl: o.url, responseHeaders: "content-type: text/html", responseText: ${JSON.stringify(folderHtml)}, response: "" });
  }
  o.onload({ status: 200, finalUrl: o.url, responseHeaders: "content-type: image/png", response: PNG });
}, 300); } };
window.confirm = () => true; window.alert = (m) => console.log("ALERT", m);
window.__posts = [];
window.__ujiExcel = () => {
  const b = Uint8Array.from(atob(${JSON.stringify(xlsxB64)}), c => c.charCodeAt(0));
  const f = new File([b], "logbook-contoh.xlsx", { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const inp = document.getElementById("autokonvert-root").shadowRoot.getElementById("excel");
  const dt = new DataTransfer(); dt.items.add(f); inp.files = dt.files; inp.dispatchEvent(new Event("change"));
};`;

const us = fs.readFileSync("website/assets/autokonvert.user.js", "utf8");
const html = `<!doctype html><html lang="id"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Uji Panel AutoKonvert</title><style>body{font-family:sans-serif;margin:0;padding:16px;background:#f4f6f9}</style></head><body>
<div class="content-wrapper"><h3>Log Aktivitas Kampus Merdeka (TIRUAN UJI)</h3>
<div>Nama Kegiatan
MPKMB IPB University 2025, Peserta</div><div>Waktu
7 Jul - 31 Agu 2025</div>
<a onclick="OpenModal('/Kegiatan/LogAktivitasKampusMerdeka/Tambah?aktivitasId=abc')">Tambah</a>
<table><tbody><tr><td>1</td><td>07 Juli 2025, 13:00 - 16:00</td><td>First meet online kelompok besar</td></tr></tbody></table></div>
<script>${shim}</script><script>${us.replace(/<\/script/gi, "<\/script")}</script></body></html>`;
fs.writeFileSync("tools/harness/index.html", html);
console.log("✓ tools/harness/index.html", (html.length / 1024).toFixed(0) + " KB");
