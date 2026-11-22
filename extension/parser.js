/*
 * AutoConvert Logbook IPB — logika parsing murni (tanpa DOM).
 * Dipakai oleh content.js dan diuji oleh tools/test-parser.mjs.
 */
(() => {
  "use strict";
  const JENIS_KEGIATAN = [
    { id: "3", label: "Berita Acara Kegiatan", kata: ["kegiatan"] },
    { id: "1", label: "Berita Acara Pembimbingan", kata: ["pembimbingan", "konsultasi", "mentoring", "coaching", "bimbingan"] },
    { id: "2", label: "Berita Acara Ujian", kata: ["ujian"] },
  ];
  // Nilai radio IsLuring di form portal: Hybrid = "", Offline = "true", Online = "false"
  const TIPE = {
    offline: { value: "true", label: "Offline" },
    online: { value: "false", label: "Online" },
    hybrid: { value: "", label: "Hybrid" },
  };

  const BULAN = {
    jan: 1, januari: 1, january: 1,
    feb: 2, februari: 2, february: 2, peb: 2,
    mar: 3, maret: 3, march: 3,
    apr: 4, april: 4,
    mei: 5, may: 5,
    jun: 6, juni: 6, june: 6,
    jul: 7, juli: 7, july: 7,
    agu: 8, agt: 8, ags: 8, agustus: 8, aug: 8, august: 8,
    sep: 9, sept: 9, september: 9,
    okt: 10, oktober: 10, oct: 10, october: 10,
    nov: 11, nopember: 11, november: 11,
    des: 12, desember: 12, dec: 12, december: 12,
  };

  // ------------------------------------------------------------ util kecil
  const pad2 = (n) => String(n).padStart(2, "0");
  const norm = (s) => String(s ?? "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();

  function isValidDate(d, m, y) {
    if (!(y >= 2000 && y <= 2100 && m >= 1 && m <= 12 && d >= 1)) return false;
    return d <= new Date(Date.UTC(y, m, 0)).getUTCDate();
  }

  /** Ubah nilai sel Excel/teks menjadi {d,m,y} atau null. */
  function parseTanggal(v) {
    if (v === null || v === undefined || v === "") return null;
    if (typeof v === "number") {
      // serial date Excel (hari sejak 1899-12-30)
      const ms = Math.round((Math.floor(v) - 25569) * 86400000);
      const dt = new Date(ms);
      const r = { d: dt.getUTCDate(), m: dt.getUTCMonth() + 1, y: dt.getUTCFullYear() };
      return isValidDate(r.d, r.m, r.y) ? r : null;
    }
    const s = String(v).trim().toLowerCase();
    let m;
    if ((m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/))) {
      const r = { y: +m[1], m: +m[2], d: +m[3] };
      return isValidDate(r.d, r.m, r.y) ? r : null;
    }
    if ((m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/))) {
      let y = +m[3];
      if (y < 100) y += 2000;
      const r = { d: +m[1], m: +m[2], y };
      return isValidDate(r.d, r.m, r.y) ? r : null;
    }
    if ((m = s.match(/^(\d{1,2})\s+([a-z]+)\.?\s+(\d{4})$/))) {
      const bln = BULAN[m[2]];
      const r = { d: +m[1], m: bln, y: +m[3] };
      return bln && isValidDate(r.d, r.m, r.y) ? r : null;
    }
    return null;
  }

  /** Ubah nilai sel Excel/teks menjadi menit sejak 00:00, atau null. */
  function parseJam(v) {
    if (v === null || v === undefined || v === "") return null;
    if (typeof v === "number") {
      const frac = v - Math.floor(v);
      const menit = Math.round(frac * 1440) % 1440;
      return menit;
    }
    // rapikan: "8:00:00 PM," / "08.30" / "8.00 p.m."
    const s = String(v)
      .trim()
      .toLowerCase()
      .replace(/(\d)\.(\d)/g, "$1:$2")
      .replace(/[^0-9:apm ]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
    const m = s.match(/^(\d{1,2})(?:\s*:\s*(\d{2}))?(?:\s*:\s*\d{2})?\s*(am|pm)?$/) || s.match(/^(\d{2})(\d{2})$/);
    if (!m) return null;
    let h = +m[1];
    const mi = m[2] === undefined ? 0 : +m[2];
    const ampm = m[3];
    if (m[2] === undefined && !ampm) return null; // "8" saja ambigu
    if (ampm) {
      if (h < 1 || h > 12) return null;
      if (ampm === "pm" && h !== 12) h += 12;
      if (ampm === "am" && h === 12) h = 0;
    }
    if (h > 23 || mi > 59) return null;
    return h * 60 + mi;
  }

  const fmtTanggal = (t) => `${pad2(t.d)}/${pad2(t.m)}/${t.y}`;
  const fmtJam = (menit) => `${pad2(Math.floor(menit / 60))}:${pad2(menit % 60)}`;
  const keyLog = (t, mulai) => `${fmtTanggal(t)} ${fmtJam(mulai)}`;

  function parseJenis(v) {
    const s = norm(v);
    if (!s) return { item: JENIS_KEGIATAN[0], defaulted: true };
    if (/^[123]$/.test(s)) return { item: JENIS_KEGIATAN.find((j) => j.id === s) };
    // cek "pembimbingan"/"ujian" dulu karena semua label mengandung "berita acara"
    const urutan = [JENIS_KEGIATAN[1], JENIS_KEGIATAN[2], JENIS_KEGIATAN[0]];
    for (const j of urutan) if (j.kata.some((k) => s.includes(k))) return { item: j };
    return { item: null };
  }

  function parseTipe(v) {
    const s = norm(v);
    if (!s) return null;
    if (s.includes("hybrid") || s.includes("hibrid") || s.includes("campur")) return TIPE.hybrid;
    if (s.includes("offline") || s.includes("luring") || s.includes("tatap")) return TIPE.offline;
    if (s.includes("online") || s.includes("daring")) return TIPE.online;
    return null;
  }

  /**
   * Tentukan tipe & lokasi. Mendukung kolom terpisah, atau satu kolom gabungan
   * seperti di tabel portal: "Online, Zoom Meeting" / "Offline - IPB Dramaga".
   */
  function tipeLokasi(rawTipe, rawLokasi) {
    let lokasi = String(rawLokasi ?? "").trim();
    let tipe = parseTipe(rawTipe);
    const m = lokasi.match(/^(online|offline|hybrid|hibrid|daring|luring)\b\s*[,:;\-–—]?\s*(.*)$/i);
    if (m) {
      if (!tipe) tipe = parseTipe(m[1]);
      lokasi = m[2].trim() || m[1].trim();
    }
    return { tipe, lokasi };
  }

  /** "7 Jul - 31 Agu 2025" / "1 - 21 Nov 2025" -> {mulai, selesai} (Date UTC) */
  function parsePeriode(s) {
    const m = String(s).toLowerCase().match(/^(\d{1,2})\s*([a-z]+)?\s*(\d{4})?\s*-\s*(\d{1,2})\s*([a-z]+)\s*(\d{4})$/);
    if (!m) return null;
    const yEnd = +m[6];
    const mEnd = BULAN[m[5]];
    const mStart = m[2] ? BULAN[m[2]] : mEnd;
    const yStart = m[3] ? +m[3] : mStart > mEnd ? yEnd - 1 : yEnd;
    if (!mEnd || !mStart) return null;
    return { mulai: Date.UTC(yStart, mStart - 1, +m[1]), selesai: Date.UTC(yEnd, mEnd - 1, +m[4]) };
  }

  // ------------------------------------------------------------ baca Excel
  const KOLOM = {
    tanggal: ["tanggal", "tgl", "date"],
    mulai: ["jam mulai", "waktu mulai", "mulai", "start"],
    selesai: ["jam selesai", "waktu selesai", "selesai", "end"],
    jenis: ["jenis kegiatan", "jenis"],
    tipe: ["tipe penyelenggaraan", "tipe", "media", "penyelenggaraan"],
    lokasi: ["lokasi", "tempat"],
    keterangan: ["keterangan", "deskripsi", "kegiatan", "uraian"],
    bukti: ["nama file bukti", "file bukti", "bukti", "foto", "nama file", "dokumen", "dokumentasi", "lampiran", "bukti aktivitas"],
  };

  function cocokKolom(header) {
    const h = norm(header);
    if (!h) return null;
    // cocok persis dulu, lalu "mengandung"
    for (const [k, alias] of Object.entries(KOLOM)) if (alias.includes(h)) return k;
    // per kata utuh, supaya mis. "pendamping" tidak dianggap "end"
    const kata = ` ${h} `;
    for (const [k, alias] of Object.entries(KOLOM)) if (alias.some((a) => kata.includes(` ${a} `))) return k;
    return null;
  }

  /** Cari baris judul kolom (yang punya Tanggal & Keterangan) dalam array-of-arrays. */
  function cariHeader(aoa) {
    for (let i = 0; i < Math.min(aoa.length, 20); i++) {
      const p = {};
      (aoa[i] || []).forEach((c, j) => {
        const k = cocokKolom(c);
        if (k && p[k] === undefined) p[k] = j;
      });
      if (p.tanggal !== undefined && p.keterangan !== undefined) return { hIdx: i, peta: p };
    }
    return null;
  }

  function sheetKeAoa(XLSX, ws) {
    if (!ws || !ws["!ref"]) return { aoa: [], r0: 0 };
    const range = XLSX.utils.decode_range(ws["!ref"]);
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "", blankrows: true });
    // sel hasil merge (mis. satu tanggal untuk beberapa baris) diisi dengan nilai sel kiri-atasnya
    for (const mg of ws["!merges"] || []) {
      const v = aoa[mg.s.r - range.s.r]?.[mg.s.c - range.s.c];
      for (let r = mg.s.r; r <= mg.e.r; r++) {
        const row = aoa[r - range.s.r];
        if (!row) continue;
        for (let c = mg.s.c; c <= mg.e.c; c++) {
          const j = c - range.s.c;
          if (row[j] === "" || row[j] === undefined) row[j] = v;
        }
      }
    }
    return { aoa, r0: range.s.r, c0: range.s.c };
  }

  /** URL dari sebuah sel: hyperlink sel, rumus =HYPERLINK("url";...), atau teks yang berupa URL. */
  function linkSel(XLSX, ws, r, c) {
    const merge = (ws["!merges"] || []).find((m) => r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c);
    const addr = XLSX.utils.encode_cell(merge ? merge.s : { r, c });
    const cell = ws[addr];
    if (!cell) return "";
    if (cell.l && cell.l.Target && /^https?:\/\//i.test(cell.l.Target)) return cell.l.Target;
    const f = cell.f && String(cell.f).match(/HYPERLINK\(\s*"([^"]+)"/i);
    if (f) return f[1];
    const v = String(cell.v ?? "").trim();
    return /^https?:\/\/\S+$/i.test(v) ? v : "";
  }

  /**
   * Kenali link Google Drive.
   * folder: drive.google.com/drive/folders/ID, .../drive/u/0/folders/ID, open?id=ID (dianggap file)
   * file:   drive.google.com/file/d/ID/..., open?id=ID, uc?id=ID, docs.google.com/uc?id=ID
   */
  function parseDriveLink(url) {
    const s = String(url || "").trim();
    if (!/^https?:\/\/(drive|docs)\.google\.com\//i.test(s)) return null;
    let m;
    if ((m = s.match(/\/folders\/([\w-]{10,})/))) return { kind: "folder", id: m[1] };
    if ((m = s.match(/\/file\/d\/([\w-]{10,})/))) return { kind: "file", id: m[1] };
    if ((m = s.match(/[?&]id=([\w-]{10,})/))) return { kind: "file", id: m[1] };
    return null;
  }

  const decodeHtml = (s) =>
    String(s)
      .replace(/&#(\d+);/g, (_, d) => String.fromCharCode(+d))
      .replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCharCode(parseInt(h, 16)))
      .replace(/&quot;/g, '"')
      .replace(/&#39;|&apos;/g, "'")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&amp;/g, "&");

  /**
   * Baca daftar isi folder dari HTML drive.google.com/embeddedfolderview?id=...
   * -> [{ id, name, jenis: "gambar"|"pdf"|"folder"|"lain" }]  (urut nama)
   */
  function parseDriveFolder(html) {
    const out = [];
    const re = /<div class="flip-entry" id="entry-([\w-]+)"([\s\S]*?)<div class="flip-entry-title">([^<]*)<\/div>/g;
    let m;
    while ((m = re.exec(String(html)))) {
      const blok = m[2];
      const name = decodeHtml(m[3]).replace(/\\\./g, ".").trim();
      const href = (blok.match(/<a href="([^"]+)"/) || [])[1] || "";
      const tipe = ((blok.match(/\/type\/([^"]+)"/) || [])[1] || "").toLowerCase();
      let jenis = "lain";
      if (/\/folders\//.test(href) || tipe.includes("folder")) jenis = "folder";
      else if (tipe.startsWith("image/") || /\.(jpe?g|png|webp|heic|gif)$/i.test(name)) jenis = "gambar";
      else if (tipe.includes("pdf") || /\.pdf$/i.test(name)) jenis = "pdf";
      out.push({ id: m[1], name, jenis });
    }
    return out.sort((a, b) => a.name.localeCompare(b.name, "id", { numeric: true }));
  }

  /**
   * Baca workbook: baris logbook dari satu sheet + daftar sheet yang berisi logbook.
   * @param namaSheet opsional; default sheet "Logbook", atau sheet logbook pertama.
   */
  function bacaExcel(XLSX, buffer, namaSheet) {
    const wb = XLSX.read(buffer, { type: "array", cellDates: false });
    const cache = {};
    const valid = [];
    for (const n of wb.SheetNames) {
      const { aoa, r0, c0 } = sheetKeAoa(XLSX, wb.Sheets[n]);
      const h = cariHeader(aoa);
      cache[n] = { aoa, r0, c0, h, ws: wb.Sheets[n] };
      if (h && norm(n) !== "contoh") valid.push(n);
    }
    if (!valid.length) {
      throw new Error(
        'Baris judul kolom tidak ditemukan. Pastikan ada kolom "Tanggal" dan "Keterangan" atau "Deskripsi Kegiatan" (lihat template).'
      );
    }
    const pilih =
      (namaSheet && valid.includes(namaSheet) && namaSheet) || valid.find((n) => norm(n) === "logbook") || valid[0];
    const { aoa, r0, c0, h, ws } = cache[pilih];
    const { hIdx, peta } = h;

    const wajib = ["tanggal", "mulai", "selesai", "lokasi", "keterangan"];
    const kurang = wajib.filter((k) => peta[k] === undefined);
    if (kurang.length) {
      throw new Error(`Sheet "${pilih}": kolom berikut tidak ada: ` + kurang.map((k) => KOLOM[k][0]).join(", "));
    }

    const rows = [];
    for (let i = hIdx + 1; i < aoa.length; i++) {
      const r = aoa[i] || [];
      const ambil = (k) => (peta[k] === undefined ? "" : r[peta[k]] ?? "");
      // baris rekap seperti "TOTAL" / "TOTAL LEARNING HOURS" dilewati
      if (r.some((c) => /^\s*(sub\s*)?total\b/i.test(String(c)))) continue;
      // baris tanpa tanggal, lokasi, dan keterangan dianggap kosong (mis. hanya ada titik nyasar di kolom jam)
      const kosong = ["tanggal", "lokasi", "keterangan"].every((k) => String(ambil(k)).trim() === "");
      if (kosong) continue;
      rows.push({
        baris: r0 + i + 1,
        raw: {
          tanggal: ambil("tanggal"),
          mulai: ambil("mulai"),
          selesai: ambil("selesai"),
          jenis: ambil("jenis"),
          tipe: ambil("tipe"),
          lokasi: String(ambil("lokasi")).trim(),
          keterangan: String(ambil("keterangan")).trim(),
          bukti: String(ambil("bukti")).trim(),
          buktiLink: peta.bukti === undefined ? "" : linkSel(XLSX, ws, r0 + i, c0 + peta.bukti),
        },
      });
    }
    return { rows, sheet: pilih, sheets: valid };
  }

  globalThis.AutoConvertParser = {
    JENIS_KEGIATAN, TIPE, BULAN, KOLOM,
    pad2, norm, isValidDate, parseTanggal, parseJam, fmtTanggal, fmtJam, keyLog,
    parseJenis, parseTipe, tipeLokasi, parsePeriode, cocokKolom, cariHeader, bacaExcel, parseDriveLink, parseDriveFolder,
  };
})();
