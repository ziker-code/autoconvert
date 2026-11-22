/*
 * AutoKonvert Logbook IPB — content script.
 *
 * Berjalan di halaman Log Aktivitas Student Portal
 * (/Kegiatan/LogAktivitasKampusMerdeka/Index/{id}) dan menambahkan panel
 * untuk mengunggah Excel + foto bukti, lalu mengirim log satu per satu
 * memakai form resmi Student Portal (sesi login milik pengguna sendiri).
 */
(() => {
  "use strict";
  if (window.__autokonvertLoaded) return;
  window.__autokonvertLoaded = true;

  // Tombol ⚡ muncul di semua halaman Student Portal (kecuali login). Panel pengisian
  // hanya di halaman Log; di halaman lain panel menunjukkan link ke Log aktivitas.
  const PATH = location.pathname.toLowerCase();
  if (PATH.startsWith("/account/")) return;
  const HALAMAN_LOG = PATH.includes("/kegiatan/logaktivitaskampusmerdeka/index/");

  // ---------------------------------------------------------------- konstanta
  const JEDA_ANTAR_KIRIM_MS = 1500;
  const BATAS_KOMPRES_BYTES = 2 * 1024 * 1024; // foto > 2 MB dikompres
  const BATAS_UPLOAD_BYTES = 10 * 1024 * 1024; // batas portal 10 MB
  const PORTAL_ORIGIN = "https://studentportal.ipb.ac.id";
  const EKSTENSI_BUKTI = ["pdf", "png", "jpg", "jpeg"];
  // gambar lain (HEIC dari iPhone, WEBP, …) diterima lalu diubah ke JPEG sebelum dikirim
  const EKSTENSI_GAMBAR_LAIN = ["heic", "heif", "webp", "gif", "bmp"];
  const ACCEPT_BUKTI = "image/*,application/pdf,.pdf,.jpg,.jpeg,.png,.heic";
  const ACCEPT_EXCEL =
    ".xlsx,.xls,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv";

  // Lingkungan tempat panel berjalan: ekstensi Chrome (default) atau userscript iPhone
  // (userscript mendefinisikan globalThis.AutoKonvertHost sebelum file ini).
  const HOST = globalThis.AutoKonvertHost || {
    jenis: "ekstensi",
    versi: chrome.runtime.getManifest().version,
    minta: (pesan) => chrome.runtime.sendMessage(pesan),
  };

  const DI_HP = HOST.jenis === "userscript";
  const LOKAL = DI_HP ? "Dari HP" : "Komputer"; // label tombol pilih file lokal
  const LOKAL_KECIL = DI_HP ? "HP" : "komputer";

  const P = globalThis.AutoKonvertParser;
  const { BULAN, pad2, norm, parseTanggal, parseJam, fmtTanggal, fmtJam, parseJenis, tipeLokasi, parsePeriode } = P;
  const bacaExcel = (buf, sheet) => P.bacaExcel(XLSX, buf, sheet);
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const esc = (s) =>
    String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  // ikon garis 24px (gaya sama dengan website)
  const IKON = {
    zap: '<path d="M13 2 4 14h7l-1 8 9-12h-7z"/>',
    x: '<path d="M18 6 6 18M6 6l12 12"/>',
    check: '<path d="M20 6 9 17l-5-5"/>',
    sheet: '<path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/><path d="M14 3v6h6M8 13h8M8 17h8M12 13v4"/>',
    image: '<rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="9" cy="9" r="1.6"/><path d="m21 15-5-5L5 21"/>',
    file: '<path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/><path d="M14 3v6h6M8 13h8M8 17h5"/>',
    cloud: '<path d="M7 18.5a4.5 4.5 0 0 1-.6-8.96A6 6 0 0 1 18 9.5a4.5 4.5 0 0 1 0 9z"/>',
    upload: '<path d="M12 21V9"/><path d="m7 14 5-5 5 5"/><path d="M5 3h14"/>',
    grid: '<rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/>',
    external: '<path d="M15 3h6v6M10 14 21 3"/><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/>',
    arrow: '<path d="M5 12h14"/><path d="m13 6 6 6-6 6"/>',
    alert: '<path d="M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/><path d="M12 9v4M12 17h.01"/>',
    xcircle: '<circle cx="12" cy="12" r="9"/><path d="m15 9-6 6M9 9l6 6"/>',
    info: '<circle cx="12" cy="12" r="9"/><path d="M12 11v5M12 8h.01"/>',
    loader: '<path d="M21 12a9 9 0 1 1-6.2-8.56"/>',
    send: '<path d="m22 2-7 20-4-9-9-4z"/><path d="M22 2 11 13"/>',
    refresh: '<path d="M21 12a9 9 0 1 1-2.64-6.36L21 8"/><path d="M21 3v5h-5"/>',
    report: '<path d="M12 3v12"/><path d="m7 10 5 5 5-5"/><path d="M5 21h14"/>',
  };
  const ic = (nama, kelas = "") =>
    `<svg class="ic ${kelas}" viewBox="0 0 24 24" aria-hidden="true">${IKON[nama] || ""}</svg>`;

  // ------------------------------------------------------------ state panel
  const state = {
    rows: [], // hasil parse Excel
    files: new Map(), // nama file (lowercase) -> File
    fotoBaris: new Map(), // nomor baris Excel -> File (dipilih manual per baris)
    drive: new Map(), // id folder Drive -> {status: "memuat"|"ok"|"gagal", files, error}
    driveBaris: new Map(), // nomor baris -> {id, name, jenis} file Drive yang dipilih manual di galeri
    excelBuffer: null,
    sheet: "",
    sheets: [],
    dosen: [], // {id, nama, checked}
    existingLogs: [],
    running: false,
    stopRequested: false,
    excelName: "",
  };

  // -------------------------------------------------- info dari halaman portal
  function getTambahUrl() {
    const a = document.querySelector('a[onclick*="LogAktivitasKampusMerdeka/Tambah"]');
    if (!a) return null;
    const m = a.getAttribute("onclick").match(/['"]([^'"]*LogAktivitasKampusMerdeka\/Tambah[^'"]*)['"]/);
    return m ? m[1] : null;
  }

  function getInfoAktivitas() {
    const teks = (document.querySelector(".content-wrapper") || document.body).innerText;
    const ambil = (label) => {
      const m = teks.match(new RegExp(label + "\\s*\\n\\s*([^\\n]+)"));
      return m ? m[1].trim() : "";
    };
    return { nama: ambil("Nama Kegiatan"), waktu: ambil("Waktu"), semester: ambil("Tahun Semester") };
  }

  /** Log yang sudah ada di tabel halaman: [{waktu: "dd/mm/yyyy HH:mm-HH:mm", ket}] */
  function getExistingLogs() {
    const list = [];
    document.querySelectorAll("table tbody tr").forEach((tr) => {
      const cells = [...tr.cells];
      const iW = cells.findIndex((c) => /\d{1,2}\s+[A-Za-z]+\s+\d{4},\s*\d{1,2}[:.]\d{2}/.test(c.innerText));
      if (iW < 0) return;
      const m = cells[iW].innerText.match(/(\d{1,2})\s+([A-Za-z]+)\s+(\d{4}),\s*(\d{1,2}[:.]\d{2})\s*-\s*(\d{1,2}[:.]\d{2})/);
      if (!m) return;
      const bln = BULAN[m[2].toLowerCase()];
      if (!bln) return;
      list.push({
        waktu: `${pad2(+m[1])}/${pad2(bln)}/${m[3]} ${fmtJam(parseJam(m[4]))}-${fmtJam(parseJam(m[5]))}`,
        ket: norm(cells[iW + 1]?.innerText || ""),
      });
    });
    return list;
  }

  const kunciWaktu = (d) => `${d.Waktu} ${d.Tmw}-${d.Tsw}`;

  /** Sudah ada di portal = tanggal, jam mulai & selesai sama, dan keterangannya sama. */
  function sudahAdaDiPortal(d) {
    const k = kunciWaktu(d);
    const ket = norm(d.Keterangan);
    return state.existingLogs.some((e) => e.waktu === k && (e.ket.startsWith(ket) || ket.startsWith(e.ket)));
  }

  /** Ambil form tambah log (modal) terbaru: token, AktivitasId, daftar dosen. */
  async function fetchFormTambah() {
    const url = getTambahUrl();
    if (!url) throw new Error("Tombol Tambah tidak ditemukan. Log aktivitas ini mungkin sudah difinalisasi.");
    const res = await fetch(url, { credentials: "same-origin", headers: { "X-Requested-With": "XMLHttpRequest" } });
    const html = await res.text();
    const doc = new DOMParser().parseFromString(html, "text/html");
    const form = doc.querySelector("form#main_form") || doc.querySelector('form[action*="LogAktivitasKampusMerdeka/Tambah"]');
    if (!res.ok || !form) {
      throw new Error("Form tambah log tidak bisa dibuka. Pastikan Anda masih login, lalu muat ulang halaman.");
    }
    return { form, doc };
  }

  function readDosen(form) {
    const list = [];
    form.querySelectorAll('input[name^="ListDosenPembimbing["][name$="].Key.PembimbingId"]').forEach((inp) => {
      const idx = inp.name.match(/\[(\d+)\]/)[1];
      const cb = form.querySelector(`input[type=checkbox][name="ListDosenPembimbing[${idx}].Value"]`);
      const label = cb ? (form.querySelector(`label[for="${cb.id}"]`) || cb.closest("div") || cb.parentElement)?.textContent : "";
      list.push({ idx, id: inp.value, nama: (label || "").trim().replace(/\s+/g, " ") || `Dosen ${+idx + 1}` });
    });
    return list;
  }

  // ------------------------------------------------------------ validasi
  // hanya ekstensi file yang dikenal yang dibuang, supaya "D.1. HARMONI 1 (2)" tetap utuh
  const tanpaEkstensi = (s) => String(s).trim().toLowerCase().replace(/\.(jpe?g|png|pdf|heic|webp)$/i, "").trim();

  /** Cocokkan isi kolom bukti dengan foto yang dipilih (dengan/tanpa ekstensi, abaikan huruf besar). */
  function cariFoto(nama) {
    const persis = state.files.get(String(nama).trim().toLowerCase());
    if (persis) return persis;
    const kunci = tanpaEkstensi(nama);
    for (const [n, f] of state.files) if (tanpaEkstensi(n) === kunci) return f;
    return null;
  }

  function validasi() {
    const periode = parsePeriode(getInfoAktivitas().waktu);
    const jadwal = [];
    const pemakaianFolder = new Map();
    for (const row of state.rows) {
      const err = [];
      const warn = [];
      const r = row.raw;
      const tgl = parseTanggal(r.tanggal);
      const mulai = parseJam(r.mulai);
      const selesai = parseJam(r.selesai);
      if (!tgl) err.push("Tanggal tidak valid (pakai format DD/MM/YYYY)");
      if (mulai === null) err.push("Jam mulai tidak valid (pakai HH:MM)");
      if (selesai === null) err.push("Jam selesai tidak valid (pakai HH:MM)");
      if (mulai !== null && selesai !== null && selesai <= mulai) err.push("Jam selesai harus setelah jam mulai");

      const jenis = parseJenis(r.jenis);
      if (!jenis.item) err.push("Jenis kegiatan tidak dikenali");
      else if (jenis.defaulted) warn.push("Jenis kosong → dipakai Berita Acara Kegiatan");

      const { tipe, lokasi } = tipeLokasi(r.tipe, r.lokasi);
      if (!tipe) err.push('Tipe tidak diketahui: isi kolom Tipe, atau tulis "Online, Zoom" di kolom Lokasi/Tempat');
      if (!lokasi) err.push("Lokasi kosong");
      if (!r.keterangan) err.push("Keterangan kosong");

      // sumber bukti, urutan prioritas: pilih manual (komputer) > pilih manual (galeri Drive)
      // > foto lokal yang namanya cocok > link Google Drive di sel > tanpa bukti
      let file = state.fotoBaris.get(row.baris) || null;
      let drive = file ? null : state.driveBaris.get(row.baris) || null;
      let memuat = false;
      if (!file && !drive && r.bukti) file = cariFoto(r.bukti);
      const link = P.parseDriveLink(r.buktiLink);
      if (!file && !drive && link) {
        if (link.kind === "file") {
          drive = { id: link.id, name: r.bukti || "bukti", jenis: "auto" };
        } else {
          const f = state.drive.get(link.id);
          const media = f?.files?.filter((x) => x.jenis === "gambar" || x.jenis === "pdf") || [];
          if (!f || f.status === "memuat") memuat = true;
          else if (f.status === "gagal") err.push(`Folder Drive "${r.bukti}" ${f.error}`);
          else if (!media.length) err.push(`Folder Drive "${r.bukti}" tidak berisi foto/PDF`);
          else {
            // baris yang memakai folder yang sama diberi foto yang berbeda-beda
            const ke = pemakaianFolder.get(link.id) || 0;
            pemakaianFolder.set(link.id, ke + 1);
            drive = media[ke % media.length];
          }
        }
      }
      row.butuhFoto = !file && !drive && !memuat && !!r.bukti;
      if (file) {
        const ext = file.name.split(".").pop().toLowerCase();
        if (!EKSTENSI_BUKTI.includes(ext) && !EKSTENSI_GAMBAR_LAIN.includes(ext) && !(file.type || "").startsWith("image/"))
          err.push("Bukti harus foto atau PDF");
        else if (ext === "pdf" && file.size > BATAS_UPLOAD_BYTES) err.push("PDF lebih dari 10 MB");
      } else if (!drive && !memuat) {
        if (r.bukti && !link) err.push(`Foto "${r.bukti}" tidak ditemukan. Pilih foto di langkah 2, atau klik "${LOKAL}" di kolom Bukti`);
        else if (!r.bukti) warn.push("Tanpa foto bukti");
      }

      if (tgl && periode) {
        const t = Date.UTC(tgl.y, tgl.m - 1, tgl.d);
        if (t < periode.mulai || t > periode.selesai) warn.push("Tanggal di luar periode aktivitas");
      }

      // jam yang bertabrakan dengan baris lain di tanggal yang sama -> peringatan (tetap boleh dikirim)
      if (tgl && mulai !== null && selesai !== null && selesai > mulai) {
        const tgK = fmtTanggal(tgl);
        const tabrak = jadwal.filter((j) => j.tgl === tgK && mulai < j.selesai && selesai > j.mulai).map((j) => j.baris);
        if (tabrak.length) warn.push(`Jam bertabrakan dengan baris ${tabrak.join(", ")}`);
        jadwal.push({ tgl: tgK, mulai, selesai, baris: row.baris });
      }

      row.data = err.length
        ? null
        : {
            Waktu: fmtTanggal(tgl),
            Tmw: fmtJam(mulai),
            Tsw: fmtJam(selesai),
            jenis: jenis.item,
            tipe,
            Lokasi: lokasi,
            Keterangan: r.keterangan,
            file,
            drive,
            durasi: selesai - mulai,
          };
      row.errors = err;
      row.warnings = warn;
      if (!row.status || ["siap", "error", "duplikat", "memuat"].includes(row.status)) {
        const duplikat = row.data && sudahAdaDiPortal(row.data);
        row.status = err.length ? "error" : duplikat ? "duplikat" : memuat ? "memuat" : "siap";
        row.pesan = duplikat ? "Sudah ada di portal, dilewati" : memuat ? "Membuka folder Google Drive…" : "";
      }
    }
  }

  // ------------------------------------------------------------ foto
  async function siapkanFile(file) {
    if (!file) return null;
    const ext = file.name.split(".").pop().toLowerCase();
    const gambar = ["jpg", "jpeg", "png"].includes(ext) || (file.type || "").startsWith("image/");
    // portal hanya menerima jpg/png/pdf: gambar lain (webp, gif, …) & foto besar diubah ke JPEG
    const perluUbah = gambar && (!["jpg", "jpeg", "png"].includes(ext) || file.size > BATAS_KOMPRES_BYTES);
    if (!perluUbah) return file;
    try {
      const bmp = await createImageBitmap(file);
      const skala = Math.min(1, 1920 / Math.max(bmp.width, bmp.height));
      const w = Math.round(bmp.width * skala);
      const h = Math.round(bmp.height * skala);
      const canvas = new OffscreenCanvas(w, h);
      const ctx = canvas.getContext("2d");
      ctx.fillStyle = "#fff";
      ctx.fillRect(0, 0, w, h);
      ctx.drawImage(bmp, 0, 0, w, h);
      const blob = await canvas.convertToBlob({ type: "image/jpeg", quality: 0.85 });
      const nama = file.name.replace(/\.[^.]+$/, "") + ".jpg";
      return new File([blob], nama, { type: "image/jpeg" });
    } catch (e) {
      console.warn("[AutoKonvert] gagal kompres, pakai file asli", e);
      return file;
    }
  }

  // ------------------------------------------------------------ Google Drive
  async function mintaBackground(pesan) {
    try {
      const res = await HOST.minta(pesan);
      if (!res) throw new Error("tidak ada jawaban dari ekstensi");
      return res;
    } catch (e) {
      // biasanya karena ekstensi di-reload tapi halaman belum
      if (/context invalidated|Receiving end/i.test(e.message)) {
        throw new Error("Ekstensi baru diperbarui. Muat ulang halaman ini (F5).");
      }
      throw e;
    }
  }

  /** Buka semua folder Drive yang dirujuk Excel (maks. 3 sekaligus), lalu perbarui tabel. */
  async function muatFolderDrive() {
    const ids = [
      ...new Set(
        state.rows
          .map((r) => P.parseDriveLink(r.raw.buktiLink))
          .filter((l) => l && l.kind === "folder" && !state.drive.has(l.id))
          .map((l) => l.id)
      ),
    ];
    if (!ids.length) return;
    ids.forEach((id) => state.drive.set(id, { status: "memuat" }));
    validasi();
    render();
    const antre = [...ids];
    const pekerja = async () => {
      while (antre.length) {
        const id = antre.shift();
        try {
          const res = await mintaBackground({ type: "drive-folder", id });
          if (res.ok) state.drive.set(id, { status: "ok", files: P.parseDriveFolder(res.html) });
          else
            state.drive.set(id, {
              status: "gagal",
              error:
                res.error === "privat"
                  ? "tidak bisa dibuka karena aksesnya dibatasi"
                  : `tidak bisa dibuka (${res.error})`,
            });
        } catch (e) {
          state.drive.set(id, { status: "gagal", error: `tidak bisa dibuka (${e.message})` });
        }
        if (!state.running) {
          validasi();
          render();
        }
      }
    };
    await Promise.all([pekerja(), pekerja(), pekerja()]);
  }

  function base64KeBlob(b64, type) {
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return new Blob([bytes], { type });
  }

  const EKSTENSI_TIPE = { "image/jpeg": "jpg", "image/png": "png", "application/pdf": "pdf", "image/webp": "webp", "image/gif": "gif" };

  /** Unduh file Drive (lewat background) menjadi File. */
  async function unduhDrive(item) {
    const mode = item.jenis === "gambar" ? "gambar" : "unduh";
    const res = await mintaBackground({ type: "drive-file", id: item.id, mode });
    if (!res.ok) throw new Error(`Foto Drive "${item.name}" ${res.error}.`);
    const ext = EKSTENSI_TIPE[res.type];
    if (!ext) throw new Error(`File Drive "${item.name}" bukan foto/PDF (${res.type || "tipe tidak dikenal"}).`);
    let nama = (res.name || item.name || "bukti").replace(/[\\/:*?"<>|]+/g, " ").trim();
    if (!nama.toLowerCase().endsWith("." + ext)) nama = nama.replace(/\.(jpe?g|png|pdf|webp|gif|heic)$/i, "") + "." + ext;
    return new File([base64KeBlob(res.base64, res.type)], nama, { type: res.type });
  }

  const thumbCache = new Map();
  async function thumbDrive(id) {
    if (!thumbCache.has(id)) {
      thumbCache.set(
        id,
        mintaBackground({ type: "drive-file", id, mode: "thumb" }).then((res) =>
          res.ok ? URL.createObjectURL(base64KeBlob(res.base64, res.type)) : null
        )
      );
    }
    return thumbCache.get(id);
  }

  // ------------------------------------------------------------ kirim
  async function kirimSatu(row) {
    const { form } = await fetchFormTambah();
    const d = row.data;
    const set = (name, value) => {
      const el = form.querySelector(`[name="${name}"]`);
      if (!el) throw new Error(`Field "${name}" tidak ada di form portal (form mungkin berubah).`);
      el.value = value;
    };
    set("Waktu", d.Waktu);
    set("Tmw", d.Tmw);
    set("Tsw", d.Tsw);
    set("JenisLogbookKegiatanKampusMerdekaId", d.jenis.id);
    set("Lokasi", d.Lokasi);
    set("Keterangan", d.Keterangan);

    const radios = [...form.querySelectorAll('input[type=radio][name="IsLuring"]')];
    const radio = radios.find((r) => r.value === d.tipe.value);
    if (!radio) throw new Error("Pilihan tipe penyelenggaraan tidak ditemukan di form portal.");
    radios.forEach((r) => (r.checked = r === radio));

    const dosenDipilih = new Set(state.dosen.filter((x) => x.checked).map((x) => x.id));
    form.querySelectorAll('input[name^="ListDosenPembimbing["][name$="].Key.PembimbingId"]').forEach((inp) => {
      const idx = inp.name.match(/\[(\d+)\]/)[1];
      const cb = form.querySelector(`input[type=checkbox][name="ListDosenPembimbing[${idx}].Value"]`);
      if (cb) cb.checked = dosenDipilih.has(inp.value);
    });

    // FormData dari form asli => ikut membawa token & field tersembunyi lain
    const fd = new FormData(form);
    const file = await siapkanFile(d.file || (d.drive ? await unduhDrive(d.drive) : null));
    if (file) {
      if (file.size > BATAS_UPLOAD_BYTES) throw new Error("File bukti lebih dari 10 MB setelah dikompres.");
      fd.set("File", file, file.name);
    }

    // alamat kirim harus tetap di Student Portal
    const action = new URL(form.getAttribute("action") || "/Kegiatan/LogAktivitasKampusMerdeka/Tambah", PORTAL_ORIGIN);
    if (action.origin !== PORTAL_ORIGIN) throw new Error("Alamat form tidak dikenal, pengiriman dibatalkan.");

    const res = await fetch(action.href, {
      method: "POST",
      body: fd,
      credentials: "same-origin",
      headers: { "X-Requested-With": "XMLHttpRequest" },
    });
    const teks = await res.text();
    let json = null;
    try {
      json = JSON.parse(teks);
    } catch {
      /* bukan JSON */
    }
    if (!json) {
      if (/Account\/Login|Login Student Portal/i.test(teks) || res.url.includes("/Account/Login")) {
        const e = new Error("Sesi login habis. Login ulang lalu muat ulang halaman ini.");
        e.fatal = true;
        throw e;
      }
      const doc = new DOMParser().parseFromString(teks, "text/html");
      const pesan = (doc.body?.innerText || "").match(/Error:\s*([^\n]+)/);
      throw new Error(pesan ? pesan[1].trim() : `Respons tidak dikenal dari portal (HTTP ${res.status}).`);
    }
    if (!json.status) throw new Error(json.message || "Ditolak oleh portal.");
    return json.message || "Berhasil";
  }

  async function mulaiKirim() {
    const antrian = state.rows.filter((r) => r.status === "siap" || r.status === "gagal");
    if (!antrian.length) return;
    if (!state.dosen.some((d) => d.checked) && state.dosen.length) {
      alert("Pilih minimal satu Dosen Penggerak terlebih dahulu.");
      return;
    }
    const info = getInfoAktivitas();
    const ok = confirm(
      `Kirim ${antrian.length} log` +
        (state.sheets.length > 1 ? ` dari sheet "${state.sheet}"` : "") +
        ` ke aktivitas:\n"${info.nama}"?\n\n` +
        "Pastikan isi logbook sudah benar. Log yang sudah terkirim hanya bisa dihapus manual satu per satu dari halaman ini."
    );
    if (!ok) return;

    state.running = true;
    state.stopRequested = false;
    render();
    for (let i = 0; i < antrian.length; i++) {
      if (state.stopRequested) break;
      const row = antrian[i];
      row.status = "mengirim";
      row.pesan = "";
      render();
      try {
        row.pesan = await kirimSatu(row);
        row.status = "berhasil";
        state.existingLogs.push({ waktu: kunciWaktu(row.data), ket: norm(row.data.Keterangan) });
      } catch (e) {
        row.status = "gagal";
        row.pesan = e.message;
        if (e.fatal) {
          state.stopRequested = true;
          alert(e.message);
        }
      }
      render();
      if (i < antrian.length - 1 && !state.stopRequested) await sleep(JEDA_ANTAR_KIRIM_MS);
    }
    state.running = false;
    render();
  }

  function unduhLaporan() {
    const header = ["Baris Excel", "Tanggal", "Jam Mulai", "Jam Selesai", "Keterangan", "Status", "Pesan"];
    const data = state.rows.map((r) => [
      r.baris,
      r.data?.Waktu ?? r.raw.tanggal,
      r.data?.Tmw ?? r.raw.mulai,
      r.data?.Tsw ?? r.raw.selesai,
      r.raw.keterangan,
      r.status,
      r.pesan || [...(r.errors || []), ...(r.warnings || [])].join("; "),
    ]);
    const ws = XLSX.utils.aoa_to_sheet([header, ...data]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Laporan");
    XLSX.writeFile(wb, "laporan-autokonvert.xlsx");
  }

  // ------------------------------------------------------------ UI
  const host = document.createElement("div");
  host.id = "autokonvert-root";
  document.documentElement.appendChild(host);
  const root = host.attachShadow({ mode: "open" });

  root.innerHTML = `
    <style>
      :host { all: initial; }
      * { box-sizing: border-box; font-family: "IBM Plex Sans", "Segoe UI", system-ui, -apple-system, sans-serif; }
      .ic { width: 16px; height: 16px; flex: none; fill: none; stroke: currentColor; stroke-width: 1.8;
        stroke-linecap: round; stroke-linejoin: round; vertical-align: -3px; }
      @keyframes putar { to { transform: rotate(360deg); } }
      .spin { animation: putar 1s linear infinite; }

      /* tombol mengambang */
      .fab { position: fixed; right: 20px; bottom: 20px; z-index: 2147483000; display: flex; align-items: center; gap: 8px;
        height: 44px; padding: 0 16px; border: 0; border-radius: 8px; cursor: pointer; background: #0B2F5B; color: #fff;
        font-size: 14px; font-weight: 500; box-shadow: 0 4px 14px rgba(11,47,91,.28); }
      .fab:hover { background: #08254A; }
      .fab .ic { width: 17px; height: 17px; stroke-width: 2; }

      /* panel */
      .panel { position: fixed; right: 20px; bottom: 20px; z-index: 2147483001; width: min(760px, calc(100vw - 40px));
        max-height: calc(100vh - 40px); background: #fff; color: #0F1B2D; border: 1px solid #E1E7EF; border-radius: 10px;
        display: none; flex-direction: column; overflow: hidden; font-size: 13.5px; line-height: 1.5;
        box-shadow: 0 16px 40px rgba(15,27,45,.16); }
      .panel.open { display: flex; }
      .head { display: flex; align-items: center; gap: 12px; padding: 12px 16px; border-bottom: 1px solid #E1E7EF; }
      .mark { width: 30px; height: 30px; border-radius: 6px; background: #0B2F5B; color: #fff; display: grid; place-items: center; flex: none; }
      .mark .ic { stroke-width: 2; }
      .head h2 { margin: 0; font-size: 15px; font-weight: 600; letter-spacing: -.01em; }
      .head h2 small { font-weight: 400; color: #55627A; font-size: 12px; margin-left: 4px; }
      .head .sub { display: block; font-size: 12px; color: #55627A; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
      .head .judul { flex: 1; min-width: 0; }
      .x { width: 32px; height: 32px; border: 0; background: none; color: #55627A; border-radius: 6px; cursor: pointer; display: grid; place-items: center; }
      .x:hover { background: #F5F8FC; color: #0F1B2D; }
      .x .ic { width: 18px; height: 18px; }
      .body { padding: 4px 16px 8px; overflow: auto; display: flex; flex-direction: column; }

      /* langkah: dipisah garis, bukan kotak */
      .step { padding: 14px 0; border-bottom: 1px solid #EEF2F7; }
      .step:last-child { border-bottom: 0; }
      .step h3 { margin: 0 0 10px; font-size: 13.5px; font-weight: 600; display: flex; gap: 10px; align-items: center; }
      .num { width: 22px; height: 22px; border-radius: 50%; border: 1.5px solid #0B2F5B; color: #0B2F5B; display: inline-grid;
        place-items: center; font-size: 11.5px; font-weight: 600; flex: none; }
      .num.ok { background: #0B2F5B; color: #fff; }
      .num.ok .ic { width: 12px; height: 12px; stroke-width: 2.6; }
      .muted { color: #55627A; }
      .row { display: flex; flex-wrap: wrap; gap: 10px; align-items: center; }

      label.file { display: inline-flex; align-items: center; gap: 8px; height: 36px; padding: 0 14px; border: 1px solid #C9D3E1;
        border-radius: 8px; color: #0B2F5B; background: #fff; font-weight: 500; cursor: pointer; }
      label.file:hover { border-color: #0B2F5B; }
      label.file input { display: none; }
      .dosen label { display: flex; gap: 10px; align-items: center; padding: 4px 0; cursor: pointer; }
      .dosen input { accent-color: #0B2F5B; width: 16px; height: 16px; }

      .alert { display: flex; gap: 10px; align-items: flex-start; padding: 10px 12px; border-radius: 8px;
        background: #FDF2F2; color: #8A1C1C; border: 1px solid #F5D0D0; }
      .alert > .ic { margin-top: 2px; }
      .alert.ok { background: #EFF7F2; color: #1D5B37; border-color: #CFE6D8; }
      .alert.info { background: #EAF0F8; color: #1D3557; border-color: #D5E0EF; }

      .summary { display: flex; gap: 6px 14px; flex-wrap: wrap; margin-bottom: 10px; font-size: 12.5px; color: #55627A; }
      .summary span { display: inline-flex; align-items: center; gap: 6px; }
      .dot { width: 7px; height: 7px; border-radius: 50%; display: inline-block; background: #94A3B8; }
      .chip { display: inline-flex; align-items: center; gap: 6px; font-size: 12px; font-weight: 500; color: #334155; }
      .chip::before { content: ""; width: 7px; height: 7px; border-radius: 50%; background: #94A3B8; }
      .c-siap::before, .dot.siap { background: #2563EB; }
      .c-error::before, .c-gagal::before, .dot.error { background: #DC2626; }
      .c-berhasil::before, .dot.berhasil { background: #16A34A; }
      .c-duplikat::before { background: #94A3B8; }
      .c-mengirim::before, .c-memuat::before, .dot.memuat { background: #D97706; }

      table { width: 100%; border-collapse: collapse; font-size: 12.5px; }
      th, td { text-align: left; padding: 8px 10px; border-bottom: 1px solid #EEF2F7; vertical-align: top; }
      th { font-weight: 500; color: #55627A; background: #F5F8FC; }
      .tablewrap { overflow-x: auto; border: 1px solid #E1E7EF; border-radius: 8px; }
      td:nth-child(2) { white-space: nowrap; }
      .ket { min-width: 200px; }
      td.bukti { min-width: 170px; max-width: 220px; word-break: break-word; }
      .src { display: flex; gap: 6px; align-items: flex-start; }
      .src .ic { color: #0B2F5B; margin-top: 1px; }
      .err, .warn { display: flex; gap: 5px; font-size: 11.5px; margin-top: 3px; }
      .err { color: #B91C1C; } .warn { color: #92600A; }
      .err .ic, .warn .ic { width: 13px; height: 13px; margin-top: 2px; }

      .foot { padding: 12px 16px; border-top: 1px solid #E1E7EF; display: flex; gap: 10px; justify-content: flex-end;
        flex-wrap: wrap; align-items: center; background: #FBFCFE; }
      .foot .chip { margin-right: auto; }
      button.btn { display: inline-flex; align-items: center; gap: 8px; height: 38px; padding: 0 16px; border-radius: 8px;
        border: 1px solid transparent; font-weight: 500; font-size: 13.5px; cursor: pointer; }
      .primary { background: #0B2F5B; color: #fff; } .primary:hover { background: #08254A; }
      .primary:disabled { background: #A8B4C5; cursor: not-allowed; }
      .secondary { background: #fff; color: #0B2F5B; border-color: #C9D3E1 !important; } .secondary:hover { border-color: #0B2F5B !important; }
      .danger { background: #fff; color: #B91C1C; border-color: #F0C4C4 !important; }
      button.mini, a.mini { display: inline-flex; align-items: center; gap: 5px; margin: 6px 6px 0 0; padding: 3px 9px;
        border: 1px solid #C9D3E1; border-radius: 6px; background: #fff; color: #0B2F5B; font-size: 12px; font-weight: 500;
        cursor: pointer; text-decoration: none; white-space: nowrap; }
      button.mini:hover, a.mini:hover { border-color: #0B2F5B; }
      .mini .ic { width: 13px; height: 13px; }
      select { height: 34px; padding: 0 8px; border-radius: 8px; border: 1px solid #C9D3E1; font-size: 13.5px; background: #fff; color: #0F1B2D; }
      .progress { flex: 1; min-width: 160px; height: 6px; background: #E1E7EF; border-radius: 99px; overflow: hidden; }
      .progress > div { height: 100%; background: #0B2F5B; transition: width .3s; }

      a.logitem { display: flex; justify-content: space-between; align-items: center; gap: 10px; padding: 10px 12px;
        border: 1px solid #E1E7EF; border-radius: 8px; margin-top: 8px; color: #0F1B2D; text-decoration: none; }
      a.logitem:hover { border-color: #0B2F5B; }
      a.logitem b { font-weight: 500; }
      ol.mulai { margin: 0; padding-left: 18px; } ol.mulai li { margin-bottom: 4px; }

      /* galeri foto Drive */
      .galeri { position: fixed; inset: 0; z-index: 2147483002; background: rgba(15,27,45,.5); display: none;
        align-items: center; justify-content: center; padding: 20px; }
      .galeri.open { display: flex; }
      .gbox { background: #fff; color: #0F1B2D; border-radius: 10px; width: min(720px, 100%); max-height: 100%;
        display: flex; flex-direction: column; overflow: hidden; font-size: 13.5px; }
      .ghead { display: flex; justify-content: space-between; align-items: center; gap: 12px; padding: 12px 16px; border-bottom: 1px solid #E1E7EF; }
      .ghead b { font-weight: 600; }
      .ggrid { display: grid; grid-template-columns: repeat(auto-fill, minmax(132px, 1fr)); gap: 10px; padding: 14px 16px; overflow: auto; }
      .gitem { border: 1px solid #E1E7EF; border-radius: 8px; padding: 6px; background: #fff; cursor: pointer; text-align: left; }
      .gitem:hover { border-color: #0B2F5B; }
      .gitem.aktif { border-color: #0B2F5B; box-shadow: inset 0 0 0 1px #0B2F5B; }
      .gthumb { height: 100px; background: #F5F8FC; border-radius: 5px; display: grid; place-items: center; overflow: hidden; color: #8592A8; }
      .gthumb .ic { width: 22px; height: 22px; }
      .gthumb img { width: 100%; height: 100%; object-fit: cover; }
      .gname { margin-top: 5px; font-size: 11.5px; word-break: break-all; color: #334155; }

      button:focus-visible, label.file:focus-within, a:focus-visible, select:focus-visible { outline: 2px solid #0B2F5B; outline-offset: 2px; }

      /* layar HP: panel penuh layar, tabel jadi kartu, sasaran sentuh lebih besar */
      @media (max-width: 640px) {
        .fab { left: 12px; right: 12px; bottom: calc(12px + env(safe-area-inset-bottom)); justify-content: center; height: 50px; font-size: 15px; }
        .panel { inset: 0; width: 100%; max-height: none; height: 100%; border-radius: 0; border: 0; font-size: 14.5px; }
        .head { padding: calc(10px + env(safe-area-inset-top)) 12px 10px; }
        .x { width: 40px; height: 40px; }
        .body { padding: 0 14px 8px; -webkit-overflow-scrolling: touch; }
        label.file { width: 100%; justify-content: center; height: 46px; font-size: 15px; }
        select { width: 100%; height: 42px; font-size: 16px; }
        .dosen label { padding: 8px 0; }
        .dosen input { width: 20px; height: 20px; }
        .tablewrap { max-height: none; border: 0; overflow: visible; }
        table, tbody, tr, td { display: block; width: 100%; }
        thead { display: none; }
        tr { border: 1px solid #E1E7EF; border-radius: 8px; padding: 10px 12px; margin-bottom: 10px; }
        td { border: 0; padding: 3px 0; max-width: none !important; }
        td:first-child { font-size: 11.5px; color: #55627A; }
        td:first-child::before { content: "Baris "; }
        button.mini, a.mini { padding: 8px 12px; font-size: 13.5px; }
        .foot { padding: 10px 12px calc(10px + env(safe-area-inset-bottom)); }
        .foot button.btn { flex: 1; justify-content: center; height: 48px; font-size: 15px; }
        .foot .chip { flex-basis: 100%; }
        .progress { flex-basis: 100%; }
        .galeri { padding: 0; }
        .gbox { width: 100%; height: 100%; border-radius: 0; }
        .ghead { padding-top: calc(12px + env(safe-area-inset-top)); }
        .ggrid { grid-template-columns: repeat(2, 1fr); }
        .gthumb { height: 130px; }
      }
    </style>
    <button class="fab" id="fab" aria-label="Buka AutoKonvert">${ic("zap")}AutoKonvert</button>
    <div class="panel" id="panel" role="dialog" aria-label="AutoKonvert Logbook">
      <div class="head">
        <span class="mark">${ic("zap")}</span>
        <div class="judul"><h2>AutoKonvert Logbook<small>v${esc(HOST.versi)}${HOST.jenis === "userscript" ? " · HP" : ""}</small></h2><span class="sub" id="aktivitas"></span></div>
        <button class="x" id="close" title="Tutup" aria-label="Tutup">${ic("x")}</button>
      </div>
      <div class="body" id="body"></div>
      <div class="foot" id="foot"></div>
    </div>
    <div class="galeri" id="galeri"></div>`;

  const $ = (id) => root.getElementById(id);
  $("fab").addEventListener("click", () => {
    $("panel").classList.add("open");
    $("fab").style.display = "none";
  });
  $("close").addEventListener("click", () => {
    if (state.running && !confirm("Pengiriman masih berjalan. Tutup panel? (pengiriman tetap berjalan)")) return;
    $("panel").classList.remove("open");
    $("fab").style.display = "";
  });

  let fatalMsg = "";
  let excelMsg = "";

  const LABEL_STATUS = {
    siap: "Siap",
    error: "Perlu diperbaiki",
    duplikat: "Sudah ada",
    memuat: "Membuka Drive",
    mengirim: "Mengirim",
    berhasil: "Terkirim",
    gagal: "Gagal",
  };

  function hitung() {
    const c = { siap: 0, error: 0, duplikat: 0, berhasil: 0, gagal: 0, mengirim: 0 };
    state.rows.forEach((r) => (c[r.status] = (c[r.status] || 0) + 1));
    return c;
  }

  // nilai yang ditampilkan di tabel: pakai hasil parse kalau bisa, kalau tidak tampilkan apa adanya
  const tampilTanggal = (r) => { const t = parseTanggal(r.raw.tanggal); return t ? fmtTanggal(t) : String(r.raw.tanggal || "?"); };
  const tampilJam = (v) => { const m = parseJam(v); return m === null ? String(v || "?") : fmtJam(m); };
  const tampilTempat = (r) => { const x = tipeLokasi(r.raw.tipe, r.raw.lokasi); return (x.tipe ? x.tipe.label + ", " : "") + x.lokasi; };

  function folderBaris(r) {
    const l = P.parseDriveLink(r.raw.buktiLink);
    return l && l.kind === "folder" ? l.id : null;
  }

  function selBukti(r) {
    const d = r.data;
    let isi;
    if (d?.file) isi = `<div class="src">${ic("image")}<div>${esc(d.file.name)}<div class="muted">dari ${LOKAL_KECIL}</div></div></div>`;
    else if (d?.drive) isi = `<div class="src">${ic("cloud")}<div>${esc(d.drive.name)}<div class="muted">dari Google Drive</div></div></div>`;
    else if (r.status === "memuat") isi = `<div class="src muted">${ic("loader", "spin")}membuka folder…</div>`;
    else if (r.raw.bukti) isi = `<span class="muted">${esc(r.raw.bukti)}</span>`;
    else isi = '<span class="muted">—</span>';
    if (state.running || r.status === "berhasil" || r.status === "duplikat") return isi;
    const fid = folderBaris(r);
    const st = fid && state.drive.get(fid)?.status;
    const galeri = st === "ok" ? `<button class="mini" data-galeri="${r.baris}">${ic("grid")}Galeri</button>` : "";
    // link hanya ditampilkan kalau memang link Google Drive yang valid
    const buka =
      st === "gagal" && !d?.file
        ? `<a class="mini" href="${esc(r.raw.buktiLink)}" target="_blank" rel="noopener noreferrer">${ic("external")}Buka folder</a>`
        : "";
    return `${isi}<div>${galeri}${buka}<button class="mini" data-pilih-foto="${r.baris}">${ic("upload")}${LOKAL}</button></div>`;
  }

  // ---- galeri foto dari folder Drive (pilih foto untuk satu baris)
  async function bukaGaleri(baris) {
    const row = state.rows.find((r) => r.baris === baris);
    const fid = row && folderBaris(row);
    const folder = fid && state.drive.get(fid);
    if (!folder || folder.status !== "ok") return;
    const media = folder.files.filter((x) => x.jenis === "gambar" || x.jenis === "pdf");
    const dipilih = row.data?.drive?.id;
    const g = $("galeri");
    g.innerHTML = `
      <div class="gbox">
        <div class="ghead"><div><b>Pilih foto bukti</b><div class="muted">Baris ${baris} · ${esc(row.raw.bukti)} · ${media.length} file</div></div>
          <button class="x" id="gtutup" title="Tutup" aria-label="Tutup">${ic("x")}</button></div>
        <div class="ggrid">${media
          .map(
            (f) => `<button class="gitem ${f.id === dipilih ? "aktif" : ""}" data-id="${esc(f.id)}">
              <div class="gthumb" data-thumb="${esc(f.id)}">${f.jenis === "pdf" ? ic("file") : ic("loader", "spin")}</div>
              <div class="gname">${esc(f.name)}</div></button>`
          )
          .join("")}</div>
      </div>`;
    g.classList.add("open");
    $("gtutup").onclick = () => g.classList.remove("open");
    g.onclick = (e) => {
      if (e.target === g) g.classList.remove("open");
    };
    g.querySelectorAll(".gitem").forEach((b) =>
      b.addEventListener("click", () => {
        const f = media.find((x) => x.id === b.dataset.id);
        state.driveBaris.set(baris, f);
        state.fotoBaris.delete(baris);
        g.classList.remove("open");
        validasi();
        render();
      })
    );
    // muat thumbnail satu per satu
    for (const f of media) {
      if (f.jenis !== "gambar" || !g.classList.contains("open")) continue;
      const url = await thumbDrive(f.id).catch(() => null);
      const el = g.querySelector(`[data-thumb="${CSS.escape(f.id)}"]`);
      if (el) el.innerHTML = url ? `<img src="${url}" alt="">` : ic("image");
    }
  }

  /**
   * Langkah 2 (foto) hanya muncul setelah Excel dipilih, dan hanya meminta foto dari komputer
   * kalau memang ada baris yang fotonya belum didapat (bukan link Drive / Drive gagal dibuka).
   */
  function langkahFoto(driveGagal, jumlahFoto) {
    if (!state.rows.length) return "";
    const aktif = state.rows.filter((r) => !["berhasil", "duplikat"].includes(r.status));
    const dariDrive = aktif.filter((r) => r.data?.drive).length;
    const dariKomputer = aktif.filter((r) => r.data?.file).length;
    const memuat = aktif.filter((r) => r.status === "memuat").length;
    const kurang = aktif.filter((r) => r.butuhFoto);
    const tanpa = aktif.filter((r) => !r.raw.bukti && !r.data?.file && !r.data?.drive).length;

    const ringkas = [
      dariDrive && `${ic("cloud")} ${dariDrive} dari Google Drive`,
      dariKomputer && `${ic("image")} ${dariKomputer} dari ${LOKAL_KECIL}`,
      memuat && `⏳ ${memuat} sedang dibuka`,
      tanpa && `${tanpa} tanpa bukti`,
    ]
      .filter(Boolean)
      .join(" · ");
    const tombolFoto = `<label class="file">${ic("upload")}Pilih foto dari perangkat<input type="file" id="foto" multiple accept="${ACCEPT_BUKTI}" ${state.running ? "disabled" : ""}></label>`;
    const ulang = driveGagal
      ? `<div class="alert info" style="margin-top:10px">${ic("info")}<span>${driveGagal} folder Google Drive tidak bisa dibuka karena aksesnya dibatasi.
          Login ke Google di browser ini dengan akun yang punya akses, lalu
          <button class="mini" id="ulangDrive" ${state.running ? "disabled" : ""}>${ic("refresh")}Coba buka ulang</button></span></div>`
      : "";

    if (!kurang.length) {
      return `<div class="step">
        <h3>${memuat ? '<span class="num">2</span>' : `<span class="num ok">${ic("check")}</span>`} Foto bukti</h3>
        <div class="muted">${ringkas || "Tidak ada kolom bukti di Excel."}</div>
        ${ulang}
      </div>`;
    }
    const contoh = [...new Set(kurang.map((r) => r.raw.bukti))].slice(0, 4);
    return `<div class="step">
        <h3><span class="num">2</span> ${kurang.length} baris belum punya foto</h3>
        ${ringkas ? `<div class="muted" style="margin-bottom:8px">${ringkas}</div>` : ""}
        <div class="row">${tombolFoto}
          <span class="muted">${jumlahFoto ? `${jumlahFoto} file dipilih` : (DI_HP ? "Bisa pilih banyak foto sekaligus" : "Bisa pilih banyak sekaligus (Ctrl+A)")}</span></div>
        <div class="muted" style="margin-top:6px">Baris ini tidak punya link Google Drive yang bisa dibuka, misalnya ${contoh
          .map((s) => `"${esc(s)}"`)
          .join(", ")}. Pilih fotonya di sini kalau nama filenya sama dengan isi kolom bukti, atau klik <b>${LOKAL}</b> di tabel untuk memilih per baris.</div>
        ${ulang}
      </div>`;
  }

  /** Panel di luar halaman Log: tunjukkan link Log aktivitas yang ada di halaman ini. */
  function bantuanHtml() {
    const links = [];
    const sudah = new Set();
    document.querySelectorAll("a[href]").forEach((a) => {
      const href = a.getAttribute("href") || "";
      if (!/\/Kegiatan\/LogAktivitasKampusMerdeka\/Index\//i.test(href)) return;
      const url = new URL(href, PORTAL_ORIGIN);
      if (url.origin !== PORTAL_ORIGIN || sudah.has(url.pathname)) return;
      sudah.add(url.pathname);
      const tr = a.closest("tr");
      const sel = tr ? [...tr.cells].map((c) => c.innerText.trim()) : [];
      // kolom tabel Aktivitas: No, Tahun Semester, Nama Kegiatan, Kelompok, Soft Skill, Waktu, Jam, Status, ...
      const nama = sel[2] || a.innerText.trim() || "Log aktivitas";
      const status = sel.find((t) => /draft|disetujui|ditolak|diverifikasi/i.test(t)) || "";
      links.push({ url: url.href, nama, status, draft: /draft/i.test(status) });
    });
    links.sort((a, b) => b.draft - a.draft);
    const daftar = links.length
      ? `<div class="step"><h3>Pilih Log aktivitas</h3>${links
          .map(
            (l) => `<a class="logitem" href="${esc(l.url)}"><b>${esc(l.nama)}</b>
              <span class="chip ${l.draft ? "c-siap" : "c-duplikat"}">${esc(l.status.replace(/\s+/g, " ") || "Log")}</span></a>`
          )
          .join("")}
          <div class="muted" style="margin-top:6px">Log hanya bisa ditambah pada aktivitas berstatus <b>Draft</b>.</div></div>`
      : `<div class="step"><h3>Cara mulai</h3><ol class="mulai">
           <li>Buka <b>Kemahasiswaan → Aktivitas</b>.</li>
           <li>Ketuk tombol <b>Log</b> (ikon daftar) pada aktivitas berstatus <b>Draft</b>.</li>
           <li>Tekan lagi tombol <b>AutoKonvert</b> di halaman Log itu.</li></ol>
           <a class="logitem" href="${PORTAL_ORIGIN}/Kegiatan/AktivitasKampusMerdeka" style="margin-top:10px"><b>Buka daftar Aktivitas</b>${ic("arrow")}</a></div>`;
    return `<div style="height:14px"></div><div class="alert ok">${ic("check")}<span>AutoKonvert aktif (v${esc(HOST.versi)}). Pengisian logbook dilakukan di halaman <b>Log</b> sebuah aktivitas.</span></div>${daftar}`;
  }

  function render() {
    const info = getInfoAktivitas();
    $("aktivitas").textContent = info.nama ? `${info.nama} · ${info.waktu}` : "";
    const body = $("body");
    const foot = $("foot");

    if (!HALAMAN_LOG) {
      $("aktivitas").textContent = "Buka halaman Log aktivitas untuk mulai";
      body.innerHTML = bantuanHtml();
      foot.innerHTML = "";
      return;
    }

    if (fatalMsg) {
      body.innerHTML = `<div style="height:14px"></div><div class="alert">${ic("alert")}<span>${esc(fatalMsg)}</span></div>`;
      foot.innerHTML = "";
      return;
    }

    const c = hitung();
    const jumlahFoto = state.files.size;
    const driveGagal = [...state.drive.values()].filter((f) => f.status === "gagal").length;
    const dosenHtml = state.dosen.length
      ? state.dosen
          .map(
            (d, i) =>
              `<label><input type="checkbox" data-dosen="${i}" ${d.checked ? "checked" : ""} ${state.running ? "disabled" : ""}> ${esc(d.nama)}</label>`
          )
          .join("")
      : `<span class="muted">Memuat daftar dosen…</span>`;

    const totalMenit = state.rows.reduce((s, r) => s + (r.data && r.status !== "duplikat" ? r.data.durasi : 0), 0);

    const baris = state.rows
      .map((r) => {
        const d = r.data;
        const catatan = [
          ...(r.errors || []).map((e) => `<div class="err">${ic("xcircle")}<span>${esc(e)}</span></div>`),
          ...(r.warnings || []).map((w) => `<div class="warn">${ic("alert")}<span>${esc(w)}</span></div>`),
          r.pesan ? `<div class="${r.status === "gagal" ? "err" : "muted"}">${esc(r.pesan)}</div>` : "",
        ].join("");
        return `<tr>
          <td>${r.baris}</td>
          <td>${esc(tampilTanggal(r))}<br><span class="muted">${esc(tampilJam(r.raw.mulai))}–${esc(tampilJam(r.raw.selesai))}</span></td>
          <td class="ket">${esc(r.raw.keterangan)}<div class="muted">${esc(tampilTempat(r))}</div></td>
          <td class="bukti">${selBukti(r)}</td>
          <td><span class="chip c-${r.status}">${LABEL_STATUS[r.status] || esc(r.status)}</span>${catatan}</td>
        </tr>`;
      })
      .join("");

    body.innerHTML = `
      <div class="step">
        <h3><span class="num">1</span> Pilih file Excel logbook</h3>
        <div class="row">
          <label class="file">${ic("sheet")}Pilih Excel<input type="file" id="excel" accept="${ACCEPT_EXCEL}" ${state.running ? "disabled" : ""}></label>
          <span class="muted">${state.excelName ? esc(state.excelName) + ` · ${state.rows.length} baris` : "Gunakan template dari website AutoKonvert"}</span>
        </div>
        ${
          state.sheets.length > 1
            ? `<div class="row" style="margin-top:10px"><span class="muted">Sheet</span>
                <select id="sheet" ${state.running ? "disabled" : ""}>${state.sheets
                  .map((s) => `<option value="${esc(s)}" ${s === state.sheet ? "selected" : ""}>${esc(s)}</option>`)
                  .join("")}</select>
                </div>
               <div class="${state.sheetOtomatis ? "alert ok" : "alert info"}" style="margin-top:10px">${ic(state.sheetOtomatis ? "check" : "info")}<span>${
                 state.sheetOtomatis
                   ? `Sheet <b>${esc(state.sheet)}</b> dipilih otomatis karena cocok dengan nama aktivitas ini.`
                   : `File ini punya ${state.sheets.length} sheet. Pastikan sheet <b>${esc(state.sheet)}</b> memang untuk aktivitas <b>${esc(getInfoAktivitas().nama)}</b>. Sheet lain dikirim dari halaman Log aktivitasnya sendiri.`
               }</span></div>`
            : ""
        }
        <input type="file" id="fotoBaris" accept="${ACCEPT_BUKTI}" style="display:none">
        ${excelMsg ? `<div class="alert" style="margin-top:10px">${ic("xcircle")}<span>${esc(excelMsg)}</span></div>` : ""}
      </div>
      ${langkahFoto(driveGagal, jumlahFoto)}
      <div class="step dosen">
        <h3><span class="num">${state.rows.length ? 3 : 2}</span> Dosen Penggerak</h3>
        ${dosenHtml}
      </div>
      ${
        state.rows.length
          ? `<div class="step">
        <h3><span class="num">4</span> Cek data sebelum dikirim</h3>
        <div class="muted" style="margin-bottom:8px">Foto di kolom Bukti bisa diganti per baris: <b>Galeri</b> (dari folder Drive) atau <b>${LOKAL}</b>.</div>
        <div class="summary">
          ${["siap", "error", "duplikat", "memuat", "berhasil", "gagal"]
            .filter((k) => c[k])
            .map((k) => `<span class="chip c-${k}">${LABEL_STATUS[k]}: ${c[k]}</span>`)
            .join("")}
          <span>Total ${Math.floor(totalMenit / 60)} jam ${totalMenit % 60} menit</span>
        </div>
        ${c.error ? `<div class="alert" style="margin-bottom:10px">${ic("alert")}<span>Ada ${c.error} baris yang perlu diperbaiki. Baris ini tidak akan dikirim sampai diperbaiki: pilih fotonya, atau perbaiki Excel lalu pilih ulang filenya.</span></div>` : ""}
        <div class="tablewrap"><table>
          <thead><tr><th>Baris</th><th>Waktu</th><th>Keterangan</th><th>Bukti</th><th>Status</th></tr></thead>
          <tbody>${baris}</tbody>
        </table></div>
      </div>`
          : ""
      }`;

    const siapKirim = c.siap + c.gagal;
    const selesai = c.berhasil + c.gagal;
    const total = state.rows.filter((r) => ["siap", "berhasil", "gagal", "mengirim"].includes(r.status)).length;
    foot.innerHTML = state.running
      ? `<div class="progress"><div style="width:${total ? (selesai / total) * 100 : 0}%"></div></div>
         <span class="muted">${c.berhasil} berhasil · ${c.gagal} gagal</span>
         <button class="btn danger" id="stop">${ic("x")}Hentikan</button>`
      : `${c.berhasil ? `<span class="chip c-berhasil">${c.berhasil} log terkirim</span>` : ""}
         ${state.rows.length ? `<button class="btn secondary" id="laporan">${ic("report")}Laporan</button>` : ""}
         ${c.berhasil ? `<button class="btn secondary" id="reload">${ic("refresh")}Muat ulang</button>` : ""}
         <button class="btn primary" id="kirim" ${siapKirim ? "" : "disabled"}>${ic("send")}${c.gagal && !c.siap ? "Kirim ulang yang gagal" : `Kirim ${siapKirim} log`}</button>`;

    // event
    root.getElementById("excel")?.addEventListener("change", onExcel);
    root.getElementById("foto")?.addEventListener("change", onFoto);
    root.getElementById("sheet")?.addEventListener("change", (e) => {
      state.sheetOtomatis = false;
      muatSheet(e.target.value);
    });
    root.getElementById("ulangDrive")?.addEventListener("click", () => {
      for (const [id, f] of state.drive) if (f.status === "gagal") state.drive.delete(id);
      muatFolderDrive();
    });
    root.querySelectorAll("[data-galeri]").forEach((b) => b.addEventListener("click", () => bukaGaleri(+b.dataset.galeri)));
    const inputBaris = root.getElementById("fotoBaris");
    root.querySelectorAll("[data-pilih-foto]").forEach((b) =>
      b.addEventListener("click", () => {
        inputBaris.dataset.baris = b.dataset.pilihFoto;
        inputBaris.value = "";
        inputBaris.click();
      })
    );
    inputBaris?.addEventListener("change", (e) => {
      const f = e.target.files[0];
      if (!f) return;
      state.fotoBaris.set(+e.target.dataset.baris, f);
      state.driveBaris.delete(+e.target.dataset.baris);
      validasi();
      render();
    });
    root.querySelectorAll("[data-dosen]").forEach((cb) =>
      cb.addEventListener("change", (e) => {
        state.dosen[+e.target.dataset.dosen].checked = e.target.checked;
      })
    );
    root.getElementById("kirim")?.addEventListener("click", mulaiKirim);
    root.getElementById("stop")?.addEventListener("click", () => {
      state.stopRequested = true;
      root.getElementById("stop").textContent = "Menghentikan…";
    });
    root.getElementById("laporan")?.addEventListener("click", unduhLaporan);
    root.getElementById("reload")?.addEventListener("click", () => location.reload());
  }

  async function onExcel(e) {
    const f = e.target.files[0];
    if (!f) return;
    state.excelName = f.name;
    state.excelBuffer = await f.arrayBuffer();
    state.sheet = "";
    state.sheets = [];
    state.sheetOtomatis = false;
    muatSheet("");
    // file punya beberapa sheet: pilih otomatis sheet yang namanya ada di nama aktivitas
    // (mis. sheet "7 habits" untuk aktivitas "7 Habits ...")
    if (state.sheets.length > 1) {
      const cocok = sheetUntukAktivitas(state.sheets, getInfoAktivitas().nama);
      if (cocok && cocok !== state.sheet) muatSheet(cocok);
      state.sheetOtomatis = !!cocok;
      render();
    }
  }

  function sheetUntukAktivitas(sheets, namaAktivitas) {
    const akt = ` ${norm(namaAktivitas)} `;
    const kandidat = sheets.filter((s) => norm(s) && akt.includes(` ${norm(s)} `));
    // ambil yang namanya paling panjang (paling spesifik)
    return kandidat.sort((a, b) => norm(b).length - norm(a).length)[0] || null;
  }

  function muatSheet(nama) {
    excelMsg = "";
    state.fotoBaris.clear(); // nomor baris berbeda antar sheet
    state.driveBaris.clear();
    try {
      const { rows, sheet, sheets } = bacaExcel(state.excelBuffer, nama || undefined);
      state.rows = rows;
      state.sheet = sheet;
      state.sheets = sheets;
      if (!rows.length) excelMsg = `Tidak ada data di sheet "${sheet}". Isi mulai baris di bawah judul kolom.`;
      validasi();
    } catch (err) {
      state.rows = [];
      excelMsg = err.message;
    }
    render();
    muatFolderDrive();
  }

  function onFoto(e) {
    for (const f of e.target.files) state.files.set(f.name.toLowerCase(), f);
    if (state.rows.length) validasi();
    render();
  }

  // ------------------------------------------------------------ init
  async function init() {
    if (!HALAMAN_LOG) {
      render();
      return;
    }
    if (!getTambahUrl()) {
      fatalMsg =
        "Log aktivitas ini tidak bisa ditambah (sudah difinalisasi atau sedang diverifikasi). " +
        "Buka Log Aktivitas dari aktivitas yang masih berstatus Draft.";
      render();
      return;
    }
    state.existingLogs = getExistingLogs();
    render();
    try {
      const { form } = await fetchFormTambah();
      state.dosen = readDosen(form).map((d) => ({ ...d, checked: true }));
      if (!state.dosen.length) state.dosen = [];
    } catch (err) {
      fatalMsg = err.message;
    }
    render();
  }

  init().catch((err) => {
    // tampilkan error di panel supaya bisa dilaporkan, alih-alih diam saja
    console.error("[AutoKonvert]", err);
    fatalMsg = `Terjadi error (v${HOST.versi}): ${err && err.message ? err.message : err}`;
    try {
      render();
    } catch {
      /* abaikan */
    }
  });
})();
