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
                  ? `tidak bisa dibuka (akses dibatasi). Login ke Google di browser ini dengan akun yang punya akses lalu muat ulang halaman, minta pemilik folder mengubah akses ke "Siapa saja yang memiliki link", atau klik "Buka folder", unduh fotonya, lalu pilih lewat "${LOKAL}"`
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
      * { box-sizing: border-box; font-family: "Poppins", "Segoe UI", system-ui, sans-serif; }
      .fab { position: fixed; right: 22px; bottom: 22px; z-index: 2147483000; border: 0; cursor: pointer;
        background: #0b5394; color: #fff; padding: 12px 18px; border-radius: 999px; font-size: 14px; font-weight: 600;
        box-shadow: 0 6px 20px rgba(11,83,148,.35); display: flex; align-items: center; gap: 8px; }
      .fab:hover { background: #094479; }
      .panel { position: fixed; right: 22px; bottom: 22px; z-index: 2147483001; width: min(760px, calc(100vw - 44px));
        max-height: calc(100vh - 44px); background: #fff; color: #1f2937; border-radius: 14px; display: none;
        flex-direction: column; box-shadow: 0 18px 50px rgba(0,0,0,.28); overflow: hidden; font-size: 13px; }
      .panel.open { display: flex; }
      .head { background: #0b5394; color: #fff; padding: 14px 18px; display: flex; align-items: center; justify-content: space-between; }
      .head h2 { margin: 0; font-size: 16px; }
      .head small { opacity: .85; display: block; font-size: 12px; margin-top: 2px; }
      .x { background: transparent; border: 0; color: #fff; font-size: 22px; cursor: pointer; line-height: 1; }
      .body { padding: 16px 18px; overflow: auto; display: flex; flex-direction: column; gap: 14px; }
      .step { border: 1px solid #e5e7eb; border-radius: 10px; padding: 12px 14px; }
      .step h3 { margin: 0 0 8px; font-size: 13px; display: flex; gap: 8px; align-items: center; }
      .num { background: #0b5394; color: #fff; width: 22px; height: 22px; border-radius: 50%; display: inline-flex;
        align-items: center; justify-content: center; font-size: 12px; flex: none; }
      .muted { color: #6b7280; }
      label.file { display: inline-flex; align-items: center; gap: 8px; background: #eef4fb; border: 1px dashed #0b5394;
        color: #0b5394; padding: 8px 12px; border-radius: 8px; cursor: pointer; font-weight: 600; }
      label.file input { display: none; }
      .row { display: flex; flex-wrap: wrap; gap: 10px; align-items: center; }
      .dosen label { display: flex; gap: 8px; align-items: center; padding: 3px 0; }
      .alert { padding: 10px 12px; border-radius: 8px; background: #fef2f2; color: #991b1b; }
      .ok { background: #ecfdf5; color: #065f46; }
      .info { background: #eff6ff; color: #1e3a8a; }
      .summary { display: flex; gap: 8px; flex-wrap: wrap; }
      .chip { padding: 3px 10px; border-radius: 999px; font-size: 12px; font-weight: 600; background: #f3f4f6; }
      .c-siap { background: #e0f2fe; color: #075985; } .c-error, .c-gagal { background: #fee2e2; color: #991b1b; }
      .c-berhasil { background: #dcfce7; color: #166534; } .c-duplikat { background: #f3f4f6; color: #4b5563; }
      .c-mengirim, .c-memuat { background: #fef9c3; color: #854d0e; }
      td.bukti { max-width: 170px; word-break: break-word; }
      .galeri { position: fixed; inset: 0; z-index: 2147483002; background: rgba(15,23,42,.55); display: none;
        align-items: center; justify-content: center; padding: 22px; }
      .galeri.open { display: flex; }
      .gbox { background: #fff; color: #1f2937; border-radius: 14px; width: min(720px, 100%); max-height: 100%;
        display: flex; flex-direction: column; overflow: hidden; font-size: 13px; }
      .ghead { display: flex; justify-content: space-between; align-items: center; padding: 12px 16px; border-bottom: 1px solid #e5e7eb; }
      .x.dark { color: #374151; }
      .ggrid { display: grid; grid-template-columns: repeat(auto-fill, minmax(130px, 1fr)); gap: 10px; padding: 14px 16px; overflow: auto; }
      .gitem { border: 2px solid #e5e7eb; border-radius: 10px; padding: 6px; background: #fff; cursor: pointer; text-align: left; }
      .gitem:hover { border-color: #93c5fd; } .gitem.aktif { border-color: #0b5394; box-shadow: 0 0 0 2px #bfdbfe; }
      .gthumb { height: 100px; background: #f3f4f6; border-radius: 6px; display: flex; align-items: center; justify-content: center;
        overflow: hidden; color: #6b7280; font-size: 12px; }
      .gthumb img { width: 100%; height: 100%; object-fit: cover; }
      .gname { margin-top: 4px; font-size: 11px; word-break: break-all; color: #374151; }
      table { width: 100%; border-collapse: collapse; font-size: 12px; }
      th, td { text-align: left; padding: 6px 8px; border-bottom: 1px solid #f0f0f0; vertical-align: top; }
      th { background: #f9fafb; position: sticky; top: 0; }
      .tablewrap { max-height: 300px; overflow: auto; border: 1px solid #e5e7eb; border-radius: 8px; }
      .ket { max-width: 240px; }
      .err { color: #b91c1c; font-size: 11px; } .warn { color: #a16207; font-size: 11px; }
      .foot { padding: 12px 18px; border-top: 1px solid #e5e7eb; display: flex; gap: 10px; justify-content: flex-end; flex-wrap: wrap; align-items: center; }
      button.btn { border: 0; border-radius: 8px; padding: 9px 16px; font-weight: 600; cursor: pointer; font-size: 13px; }
      .primary { background: #0b5394; color: #fff; } .primary:disabled { background: #9ca3af; cursor: not-allowed; }
      .secondary { background: #f3f4f6; color: #111827; } .danger { background: #dc2626; color: #fff; }
      a.mini { display: inline-block; text-decoration: none; }
      a.logitem { display: flex; justify-content: space-between; align-items: center; gap: 10px; padding: 10px 12px;
        border: 1px solid #e5e7eb; border-radius: 8px; margin-top: 8px; color: #0b5394; text-decoration: none; }
      a.logitem:hover { border-color: #0b5394; background: #eef4fb; }
      button.mini, a.mini { margin-top: 4px; border: 1px solid #0b5394; background: #fff; color: #0b5394; border-radius: 6px;
        padding: 2px 8px; font-size: 11px; font-weight: 600; cursor: pointer; }
      select { padding: 5px 8px; border-radius: 6px; border: 1px solid #d1d5db; font-size: 13px; }
      .progress { flex: 1; min-width: 160px; height: 8px; background: #e5e7eb; border-radius: 99px; overflow: hidden; }
      .progress > div { height: 100%; background: #16a34a; transition: width .3s; }

      /* ---- layar HP: panel penuh layar, tabel jadi kartu, tombol lebih besar ---- */
      @media (max-width: 640px) {
        .fab { right: 12px; left: 12px; bottom: calc(12px + env(safe-area-inset-bottom)); justify-content: center; padding: 14px 18px; font-size: 15px; }
        .panel { inset: 0; right: 0; bottom: 0; width: 100%; max-height: none; height: 100%; border-radius: 0; font-size: 14px; }
        .head { padding: calc(12px + env(safe-area-inset-top)) 14px 12px; }
        .x { font-size: 30px; padding: 0 6px; }
        .body { padding: 12px; gap: 12px; -webkit-overflow-scrolling: touch; }
        .step { padding: 12px; }
        label.file { width: 100%; justify-content: center; padding: 12px; font-size: 14px; }
        select { width: 100%; font-size: 16px; padding: 8px; }
        .dosen label { padding: 6px 0; }
        .dosen input { width: 20px; height: 20px; }
        .tablewrap { max-height: none; border: 0; overflow: visible; }
        table, tbody, tr, td { display: block; width: 100%; }
        thead { display: none; }
        tr { border: 1px solid #e5e7eb; border-radius: 10px; padding: 10px 12px; margin-bottom: 10px; }
        td { border: 0; padding: 3px 0; max-width: none !important; }
        td:first-child { font-size: 11px; color: #6b7280; }
        td:first-child::before { content: "Baris "; }
        td.bukti { padding-top: 6px; }
        button.mini, a.mini { padding: 7px 12px; font-size: 13px; margin: 6px 6px 0 0; }
        .foot { padding: 10px 12px calc(10px + env(safe-area-inset-bottom)); }
        .foot button.btn { flex: 1; padding: 13px 12px; font-size: 15px; }
        .progress { flex-basis: 100%; }
        .galeri { padding: 0; }
        .gbox { width: 100%; height: 100%; border-radius: 0; }
        .ghead { padding-top: calc(12px + env(safe-area-inset-top)); }
        .ggrid { grid-template-columns: repeat(2, 1fr); }
        .gthumb { height: 130px; }
      }
    </style>
    <button class="fab" id="fab">⚡ AutoKonvert Logbook</button>
    <div class="panel" id="panel">
      <div class="head">
        <div><h2>⚡ AutoKonvert Logbook <small style="display:inline;opacity:.7">v${esc(HOST.versi)}${HOST.jenis === "userscript" ? " · HP" : ""}</small></h2><small id="aktivitas"></small></div>
        <button class="x" id="close" title="Tutup">×</button>
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
    if (d?.file) isi = `📷 ${esc(d.file.name)}<div class="muted">dari ${LOKAL_KECIL}</div>`;
    else if (d?.drive) isi = `☁️ ${esc(d.drive.name)}<div class="muted">dari Google Drive</div>`;
    else if (r.status === "memuat") isi = `☁️ <span class="muted">membuka folder…</span>`;
    else if (r.raw.bukti) isi = `<span class="muted">${esc(r.raw.bukti)}</span>`;
    else isi = '<span class="muted">—</span>';
    if (state.running || r.status === "berhasil" || r.status === "duplikat") return isi;
    const fid = folderBaris(r);
    const st = fid && state.drive.get(fid)?.status;
    const galeri = st === "ok" ? `<button class="mini" data-galeri="${r.baris}">Galeri</button> ` : "";
    // link hanya ditampilkan kalau memang link Google Drive yang valid
    const buka =
      st === "gagal" && !d?.file
        ? `<a class="mini" href="${esc(r.raw.buktiLink)}" target="_blank" rel="noopener noreferrer">Buka folder ↗</a> `
        : "";
    return `${isi}<div>${galeri}${buka}<button class="mini" data-pilih-foto="${r.baris}">${LOKAL}</button></div>`;
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
          <button class="x dark" id="gtutup" title="Tutup">×</button></div>
        <div class="ggrid">${media
          .map(
            (f) => `<button class="gitem ${f.id === dipilih ? "aktif" : ""}" data-id="${esc(f.id)}">
              <div class="gthumb" data-thumb="${esc(f.id)}">${f.jenis === "pdf" ? "📄 PDF" : "⏳"}</div>
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
      if (el) el.innerHTML = url ? `<img src="${url}" alt="">` : "🖼️";
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
      dariDrive && `☁️ ${dariDrive} dari Google Drive`,
      dariKomputer && `📷 ${dariKomputer} dari ${LOKAL_KECIL}`,
      memuat && `⏳ ${memuat} sedang dibuka`,
      tanpa && `${tanpa} tanpa bukti`,
    ]
      .filter(Boolean)
      .join(" · ");
    const tombolFoto = `<label class="file">📷 Pilih foto dari perangkat<input type="file" id="foto" multiple accept="${ACCEPT_BUKTI}" ${state.running ? "disabled" : ""}></label>`;
    const ulang = driveGagal
      ? `<div class="alert info" style="margin-top:8px">${driveGagal} folder Google Drive tidak bisa dibuka karena aksesnya dibatasi.
          Login ke Google di browser ini dengan akun yang punya akses, lalu
          <button class="mini" id="ulangDrive" ${state.running ? "disabled" : ""}>Coba buka ulang</button></div>`
      : "";

    if (!kurang.length) {
      return `<div class="step">
        <h3><span class="num">2</span> Foto bukti ${memuat ? "" : "✅"}</h3>
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
      : `<div class="step"><h3>Cara mulai</h3><ol style="margin:0;padding-left:18px">
           <li>Buka <b>Kemahasiswaan → Aktivitas</b>.</li>
           <li>Ketuk tombol <b>Log</b> (ikon daftar) pada aktivitas berstatus <b>Draft</b>.</li>
           <li>Ketuk lagi tombol <b>⚡ AutoKonvert</b> di halaman Log itu.</li></ol>
           <a class="logitem" href="${PORTAL_ORIGIN}/Kegiatan/AktivitasKampusMerdeka" style="margin-top:10px"><b>Buka daftar Aktivitas →</b></a></div>`;
    return `<div class="alert info">AutoKonvert aktif ✅ (v${esc(HOST.versi)}). Pengisian logbook dilakukan di halaman <b>Log</b> sebuah aktivitas.</div>${daftar}`;
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
      body.innerHTML = `<div class="alert">${esc(fatalMsg)}</div>`;
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
          ...(r.errors || []).map((e) => `<div class="err">✖ ${esc(e)}</div>`),
          ...(r.warnings || []).map((w) => `<div class="warn">⚠ ${esc(w)}</div>`),
          r.pesan ? `<div class="${r.status === "gagal" ? "err" : "muted"}">${esc(r.pesan)}</div>` : "",
        ].join("");
        return `<tr>
          <td>${r.baris}</td>
          <td>${esc(tampilTanggal(r))}<br><span class="muted">${esc(tampilJam(r.raw.mulai))}–${esc(tampilJam(r.raw.selesai))}</span></td>
          <td class="ket">${esc(r.raw.keterangan)}<div class="muted">${esc(tampilTempat(r))}</div></td>
          <td class="bukti">${selBukti(r)}</td>
          <td><span class="chip c-${r.status}">${esc(r.status)}</span>${catatan}</td>
        </tr>`;
      })
      .join("");

    body.innerHTML = `
      <div class="step">
        <h3><span class="num">1</span> Pilih file Excel logbook</h3>
        <div class="row">
          <label class="file">📄 Pilih Excel<input type="file" id="excel" accept="${ACCEPT_EXCEL}" ${state.running ? "disabled" : ""}></label>
          <span class="muted">${state.excelName ? esc(state.excelName) + ` · ${state.rows.length} baris` : "Gunakan template dari website AutoKonvert"}</span>
        </div>
        ${
          state.sheets.length > 1
            ? `<div class="row" style="margin-top:8px"><span>Sheet:</span>
                <select id="sheet" ${state.running ? "disabled" : ""}>${state.sheets
                  .map((s) => `<option value="${esc(s)}" ${s === state.sheet ? "selected" : ""}>${esc(s)}</option>`)
                  .join("")}</select>
                </div>
               <div class="${state.sheetOtomatis ? "alert ok" : "alert info"}" style="margin-top:8px">${
                 state.sheetOtomatis
                   ? `Sheet <b>${esc(state.sheet)}</b> dipilih otomatis karena cocok dengan nama aktivitas ini.`
                   : `File ini punya ${state.sheets.length} sheet. Pastikan sheet <b>${esc(state.sheet)}</b> memang untuk aktivitas <b>${esc(getInfoAktivitas().nama)}</b>. Sheet lain dikirim dari halaman Log aktivitasnya sendiri.`
               }</div>`
            : ""
        }
        <input type="file" id="fotoBaris" accept="${ACCEPT_BUKTI}" style="display:none">
        ${excelMsg ? `<div class="alert" style="margin-top:8px">${esc(excelMsg)}</div>` : ""}
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
        <div class="summary" style="margin-bottom:8px">
          <span class="chip c-siap">Siap: ${c.siap}</span>
          <span class="chip c-error">Perlu diperbaiki: ${c.error}</span>
          <span class="chip c-duplikat">Sudah ada: ${c.duplikat}</span>
          <span class="chip c-berhasil">Berhasil: ${c.berhasil}</span>
          <span class="chip c-gagal">Gagal: ${c.gagal}</span>
          ${c.memuat ? `<span class="chip c-memuat">Membuka Drive: ${c.memuat}</span>` : ""}
          <span class="chip">Total durasi: ${Math.floor(totalMenit / 60)} jam ${totalMenit % 60} menit</span>
        </div>
        ${c.error ? `<div class="alert" style="margin-bottom:8px">Ada ${c.error} baris yang perlu diperbaiki. Baris ini tidak akan dikirim sampai diperbaiki: pilih fotonya, atau perbaiki Excel lalu pilih ulang filenya.</div>` : ""}
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
         <button class="btn danger" id="stop">Hentikan</button>`
      : `${c.berhasil ? `<span class="chip c-berhasil">${c.berhasil} log terkirim</span>` : ""}
         ${state.rows.length ? `<button class="btn secondary" id="laporan">Unduh laporan</button>` : ""}
         ${c.berhasil ? `<button class="btn secondary" id="reload">Muat ulang halaman</button>` : ""}
         <button class="btn primary" id="kirim" ${siapKirim ? "" : "disabled"}>${c.gagal && !c.siap ? "Kirim ulang yang gagal" : `Kirim ${siapKirim} log`}</button>`;

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
