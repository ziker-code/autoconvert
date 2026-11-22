/*
 * AutoKonvert Logbook IPB — akses Google Drive (dipakai bersama).
 *
 * Dipakai oleh background.js (ekstensi Chrome, lewat fetch) dan oleh
 * userscript iPhone (lewat GM.xmlHttpRequest). Pemanggil menyediakan fungsi
 *   ambil(url, { biner }) -> Promise<{ status, url, type, disposition, teks?, buffer? }>
 * URL selalu dibentuk dari ID Drive yang sudah divalidasi, bukan URL bebas.
 */
(() => {
  "use strict";
  const ID_RE = /^[\w-]{10,200}$/;
  const BATAS_BYTES = 40 * 1024 * 1024;

  function keBase64(buf) {
    const bytes = new Uint8Array(buf);
    let bin = "";
    for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(bin);
  }

  function namaDariDisposition(cd) {
    cd = cd || "";
    const star = cd.match(/filename\*\s*=\s*(?:UTF-8'')?([^;]+)/i);
    if (star) {
      try {
        return decodeURIComponent(star[1].trim().replace(/^"|"$/g, ""));
      } catch {
        /* abaikan */
      }
    }
    const biasa = cd.match(/filename\s*=\s*"?([^";]+)"?/i);
    return biasa ? biasa[1].trim() : "";
  }

  async function ambilFolder(ambil, id) {
    if (!ID_RE.test(String(id))) return { ok: false, error: "ID tidak valid" };
    const res = await ambil(`https://drive.google.com/embeddedfolderview?id=${id}`, { biner: false });
    const html = res.teks || "";
    if (res.status !== 200 || !html.includes("flip-entry")) {
      const login =
        [401, 403].includes(res.status) || /accounts\.google\.com|ServiceLogin/i.test((res.url || "") + html.slice(0, 3000));
      return { ok: false, error: login ? "privat" : res.status === 404 ? "tidak ditemukan" : `HTTP ${res.status}` };
    }
    return { ok: true, html };
  }

  /** mode: "thumb" (kecil, untuk galeri) | "gambar" (resolusi besar) | "unduh" (file asli, mis. PDF) */
  async function ambilFile(ambil, id, mode) {
    if (!ID_RE.test(String(id))) return { ok: false, error: "ID tidak valid" };
    const url =
      mode === "thumb"
        ? `https://drive.google.com/thumbnail?id=${id}&sz=w320`
        : mode === "gambar"
          ? `https://drive.google.com/thumbnail?id=${id}&sz=w2000`
          : `https://drive.usercontent.google.com/download?id=${id}&export=download`;
    const res = await ambil(url, { biner: true });
    const type = String(res.type || "").split(";")[0].trim().toLowerCase();
    if (res.status !== 200 || type.startsWith("text/html") || !res.buffer) {
      return { ok: false, error: res.status === 404 ? "tidak ditemukan" : "tidak bisa diakses (cek izin berbagi)" };
    }
    if (res.buffer.byteLength > BATAS_BYTES) return { ok: false, error: "file terlalu besar" };
    return { ok: true, type, name: namaDariDisposition(res.disposition), base64: keBase64(res.buffer) };
  }

  /** Jalankan pesan {type: "drive-folder"|"drive-file", id, mode}. */
  function proses(ambil, pesan) {
    if (!pesan || !ID_RE.test(String(pesan.id || ""))) return Promise.resolve({ ok: false, error: "permintaan tidak valid" });
    if (pesan.type === "drive-folder") return ambilFolder(ambil, pesan.id);
    if (pesan.type === "drive-file") return ambilFile(ambil, pesan.id, pesan.mode);
    return Promise.resolve({ ok: false, error: "permintaan tidak dikenal" });
  }

  globalThis.AutoKonvertDrive = { ID_RE, proses, ambilFolder, ambilFile, keBase64, namaDariDisposition };
})();
