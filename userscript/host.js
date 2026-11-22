/*
 * AutoKonvert Logbook IPB — lapisan host untuk userscript (iPhone/Safari: aplikasi
 * "Userscripts" atau Tampermonkey). Menggantikan peran background.js di ekstensi:
 * permintaan ke Google Drive dikirim lewat GM.xmlHttpRequest (tidak terkena CORS).
 */
(() => {
  "use strict";

  const gmXhr =
    (typeof GM !== "undefined" && GM && GM.xmlHttpRequest && GM.xmlHttpRequest.bind(GM)) ||
    (typeof GM_xmlhttpRequest !== "undefined" && GM_xmlhttpRequest) ||
    null;

  function header(raw, nama) {
    const m = String(raw || "").match(new RegExp("^" + nama + ":\\s*(.*)$", "im"));
    return m ? m[1].trim() : "";
  }

  function kirim(url, responseType) {
    return new Promise((resolve, reject) => {
      let selesai = false;
      const ok = (r) => {
        if (!selesai) {
          selesai = true;
          resolve(r);
        }
      };
      const gagal = (e) => {
        if (!selesai) {
          selesai = true;
          reject(new Error((e && (e.error || e.statusText)) || "koneksi ke Google Drive gagal"));
        }
      };
      const hasil = gmXhr({
        method: "GET",
        url,
        responseType,
        anonymous: false,
        timeout: 60000,
        onload: ok,
        onerror: gagal,
        ontimeout: () => gagal({ error: "waktu habis" }),
      });
      // sebagian pengelola userscript mengembalikan Promise alih-alih memanggil onload
      if (hasil && typeof hasil.then === "function") hasil.then(ok, gagal);
    });
  }

  async function keBuffer(resp) {
    const r = resp.response;
    if (r instanceof ArrayBuffer) return r;
    if (r && typeof r.arrayBuffer === "function") return r.arrayBuffer(); // Blob
    if (r && r.buffer instanceof ArrayBuffer) return r.buffer; // Uint8Array
    return null;
  }

  /** Adapter untuk AutoKonvertDrive: sama bentuknya dengan versi fetch di background.js */
  async function ambil(url, { biner }) {
    if (!gmXhr) throw new Error("Pengelola userscript tidak mengizinkan akses Google Drive (GM.xmlHttpRequest).");
    let resp = await kirim(url, biner ? "arraybuffer" : "text");
    const hasil = {
      status: resp.status,
      url: resp.finalUrl || url,
      type: header(resp.responseHeaders, "content-type"),
      disposition: header(resp.responseHeaders, "content-disposition"),
    };
    if (biner && !hasil.type.toLowerCase().startsWith("text/html")) {
      let buf = await keBuffer(resp);
      if (!buf && resp.status === 200) {
        // cadangan: sebagian versi hanya mendukung responseType "blob"
        resp = await kirim(url, "blob");
        buf = await keBuffer(resp);
      }
      hasil.buffer = buf;
    } else {
      hasil.teks = typeof resp.responseText === "string" ? resp.responseText : String(resp.response || "");
    }
    return hasil;
  }

  globalThis.AutoKonvertHost = {
    jenis: "userscript",
    versi: "__VERSI__",
    minta: (pesan) => globalThis.AutoKonvertDrive.proses(ambil, pesan),
  };
})();
