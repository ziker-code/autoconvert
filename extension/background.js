/*
 * AutoConvert Logbook IPB — service worker (ekstensi Chrome).
 *
 * Mengambil isi folder / file Google Drive untuk content script, karena
 * halaman Student Portal tidak bisa melakukannya sendiri (dibatasi CORS).
 * Hanya melayani pesan dari halaman studentportal.ipb.ac.id. Logikanya ada
 * di drive.js (dipakai bersama dengan userscript iPhone).
 */
"use strict";
importScripts("drive.js");

function bolehDariPengirim(sender) {
  try {
    return sender.id === chrome.runtime.id && new URL(sender.url).origin === "https://studentportal.ipb.ac.id";
  } catch {
    return false;
  }
}

/** Adapter fetch untuk drive.js. Cookie Google ikut, supaya folder yang dibagikan ke akun pengguna bisa dibuka. */
async function ambil(url, { biner }) {
  const res = await fetch(url, { credentials: "include" });
  const hasil = {
    status: res.status,
    url: res.url,
    type: res.headers.get("content-type") || "",
    disposition: res.headers.get("content-disposition") || "",
  };
  const htmlPadaBiner = biner && hasil.type.toLowerCase().startsWith("text/html");
  if (biner && !htmlPadaBiner) hasil.buffer = await res.arrayBuffer();
  else hasil.teks = await res.text();
  return hasil;
}

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (!bolehDariPengirim(sender)) return false;
  if (!msg || !["drive-folder", "drive-file"].includes(msg.type)) return false;
  AutoConvertDrive.proses(ambil, msg).then(sendResponse, (e) => sendResponse({ ok: false, error: e?.message || String(e) }));
  return true; // jawab async
});
