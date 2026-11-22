// Tes lapisan userscript (drive.js + host.js) dengan GM.xmlHttpRequest tiruan berbasis fetch.
// Butuh internet. Jalankan: node tools/test-userscript.mjs
import fs from "node:fs";
import vm from "node:vm";
import assert from "node:assert/strict";
import * as XLSX from "xlsx";

// ID folder Google Drive publik untuk dites (berisi foto). Contoh:
//   node tools/test-userscript.mjs 1AbC...xyz
const FOLDER_PUBLIK = process.argv[2] || process.env.TEST_DRIVE_FOLDER;
if (!FOLDER_PUBLIK) {
  console.log("Lewati: berikan ID folder Drive publik sebagai argumen (node tools/test-userscript.mjs <id>).");
  process.exit(0);
}

function buatGM(responseTypeDidukung) {
  return {
    xmlHttpRequest(opt) {
      (async () => {
        const res = await fetch(opt.url, { headers: { "User-Agent": "Mozilla/5.0" } });
        const buf = await res.arrayBuffer();
        const headers = [...res.headers].map(([k, v]) => `${k}: ${v}`).join("\r\n");
        let response;
        const rt = responseTypeDidukung.includes(opt.responseType) ? opt.responseType : "text";
        if (rt === "arraybuffer") response = buf;
        else if (rt === "blob") response = new Blob([buf]);
        else response = new TextDecoder().decode(buf);
        opt.onload({ status: res.status, finalUrl: res.url, responseHeaders: headers, response, responseText: rt === "text" ? response : undefined });
      })().catch((e) => opt.onerror({ error: e.message }));
    },
  };
}

for (const dukung of [["arraybuffer", "blob", "text"], ["blob", "text"]]) {
  const ctx = vm.createContext({ GM: buatGM(dukung), btoa, atob, TextDecoder, Blob, ArrayBuffer, Uint8Array, console });
  ctx.globalThis = ctx;
  for (const f of ["extension/parser.js", "extension/drive.js", "userscript/host.js"]) vm.runInContext(fs.readFileSync(f, "utf8"), ctx);
  const H = ctx.AutoConvertHost;
  assert.equal(H.jenis, "userscript");

  const folder = await H.minta({ type: "drive-folder", id: FOLDER_PUBLIK });
  assert.ok(folder.ok, "folder publik harus bisa dibuka: " + folder.error);
  const files = ctx.AutoConvertParser.parseDriveFolder(folder.html);
  const gambar = files.filter((f) => f.jenis === "gambar");
  assert.ok(gambar.length > 0);

  const foto = await H.minta({ type: "drive-file", id: gambar[0].id, mode: "thumb" });
  assert.ok(foto.ok, foto.error);
  assert.match(foto.type, /^image\//);
  const bytes = Buffer.from(foto.base64, "base64");
  assert.ok(bytes.length > 1000);

  const tolak = await H.minta({ type: "drive-file", id: "../../etc", mode: "thumb" });
  assert.equal(tolak.ok, false);
  console.log(`  ✓ responseType [${dukung}]: folder ${files.length} file, thumbnail ${foto.type} ${(bytes.length / 1024).toFixed(0)} KB`);
}

// SheetJS versi mini (dipakai userscript) tetap bisa membaca template
const mini = vm.createContext({ console });
vm.runInContext(fs.readFileSync("node_modules/xlsx/dist/xlsx.mini.min.js", "utf8"), mini);
vm.runInContext(fs.readFileSync("extension/parser.js", "utf8").replace("globalThis.AutoConvertParser", "this.AutoConvertParser"), mini);
const buf = new Uint8Array(fs.readFileSync("website/assets/template-logbook-autoconvert.xlsx"));
const res = mini.AutoConvertParser.bacaExcel(mini.XLSX, buf);
assert.equal(res.sheet, "Logbook");
console.log("  ✓ xlsx.mini membaca template");
console.log("\nTes userscript lulus");
