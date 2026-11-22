// Server uji lokal: menyajikan tools/harness/index.html di alamat mirip Student Portal,
// supaya panel bisa dilihat dalam mode Log maupun mode bantuan. Tidak menghubungi portal asli.
//   node tools/harness-server.mjs  ->  http://localhost:4173/Kegiatan/LogAktivitasKampusMerdeka/Index/uji
import http from "node:http";
import fs from "node:fs";
const PORT = Number(process.env.PORT || 4173);
http
  .createServer((req, res) => {
    const html = fs.readFileSync(new URL("./harness/index.html", import.meta.url));
    res.writeHead(200, { "content-type": "text/html; charset=utf-8", "cache-control": "no-store" });
    res.end(html);
  })
  .listen(PORT, () => console.log(`Harness: http://localhost:${PORT}/Kegiatan/LogAktivitasKampusMerdeka/Index/uji`));
