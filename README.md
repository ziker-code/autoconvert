# AutoKonvert Logbook IPB

🌐 **Website:** https://autokonvert.vercel.app/ (template Excel, ekstensi Chrome, dan versi iPhone)

Ekstensi Chrome dan website panduan untuk mengisi **Log Aktivitas** Student Portal IPB secara otomatis dari file Excel, lengkap dengan foto bukti.

## Struktur

```
extension/          Ekstensi Chrome (Manifest V3), bisa langsung "Load unpacked"
  manifest.json
  parser.js         Logika baca Excel & validasi (murni, tanpa DOM, ada tesnya)
  content.js        Panel ⚡ di halaman Log Aktivitas + pengiriman ke portal
  background.js     Service worker: baca folder & unduh foto Google Drive (lewati CORS)
  popup.html/js     Popup petunjuk singkat
  vendor/           SheetJS (disalin otomatis saat build)
website/            Website statis: panduan + download
  index.html
  assets/           template Excel, zip ekstensi, ikon (dibuat oleh build)
userscript/
  host.js           Pengganti background.js untuk versi iPhone (GM.xmlHttpRequest)
tools/
  build.mjs         Buat ikon, template Excel, dan zip ekstensi
  test-parser.mjs   Tes parser
  test-userscript.mjs  Tes lapisan Drive versi userscript (butuh internet)
  make-harness.mjs  Halaman uji tiruan Log Aktivitas (tools/harness/, tanpa portal asli)
```

## Perintah

```powershell
npm install
npm run build   # perbarui vendor, ikon, template, dan zip
npm test        # tes parser
```

Website di-hosting di **Vercel** (lihat `vercel.json`): setiap push ke `main` otomatis menjalankan `npm run build && npm test` lalu menerbitkan folder `website/`.
Jalankan `npm run build` setiap kali isi `extension/` berubah supaya zip di website ikut diperbarui.

## Cara kerja teknis

Halaman log: `https://studentportal.ipb.ac.id/Kegiatan/LogAktivitasKampusMerdeka/Index/{id}`

Untuk setiap baris Excel, ekstensi:

1. Mengambil form modal `GET /Kegiatan/LogAktivitasKampusMerdeka/Tambah?aktivitasId=…`, yang memberi
   `__RequestVerificationToken`, `AktivitasId`, dan daftar dosen yang masih berlaku.
2. Mengisi field form tersebut lalu mengirim `POST /Kegiatan/LogAktivitasKampusMerdeka/Tambah` (multipart).
   Responsnya JSON `{ status, message, url }`.

| Field portal | Isi |
|---|---|
| `Waktu` | `dd/mm/yyyy` |
| `Tmw`, `Tsw` | `HH:mm` (jam mulai/selesai) |
| `JenisLogbookKegiatanKampusMerdekaId` | 1 = Pembimbingan, 2 = Ujian, 3 = Kegiatan |
| `ListDosenPembimbing[i].Key.PembimbingId` / `.Value` | dosen penggerak (checkbox) |
| `IsLuring` | `"true"` = Offline, `"false"` = Online, `""` = Hybrid |
| `Lokasi`, `Keterangan` | teks |
| `File` | bukti `.pdf/.png/.jpg/.jpeg` |

Pengiriman diberi jeda 1,5 detik per baris. Log yang tanggal dan jam mulainya sudah ada di tabel dilewati.

## Foto dari Google Drive

Jika sel kolom bukti berisi link Drive (hyperlink sel, rumus HYPERLINK, atau teks URL):

- **Folder**: daftar isi dibaca dari `drive.google.com/embeddedfolderview?id=…`. Satu foto dipilih otomatis per baris (baris yang memakai folder sama mendapat foto berbeda), dan bisa diganti lewat Galeri.
- **File**: diunduh langsung.
- Gambar diambil lewat `drive.google.com/thumbnail?id=…&sz=w2000`, PDF lewat `drive.usercontent.google.com/download`.
- Permintaan dikirim dari service worker dengan cookie Google pengguna, jadi folder yang hanya bisa diakses akun tertentu tetap bisa dibuka selama akun itu login di Chrome.

## Versi iPhone (userscript)

`npm run build` juga menghasilkan `website/assets/autokonvert.user.js` (SheetJS mini + parser.js + drive.js + userscript/host.js + content.js).
Di iPhone dipasang lewat Safari + aplikasi **Userscripts**. Saat di-build di Vercel, `@updateURL` otomatis memakai domain produksi Vercel (atau atur `SITE_URL` sendiri).

## Lisensi

[MIT](LICENSE) © 2026 ziker-code. Ekstensi dan userscript juga berisi SheetJS Community Edition (Apache-2.0); lihat [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).

AutoKonvert adalah alat bantu buatan mahasiswa, bukan aplikasi resmi IPB University.
