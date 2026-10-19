# AT FARM PWA — PC Baru (V11.05)

Folder ini siap dijadikan repository GitHub **at-farm-pwa** di PC baru.

## Struktur
- `index.html`, `config.js`, `sw.js`, `manifest.webmanifest`, ikon: **PWA GitHub Pages**.
- `APPS_SCRIPT/Index.html` dan `APPS_SCRIPT/Code.gs`: **sumber aplikasi Google Apps Script**, bukan halaman GitHub Pages.

## Cara pakai
1. Ekstrak ZIP ke `Documents\GitHub\at-farm-pwa` (jangan sampai folder bersarang ganda).
2. **Cadangkan repository lama** sebelum mengganti apa pun. Bila repo `at-farm-pwa` sudah ada, salin hanya file PWA dari paket ini ke folder repository yang sudah ada, bukan folder `.git`.
3. Buka GitHub Desktop pada repo `at-farm-pwa`, lihat Changes, lalu Commit dan Push origin.
4. Untuk fitur Export JPG dan tata letak Data Terbaru, buka proyek Google Apps Script yang benar dan perbarui file `Index.html` menggunakan `APPS_SCRIPT/Index.html`. `Code.gs` hanya dipakai jika perlu dan sesudah mencadangkan versi aktif.
5. Di Apps Script: Deploy > Manage deployments > Edit > New version > Deploy.
6. Pastikan `config.js` berisi URL `/exec` dari deployment **yang benar dan aktif**. URL bawaan berasal dari arsip V11.05 dan belum diverifikasi apakah masih aktif.
7. Refresh paksa PWA / hapus cache aplikasi jika tampilan lama masih tersimpan.

**Penting:** GitHub Pages hanya membungkus aplikasi Apps Script melalui iframe; mengubah GitHub saja tidak menambah tombol Export JPG di halaman Apps Script.


## V11.06 - Print Bersih + Total Keseluruhan
- Navigasi/dashboard tidak tercetak pada PDF laporan.
- Baris TOTAL KESELURUHAN di bagian akhir tabel pada semua fungsi cetak laporan.
- Uang Masuk/Uang Keluar dijumlahkan; Saldo menampilkan saldo terakhir, bukan penjumlahan saldo berjalan.
- Export JPG laporan menyertakan baris total.
- Pasang `APPS_SCRIPT/Index.html` pada Google Apps Script, simpan, lalu Deploy > Manage deployments > Edit > New version.
- `Code.gs` dan berkas PWA lain tetap sama.
