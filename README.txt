AT FARM PWA V10.50 - LOGO LOGIN + ICON MOBILE

PERUBAHAN:
1. Ikon aplikasi PWA di layar utama HP menggunakan LOGO AT baru.
2. Logo baru dibuat dengan safe-area supaya tidak terpotong Android.
3. Logo dashboard/menu lain TIDAK diubah.
4. Disertakan file Index_APPS_SCRIPT_LOGIN_LOGO_BARU.html untuk mengganti Index.html
   di Google Apps Script agar gambar sayur pada halaman LOGIN berubah menjadi LOGO AT baru.

CARA PWA:
- Ganti isi folder repository at-farm-pwa dengan file PWA dari paket ini.
- Commit lalu Push origin di GitHub Desktop.
- Tunggu GitHub Pages memperbarui versi.

CARA LOGIN:
- Di Google Apps Script, backup Index.html lama.
- Gunakan isi Index_APPS_SCRIPT_LOGIN_LOGO_BARU.html sebagai Index.html.
- Deploy / Kelola penerapan -> buat versi baru / perbarui penerapan.
- Fungsi dashboard dan logo lain tidak sengaja diubah; hanya elemen logo login pada file dasar ini.

CATATAN:
Jika aplikasi HP masih menampilkan ikon lama, hapus shortcut/PWA lama lalu pasang kembali
setelah GitHub Pages selesai memperbarui cache.
