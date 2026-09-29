AT FARM V10.51 - WAJIB LOGIN ULANG

Perubahan:
- Saat aplikasi / halaman dibuka ulang, selalu kembali ke halaman Login.
- Username dan password harus diisi kembali.
- Token login tidak lagi disimpan permanen di localStorage.
- Token hanya hidup selama halaman aplikasi yang sedang berjalan.
- Token lama dari versi sebelumnya dibersihkan.
- Logout tetap bekerja.
- Logo login baru dan ikon PWA dari V10.50 tetap dipertahankan.
- Dashboard, Kasir, data, dan fitur lain tidak dimaksudkan untuk diubah.

Pemasangan:
1. Google Apps Script:
   ganti seluruh isi Index.html dengan Index_APPS_SCRIPT_V10_51_WAJIB_LOGIN_ULANG.html
2. Save.
3. Deploy / Manage deployments -> Edit -> New version -> Deploy.
4. PWA GitHub tidak wajib diubah untuk mekanisme login ini bila PWA V10.50 sudah terpasang.
