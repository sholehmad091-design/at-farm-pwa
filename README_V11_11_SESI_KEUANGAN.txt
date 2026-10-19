PASAR AT FARM V11.11 - TOTAL KEUANGAN PER SESI

Perubahan hanya pada APPS_SCRIPT/Code.gs.
- Transaksi baru diberi metadata sesi di sheet AT_SESI_KEUANGAN.
- Transaksi historis tidak diubah dan tetap tampil seperti sebelumnya.
- Pada Total Keuangan, transaksi baru menampilkan nama sumber pada kolom keterangan.
- Beberapa baris yang memiliki ID sesi, tanggal, sumber, dan jenis yang sama ditampilkan menjadi satu baris dengan nominal dijumlahkan.
- Satu klik Simpan Form Penjualan adalah satu sesi; satu klik Bayar Kasir adalah satu sesi berisi seluruh item keranjang.
- Input manual dan pengeluaran satu klik Simpan juga masing-masing satu sesi.
- Total, saldo, laba, dan laporan lainnya tetap dihitung dari data transaksi asli.

PEMASANGAN
1. Cadangkan Code.gs yang sedang aktif.
2. Salin APPS_SCRIPT/Code.gs ke Google Apps Script, mengganti Code.gs lama.
3. Index.html dan seluruh file PWA tidak perlu diganti.
4. Deploy > Manage deployments > Edit > New version > Deploy.
5. Uji dengan transaksi baru; cek sheet AT_SESI_KEUANGAN dan TOTAL KEUANGAN.
