PASAR AT FARM V11.07 - PRINT BERSIH + TOTAL KESELURUHAN

Perbaikan:
- Hilangkan navigasi bawah FORM/INPUT/TOTAL KEUANGAN/DATA/REKAP dari print PDF.
- Baris TOTAL KESELURUHAN berada di akhir tabel print dan Export JPG.
- Total hanya menjumlahkan kolom uang aditif (bukan harga satuan/saldo berjalan).
- Tampilan aplikasi biasa tidak berubah.

Pemasangan:
1. Cadangkan Index.html Apps Script lama.
2. Salin APPS_SCRIPT/Index.html ke Index.html proyek Apps Script AT FARM.
3. Save, Deploy > Manage deployments > Edit > New version > Deploy.
4. Untuk GitHub Pages, jangan unggah folder APPS_SCRIPT ke root; PWA GitHub terpisah dari Apps Script.
