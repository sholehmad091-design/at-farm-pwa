AT FARM V10.90 - TOTAL KEUANGAN READ ONLY

Perbaikan:
- getTotalKeuangan tidak lagi menjalankan syncRekapKeuangan_ saat dashboard dibuka.
- Tidak lagi memanggil getCombinedCashFlowRecap dari getTotalKeuangan.
- Tabel membaca langsung REKAP_KEUANGAN kolom A:F.
- Filter tanggal dan saldo berjalan dihitung dari data yang dibaca.
- PWA tetap memakai deployment Apps Script terbaru pada config.js.
