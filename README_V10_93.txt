AT FARM V10.93 - FIX TABEL TOTAL KEUANGAN

Perbaikan utama:
- Memperbaiki error JavaScript pada render tabel TOTAL KEUANGAN.
- Penyebab: bagian TOTAL KEUANGAN memanggil fungsi esc() yang tidak tersedia; fungsi aplikasi yang benar adalah escapeHtml().
- Akibat bug lama: kartu ringkasan berhasil berubah, lalu render baris berhenti sehingga tabel tetap "Memuat Total Keuangan...".
- Filter tanggal dan Print Total Keuangan diperbaiki menggunakan escapeHtml().
- Sumber data tetap langsung dari sheet transaksi/lokasi, KASIR_TRANSAKSI, dan PENGELUARAN; tidak memakai REKAP_KEUANGAN.
- Endpoint PWA tetap Apps Script terbaru pengguna.
