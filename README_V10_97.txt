V10.97 - Fix tombol MASUK tidak berfungsi.
Penyebab: karakter literal \\n sebelum deklarasi APP_SECTIONS menyebabkan SyntaxError dan menghentikan semua JavaScript, termasuk login.
Perbaikan: hapus karakter tersebut dan validasi syntax JavaScript dengan node --check.
Sidebar tetap dihapus, bottom navigation tetap INPUT | DATA | REKAP. Code.gs dan ID spreadsheet tidak diubah.
Pasang Index.html lalu deploy New Version.
