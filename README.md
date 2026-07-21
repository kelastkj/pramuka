# Sistem Pramuka SMKN 2 Marabahan

Frontend aplikasi berjalan sebagai halaman statis GitHub Pages dari `index.html`.
Backend tetap memakai Google Apps Script dari `Kode.gs`.

## Deploy

1. Upload/push `index.html` ke repository GitHub.
2. Aktifkan GitHub Pages dari branch utama dan folder root.
3. Tempel isi `Kode.gs` ke project Apps Script yang terhubung dengan spreadsheet.
4. Deploy ulang Apps Script sebagai Web App dengan akses sesuai kebutuhan pengguna.
   Gunakan deployment versi baru agar endpoint `/exec` menjalankan kode API terbaru.

Endpoint Apps Script yang dipakai frontend saat ini:

`https://script.google.com/macros/s/AKfycbzy5D-tmKoq0kJI4FbTJY0hvsGNSuqXsUfEfSPmMpMhny7WJ94y280Ddp_EU93_Gof0BQ/exec`

Frontend memanggil endpoint tersebut dengan JSONP agar GitHub Pages tidak terkena batasan CORS Apps Script.
