# PL REALISASI untuk Google Apps Script

Paket ini berisi frontend dan backend yang berjalan di Google Apps Script. Template `Tamplate_PL.xlsx` sudah tertanam di `TemplateData.gs`. Saat website pertama kali dibuka, aplikasi otomatis membuat folder kerja pribadi dan template Google Sheets di Drive; tidak perlu menjalankan setup secara manual.

## Deploy lewat editor Google Apps Script

1. Ekstrak `PL_Google_Apps_Script.zip`, lalu buka [script.google.com](https://script.google.com/) dan buat **New project**.
2. Isi file **Code.gs** dengan isi `Code.gs` dari paket. Tambahkan dua file **Script** bernama **Core** dan **TemplateData**, lalu salin isi file `.gs` yang sesuai.
3. Tambahkan file **HTML** bernama **Index**, lalu salin isi `Index.html`.
4. Di **Project Settings**, aktifkan **Show "appsscript.json" manifest file in editor**. Buka `appsscript.json` dan ganti seluruh isinya dengan manifest dari paket. Simpan.
5. Manifest dari paket sudah mendaftarkan **Drive API v3**, identifier **Drive**. Jangan menambahkan Drive lagi lewat tombol **+ Services**. Jika daftar Services belum diperbarui, muat ulang editor setelah manifest disimpan. Untuk proyek dengan Google Cloud project khusus, aktifkan Google Drive API pada proyek Cloud tersebut juga. [Petunjuk layanan Drive Google](https://developers.google.com/apps-script/advanced/drive).
6. Pilih **Deploy → New deployment → Web app**. Pilih **Execute as: Me**. Untuk pemakaian pribadi pilih **Who has access: Only myself**. Untuk penggunaan kantor, pilih akses domain bila tersedia pada akun Workspace. Klik **Deploy**, setujui izin akun Google jika diminta, lalu buka URL yang berakhiran **/exec**. [Panduan deployment Google](https://developers.google.com/apps-script/guides/web#deploy_a_script_as_a_web_app).
7. Tunggu pesan **Menyiapkan aplikasi dan template…** pada kunjungan pertama. Setelah status menjadi **Siap memproses**, langsung unggah GL. Folder dan template disiapkan otomatis; kunjungan berikutnya memakai data setup yang sama.

Semua file bernama tepat seperti di atas. Tidak perlu memasukkan README atau folder pengujian ke editor.

Jika muncul **Found a service identifier used more than once: Drive**, terdapat dua entri Drive pada manifest proyek Google. Buka `appsscript.json`, pilih seluruh isinya, lalu ganti dengan seluruh manifest dari paket; jangan menempelkan manifest di bawah konfigurasi lama. `dependencies.enabledAdvancedServices` harus berisi tepat satu entri Drive v3. Simpan, muat ulang editor, dan deploy versi terbaru.

## Pemakaian

1. Unggah GL `.xlsx`, pilih **Tahun Laporan** dan **Bulan YTD**, lalu klik **Proses**. Header sumber harus tepat lima: `Project`, `GL Account`, `Posting Date`, `In profit center local currency`, `Text`; urutannya boleh berbeda. Sheet pertama dipakai, dan satu file hanya boleh berisi satu proyek.
2. Isi nama serta jabatan **Pembuat**, **Mengetahui**, dan **Menyetujui** sesuai kebutuhan.
3. Periksa **Preview PL**, **Perlu Dicek**, dan **Ringkasan Carry Over**. Edit bulan/tahun pada tab **Perlu Dicek**, lalu klik **Terapkan Koreksi**. Filter dan halaman tabel menjaga baris tetap mudah dicari.
4. Edit justifikasi. **Hasil Evaluasi** di bawahnya mengambil nilai yang sama dengan **P37**.
5. Klik **Download Excel**. Koreksi yang belum diterapkan akan diproses sebelum file dibuat. Output memiliki sheet sumber dengan F:I, sheet PL dengan rumus, tanda tangan, evaluasi, dan YTD di kolom AA yang disembunyikan. Kolom J tersembunyi pada sheet GL dipakai untuk mencocokkan akun dengan spasi tambahan, sambil mempertahankan lima kolom sumber.

## Aturan yang dibawa dari versi Python

- **DIRECT LABOR**, **SALARY AND COMPENSATION**, dan **DEPRECIATION & AMORTIZATION** memakai Posting Date.
- Singkatan bulan diikuti tahun dua digit dibaca sebagai tahun 2000-an: `DES 24` → Desember 2024; `NOV-24` → November 2024. Tahun sebelum Tahun Laporan menjadi `CO <tahun>`.
- Bulan tanpa tahun sampai batas YTD memakai Tahun Laporan; bulan setelah YTD memakai tahun sebelumnya.
- Rentang bulan pada Tahun Laporan memakai Posting Date. Tahun Laporan tanpa bulan memakai Posting Date.
- Angka hanya memasukkan Posting Date pada Tahun Laporan sampai akhir YTD. Kolom bulan setelah YTD menghasilkan nol.
- OPERATING REVENUE dibalik tandanya dari GL. Nilai negatif memakai minus berjarak; nilai nol memakai strip yang posisinya seragam.
- Carry Over pada kolom C dan ringkasannya menampilkan tahun tepat sebelum Tahun Laporan, sesuai template. Carry Over yang lebih lama tetap terlihat pada kolom periode GL dan memengaruhi SELISIH terhadap TOTAL SAP.
- Evaluasi memakai persentase P37 sesuai permintaan. Label NPAT pada kalimat evaluasi mengikuti teks yang sudah digunakan dalam laporan.

Perhitungan backend dan formula output diperiksa sebelum ekspor; bila nilainya berbeda, unduhan dibatalkan dengan pesan error. Rumus operating income per bulan menggunakan `MAX(0, gross profit - overhead)` agar sama dengan preview Python.

## File dan batas pemrosesan

| File | Fungsi |
|---|---|
| `Index.html` | Form unggah, preview, koreksi, carry over, justifikasi, evaluasi, unduh |
| `Code.gs` | Endpoint, konversi GL/template, sesi, pembuatan rumus, ekspor Excel |
| `Core.gs` | Parsing periode dan perhitungan PL murni JavaScript |
| `TemplateData.gs` | Template Excel yang tertanam dan pengaturan cetaknya |
| `appsscript.json` | Runtime V8, zona waktu, Drive v3, izin layanan |

Batas aplikasi adalah **5 MB** dan **5.000 baris GL** per unggahan. Google membatasi satu eksekusi Apps Script hingga **6 menit** dan ekspor Drive hingga **10 MB**. Penggunaan layanan juga mengikuti kuota akun Google. [Kuota Apps Script](https://developers.google.com/apps-script/guides/services/quotas), [batas ekspor Drive](https://developers.google.com/workspace/drive/api/reference/rest/v3/files/export).

GL sementara dan data sesi disimpan dalam folder Drive pemilik aplikasi. Sesi berlaku **4 jam**, terikat token acak dan identitas tab browser. Trigger setiap jam memindahkan file sesi kedaluwarsa ke Trash. Salinan laporan sementara dipindahkan ke Trash setelah unduhan dibuat. File tidak dibagikan lewat tautan publik; Excel dikirim langsung ke browser.

Template melalui konversi Google Sheets sebelum ekspor Excel. Paket mengembalikan area cetak dan pengaturan halaman template, serta format angka custom setelah ekspor. Beberapa detail tampilan Excel yang tidak didukung Google Sheets dapat berubah; periksa satu hasil contoh di Excel setelah deployment.

## Verifikasi setelah deployment

Jalankan fungsi **runSmokeTest_** dari editor. Pengujian membuat GL contoh kecil, melakukan impor, koreksi, penghitungan, dan ekspor, lalu menyimpan satu file uji di folder kerja serta menulis linknya di log. Periksa tanda tangan, justifikasi, P37, format nol/minus, dan batas cetak pada file tersebut. Fungsi setup dan pengujian bersifat privat terhadap panggilan frontend.

Uji logika lokal telah membandingkan semua hasil parser dan preview terhadap aplikasi Python menggunakan GL contoh, YTD penuh dan Februari, singkatan tahun, akun pengecualian, urutan header, dan koreksi F/G. Layanan Google Drive/Sheets serta izin deployment harus diuji di akun Google melalui langkah di atas.

Jika memperbarui kode, pilih **Deploy → Manage deployments → Edit → New version → Deploy**. URL `/exec` mengikuti versi deployment, bukan perubahan editor yang belum dideploy.

## Pemeliharaan kode lokal

Dari folder utama proyek, bangun ulang paket setelah template/kode berubah:

```powershell
python google_apps_script/tools/build_package.py
```

Jalankan pengujian port menggunakan Node.js dan Python yang memiliki pandas/openpyxl:

```powershell
node google_apps_script/tests/parity.test.js --python python
node google_apps_script/tests/backend.test.js --python python
```
