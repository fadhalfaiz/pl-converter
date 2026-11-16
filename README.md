# PL REALISASI Converter

Aplikasi Streamlit untuk membuat laporan **PL REALISASI** per proyek dari GL SAP. Sheet PL mengikuti `Tamplate_PL.xlsx`; data GL tetap di sheet asal, dengan kolom bantu F:I dan rumus yang dapat dihitung ulang di Excel.

Nilai OPERATING REVENUE dibalik tanda dari GL saat masuk PL; akun lain memakai tanda aslinya. Angka hasil yang negatif menampilkan tanda minus terpisah dari angkanya, dan nol ditampilkan sebagai `-`, tanpa mengubah nilai yang dipakai rumus.

## Versi Google Apps Script

Frontend dan backend Google Apps Script tersedia pada folder `google_apps_script`, dengan panduan di `google_apps_script/README.md`. Paket deployment dibuat dengan `python google_apps_script/tools/build_package.py` dan tersimpan di `dist/PL_Google_Apps_Script.zip`. Versi ini memakai Google Drive/Sheets untuk membaca dan mengekspor Excel, dengan template yang sudah tertanam dalam paket.

## Jalankan dengan Docker di localhost

Pastikan Docker Desktop sudah terpasang dan berjalan. Dari folder proyek:

```powershell
docker compose up --build -d
```

Jika PowerShell tidak mengenali `docker` pada instalasi Docker Desktop di komputer ini, jalankan CLI dari lokasi instalasinya:

```powershell
& "$env:LOCALAPPDATA\Programs\DockerDesktop\resources\bin\docker.exe" compose up --build -d
```

Buka [http://localhost:8501](http://localhost:8501). Port dipetakan hanya ke `127.0.0.1`, sehingga aplikasi diakses dari komputer ini. Untuk melihat log atau menghentikan aplikasi:

```powershell
docker compose logs -f
docker compose down
```

File GL diunggah melalui browser dan hasil Excel diunduh melalui tombol aplikasi; tidak perlu memasang folder tambahan ke container.

## Jalankan tanpa Docker

Gunakan Python 3.11 atau lebih baru.

```powershell
pip install -r requirements.txt
streamlit run app.py
```

1. Unggah GL `.xlsx` dengan lima header `Project`, `GL Account`, `Posting Date`, `In profit center local currency`, dan `Text`. Urutan kolom bebas; aplikasi membaca sheet pertama. Satu file harus berisi satu proyek.
2. Pilih **Tahun Laporan** dan **Bulan YTD (sampai)**. Bulan YTD bawaan adalah Desember; pilih bulan lebih awal untuk membatasi laporan sampai akhir bulan tersebut. Aplikasi memakai `Tamplate_PL.xlsx` bawaan secara otomatis.
3. Edit nama dan jabatan pada bagian **Pembuat**, **Mengetahui**, dan **Menyetujui** bila perlu; isian tersebut akan tampil pada bagian tanda tangan di Excel. Jabatan bawaan "Mengetahui" adalah "Manager TOPM Dept.". Klik **Proses**. Periksa tab **Preview PL**, koreksi kolom **Bulan (Text)** dan **Tahun (Text)** pada tab **Perlu Dicek**, serta lihat rincian di **Ringkasan Carry Over**.
4. Di bawah poin justifikasi, website dan Excel menampilkan **Hasil Evaluasi**; persentase NPAT pada kalimat evaluasi dihitung dari TOTAL NPAT (P48) dibagi OPERATING REVENUE (P4).
4. Edit poin **Justifikasi** bila diperlukan, lalu klik **Download Excel**.

Kolom F/G di Excel berwarna biru dan dapat diedit. **DIRECT LABOR, SALARY AND COMPENSATION, dan DEPRECIATION & AMORTIZATION selalu mengikuti Posting Date**, apa pun isi Bulan (Text) dan Tahun (Text). Untuk GL lain, bulan pada teks tanpa tahun memakai Tahun Laporan jika bulannya sampai dengan Bulan YTD. Jika bulan melewati YTD, Tahun (Text) dianggap tahun sebelumnya (2025 bila Tahun Laporan 2026) dan masuk Carry Over. Singkatan bulan yang langsung diikuti tahun dua digit dibaca sebagai tahun 2000-an, misalnya `DES 24` menjadi Desember 2024. Tahun (Text) sebelum Tahun Laporan juga masuk Carry Over. Jika Tahun (Text) sama dengan Tahun Laporan tetapi Bulan (Text) kosong, Periode PL mengikuti Posting Date. Bulan dan tahun yang lengkap mengikuti periode pada teks. Kolom H dan seluruh angka PL memakai rumus sehingga perubahan F/G atau tambahan baris GL dihitung ulang saat workbook dibuka di Excel. Kolom H merujuk tahun laporan di `PL!C2` untuk menentukan carry over. Pilih ulang Bulan YTD dan klik **Proses** bila ingin mengubah tahun yang diasumsikan untuk bulan tanpa tahun.

Bulan YTD tersimpan pada kolom `AA` yang disembunyikan di sheet PL agar tidak tampil pada laporan, tetapi tetap dipakai rumus. Baris GL dengan Posting Date setelah bulan YTD tidak masuk preview, carry over, maupun total Excel; kolom bulan setelah batas YTD menghasilkan nol. Rentang dua bulan pada tahun laporan (misalnya `JAN-MAR 26`) mengikuti bulan Posting Date. Rentang tahun sebelumnya tetap masuk Carry Over. File unduhan memakai nama `PL_<Project>_<Tahun>_YTD_<Bulan>.xlsx`.

Jika LibreOffice tersedia pada sistem, aplikasi menjalankan `soffice` headless untuk menghitung cache rumus dan memeriksa error `#NAME?`, `#VALUE!`, serta `#REF!` sebelum unduhan. Bila tidak tersedia, Excel menghitung rumus saat file dibuka. Proses LibreOffice dapat memerlukan waktu pada GL besar karena rumus memakai rentang kolom penuh.

## Uji

```powershell
pytest -q
```

Contoh GL ada di `examples/GL Aviasi Pulau Layang.xlsx`. Uji integrasi memeriksa enam angka penerimaan utama untuk tahun 2026, semua akun GL, rumus, format template, dan perubahan preview setelah koreksi F/G.
