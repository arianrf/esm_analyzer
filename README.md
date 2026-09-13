# ESM Activity Analyzer

Program ini membaca langsung file Excel laporan aktivitas yang diunduh dari LMS Edlink. Program menggabungkan data dari seluruh sheet dan menghasilkan rekap peserta, ringkasan sesi, daftar tindak lanjut, dan detail aktivitas.

## Indikator

- Akses materi: 25%
- Pengumpulan tugas: 40%
- Keikutsertaan kuis: 35%

Kategori aktivitas:

- Skor 75-100: Aktif
- Skor 50-74,99: Cukup Aktif
- Skor di bawah 50: Perlu Ditindaklanjuti

Jumlah views, komentar, dan unduhan ditampilkan sebagai informasi tambahan. Banyaknya klik tidak langsung dianggap sebagai hasil belajar.

## Instalasi

1. Instal Python versi 3.10 atau lebih baru dari https://www.python.org/downloads/.
2. Ekstrak paket program.
3. Buka Terminal atau Command Prompt pada folder `ESM_Activity_Analyzer`.
4. Jalankan:

   ```bash
   python -m pip install -r requirements.txt
   ```

## Menjalankan program

1. Jalankan:

   ```bash
   python -m streamlit run app.py
   ```

2. Browser biasanya terbuka otomatis pada `http://localhost:8501`.
3. Klik **Browse files** dan pilih file Excel laporan Edlink.
4. Pada **Pengaturan Analisis**, pilih hanya materi, tugas, dan kuis yang sudah diberikan atau sudah jatuh tempo.
5. Atur batas minimal kategori Cukup Aktif dan Aktif bila diperlukan.
6. Periksa ringkasan kelas, rekap peserta, ringkasan sesi, dan daftar tindak lanjut.
7. Klik **Unduh Hasil Analisis Excel** untuk menyimpan hasil.

## Isi file hasil

- `Rekap Peserta`: skor dan status setiap peserta.
- `Ringkasan Sesi`: ringkasan aktivitas per sheet atau sesi.
- `Tindak Lanjut`: peserta dengan skor aktivitas di bawah 50.
- `Detail Aktivitas`: data terstruktur per peserta dan konten.

## Catatan

- Program dirancang untuk struktur laporan Edlink dengan identitas pada baris 1-3, header konten pada baris 5, subheader pada baris 6, dan data peserta mulai baris 7.
- Email digunakan sebagai kunci penggabungan peserta antar-sheet.
- Program tidak login ke Edlink dan tidak memerlukan kata sandi.
- Kegiatan yang belum jatuh tempo dapat dikeluarkan dari perhitungan tanpa mengubah file Excel asli.
