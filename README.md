# KPI NOS Trend Dashboard

Dashboard Streamlit untuk membandingkan nilai `RealKPI` dan `NoteUserKPI` berdasarkan KPI NOS, region, PIC NOS, dan minggu.

## Fungsi Utama

- Menampilkan perbandingan `RealKPI` dan `NoteUserKPI`.
- Memisahkan analisis berdasarkan setiap `KPI INDICES` dalam bentuk tab.
- Menyediakan ringkasan KPI dalam card.
- Menampilkan tren mingguan, perbandingan per region, dan perbandingan per PIC NOS.
- Mendukung filter `Jenis KPI`, `Region`, `PIC NOS`, dan `Week`.

## Cara Menggunakan

1. Buka aplikasi dashboard.
2. Gunakan filter di sidebar kiri untuk membatasi data yang ingin dianalisis.
3. Pilih tab `KPI INDICES` yang ingin dilihat.
4. Baca card ringkasan untuk melihat nilai total atau persentase utama.
5. Gunakan grafik tren mingguan untuk melihat perubahan dari minggu ke minggu.
6. Gunakan grafik per region dan per PIC NOS untuk melihat area atau PIC yang paling terdampak.
7. Lihat tabel detail di bagian bawah untuk memeriksa angka per baris data.

## Catatan Filter

- Filter `PIC NOS` akan menyesuaikan pilihan berdasarkan `Region` dan `Jenis KPI` yang sedang dipilih.
- Jika suatu filter dikosongkan, aplikasi menganggap semua opsi pada filter tersebut dipilih.
- Filter `Week` tetap berdiri sendiri agar pilihan minggu tidak terkunci oleh filter lain.

## Sumber Data

Data utama dibaca langsung dari Google Sheets, sedangkan format KPI dibaca dari file `FORMATTING.xlsx`.

File `FORMATTING.xlsx` menentukan:

- jenis metrik: `Percentage` atau `Absolute Number`
- kategori KPI: `Main KPI` atau `Support KPI`
- arah KPI: `Higher Better` atau `Lower Better`

## Menjalankan Lokal

Install dependency:

```bash
pip install -r requirements.txt
```

Jalankan aplikasi:

```bash
streamlit run app.py
```
