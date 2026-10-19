# Pindai

Aplikasi scan untuk HP Scanjet Enterprise Flow 7000 s3 di Windows. Hasil scan tampil dulu untuk diperiksa, lalu disimpan sebagai JPG per halaman atau satu PDF.

## Kebutuhan

- Windows 10 atau 11
- Python 3.9 atau lebih baru (centang "Add python.exe to PATH" saat memasang)
- Driver HP Scanjet Enterprise Flow 7000 s3 (paket HP Scan)

Modul Python yang dipakai: `pywin32`, `pillow`, `pytwain`. Semuanya dipasang otomatis oleh `Pasang Pindai.bat`.

## Pasang

1. Simpan folder ini di lokasi tetap, misalnya `C:\BIT2026\Pindai`.
2. Jalankan `Pasang Pindai.bat`. File ini memasang modul, membuat shortcut di Desktop dan Start Menu, lalu membuka Pindai.
3. Untuk menyematkan ke taskbar: buka Start, ketik Pindai, klik kanan, pilih Pin to taskbar.

Kalau folder dipindah, jalankan `Pasang Pindai.bat` lagi agar shortcut menunjuk ke lokasi baru.

## Cara kerja

**Mesin scan.** Pindai memakai driver TWAIN HP ("HP TWAIN USB") bila tersedia, sehingga deteksi ukuran, pelurusan, orientasi otomatis, warna otomatis, buang halaman kosong, dan pencahayaan dikerjakan driver HP seperti di HP Scan. Bila driver TWAIN tidak ada, Pindai memakai WIA dan mengolah gambar sendiri.

**Alur.** Scan, periksa (putar, hapus, buka besar), lalu Simpan gambar atau Simpan PDF. Halaman belum ditulis ke folder tujuan sebelum tombol simpan ditekan.

| Pengaturan HP Scan | Di Pindai |
|---|---|
| Input Page Size: Detect Size | Tab Scan, Ukuran input: Otomatis |
| Output Page Size: Same as Input | Tab Hasil, Ukuran hasil: Sama dengan input |
| Page Orientation: Auto Orient | Tab Hasil, Orientasi otomatis |
| Color Mode: Auto Detect Color | Tab Scan, Mode warna: Otomatis (warna / abu-abu) |
| Exposure: Default | Tab Hasil, Kecerahan 0 dan Kontras 0 |

## Pintasan

| Tombol | Fungsi |
|---|---|
| F5 | Scan |
| Esc | Berhenti setelah lembar yang sedang ditarik |
| Ctrl+S | Simpan gambar (JPG per halaman) |
| Ctrl+Shift+S | Simpan PDF (satu berkas) |
| Ctrl+O | Buka folder hasil |
| L / R / Delete | Putar kiri, putar kanan, hapus halaman terpilih |

## Penamaan berkas

- Gambar: `NAMA_001.jpg`, `NAMA_002.jpg`, dan seterusnya. Nomor melanjutkan berkas yang sudah ada, jadi tidak ada yang tertimpa.
- PDF: `NAMA.pdf`. Bila sudah ada, berkas baru menjadi `NAMA_2.pdf`.
- Nama kosong: `SCAN_ddmmyy_jamdetik`.

## Masalah umum

- **Scanner tidak muncul di daftar:** nyalakan scanner, cek kabel USB, tekan Muat ulang.
- **"Scanner sedang dipakai aplikasi lain":** tutup HP Scan lalu coba lagi.
- **Driver HP tidak terdeteksi:** jalankan `alat/Cek Driver HP.bat` dan periksa laporan di Desktop.

## Isi folder

- `pindai.py`: aplikasi
- `pindai.ico`, `pindai.svg`: ikon
- `Pasang Pindai.bat`: pemasang dan pembuat shortcut
- `Push ke GitHub.bat`: mengirim perubahan ke GitHub
- `alat/`: pemeriksa driver HP
