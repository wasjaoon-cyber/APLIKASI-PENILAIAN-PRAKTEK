# Aplikasi Penilaian Praktik TSM

Aplikasi penilaian praktik Teknik Sepeda Motor (TSM) untuk digunakan melalui GitHub Pages.

## Isi
- Aplikasi HTML penilaian praktik (`index.html`)
- Daftar siswa TSM dari data Excel
- Pilihan kelas dan nama siswa
- NIS/NISN otomatis
- Penilaian 7 aspek berbobot
- Rekap nilai dan ekspor CSV
- Penyimpanan online melalui Google Apps Script

## URL Google Apps Script
Web App sudah diarahkan ke URL yang diberikan pengguna.

## Cara upload ke GitHub
1. Buat repository baru, misalnya `penilaian-praktik-tsm`.
2. Upload semua file dalam folder ini ke repository.
3. Pastikan file `index.html` berada di root repository. File ini sudah disiapkan otomatis.
4. Buka **Settings → Pages**.
5. Pada **Build and deployment**, pilih **Deploy from a branch**.
6. Pilih branch `main` dan folder `/ (root)`.
7. Klik **Save**.
8. Setelah aktif, buka URL GitHub Pages yang diberikan GitHub.

Catatan: Google Apps Script harus sudah dideploy sebagai Web App dengan akses yang sesuai agar penyimpanan online berjalan.
