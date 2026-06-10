# Checklist Pengembangan — Dairy Meruno
> Sistem Informasi Arsip Pribadi - Rumah Sakit

---

## ✅ Selesai

- [x] Buat landing page `index.html` dengan desain teal + orange
- [x] Navbar sticky + logo + CTA button
- [x] Hero section dengan search bar dan dashboard mockup
- [x] Section fitur: Baca Excel, Baca PDF, Database Search
- [x] Upload zone drag & drop (simulasi)
- [x] Search demo real-time dengan filter
- [x] Stats counter animasi
- [x] Footer lengkap
- [x] `jalankan-server.bat` untuk localhost
- [x] `changelog.md` dan `changelog.html`
- [x] `checklist.md` dan `checklist.html`
- [x] Section Changelog & Checklist di halaman utama
- [x] `app.py` Flask backend
- [x] `excel.html` — Excel viewer dengan statistik & search
- [x] `pdf.html` — PDF viewer dengan ekstraksi teks & tabel

---

## 🔄 Sedang Dikerjakan

- [x] Konten diperbarui ke tema arsip rumah sakit
- [x] Upload terhubung ke Flask backend (POST /api/upload)

---

## 📋 Antrian (Prioritas Tinggi)

- [ ] Integrasi **SheetJS** — parsing file Excel nyata di browser
- [ ] Integrasi **PDF.js** — baca dan ekstrak teks dari PDF
- [ ] Simpan hasil parsing ke **IndexedDB** (lokal, tanpa server)
- [ ] Halaman **Dashboard** — statistik file, aktivitas terbaru
- [ ] Halaman **Manajemen Arsip** — daftar file, hapus, rename

---

## 📋 Antrian (Prioritas Menengah)

- [ ] Autentikasi sederhana (PIN 4-6 digit)
- [ ] Fitur **Export** hasil pencarian ke Excel / PDF
- [ ] Filter pencarian lanjutan (tanggal upload, ukuran file, tipe)
- [ ] Highlight kata kunci di hasil pencarian
- [ ] Dark mode toggle
- [ ] Preview dokumen (PDF viewer, Excel viewer)

---

## 💡 Ide & Backlog

- [ ] Tagging / kategori arsip (Rekam Medis, Administrasi, Laporan, dll)
- [ ] Versi mobile-responsive penuh
- [ ] Print / cetak daftar arsip
- [ ] Backup & restore database lokal
- [ ] Multi-user dengan profil terpisah
- [ ] Notifikasi dokumen kadaluarsa / perlu diperbarui
