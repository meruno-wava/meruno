# Changelog — Dairy Meruno
> Sistem Informasi Arsip Pribadi - Rumah Sakit

Semua perubahan signifikan pada project dicatat di sini.
Format mengikuti [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [0.1.0] — 2026-06-10

### Ditambahkan
- `index.html` — Landing page utama Dairy Meruno
  - Navbar sticky dengan logo dan CTA
  - Hero section: judul animasi, search bar, dashboard mockup interaktif
  - Section Fitur: 3 card (Baca Excel, Baca PDF, Database Search)
  - Section Upload: drag & drop zone, progress bar, preview hasil parsing
  - Section Search Demo: pencarian real-time + filter (Semua / Excel / PDF / Tahun / Kategori)
  - Section Stats: 4 counter animasi
  - Section CTA dan Footer lengkap
  - Section Changelog & Checklist (terintegrasi di halaman)
- `jalankan-server.bat` — Script start localhost:8080 via Python
- `changelog.md` + `changelog.html` — Riwayat perubahan
- `checklist.md` + `checklist.html` — Daftar tugas pengembangan
- `claude.md` — Dokumentasi sesi dan keputusan teknis

### Keputusan Desain
- Palet warna: teal (#0d6e7a) + orange (#e8821e), terinspirasi wavahusada.com
- Font: Playfair Display (heading) + Nunito Sans (body) + Space Mono (data/angka)
- Single HTML file — portabel, tidak perlu build tool
- Vanilla JS — ringan, cocok untuk prototype awal

### Catatan
- Fitur upload dan search masih **simulasi** (frontend only)
- Backend belum diintegrasikan

---

## [Unreleased]

### Direncanakan
- Integrasi SheetJS untuk parsing Excel nyata di browser
- Integrasi PDF.js untuk membaca PDF
- Halaman Dashboard pasca-login
- Halaman manajemen file/arsip
- Fitur export hasil pencarian ke Excel/PDF
- Autentikasi sederhana (PIN/password)
- Penyimpanan data lokal (IndexedDB)
