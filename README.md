# Aplikasi Perpustakaan SMP Negeri 1 Bandar Lampung

Vite + JavaScript, Firebase (Authentication + Firestore), PWA, deploy lewat GitHub dan Vercel.

Menu: Beranda, Katalog Buku, Peminjaman, Pengembalian, **Anggota Siswa**, **Anggota GTK**, Laporan, Pengaturan.

## Menjalankan lokal
```bash
npm install
npm run dev        # http://localhost:5173
npm run build      # hasil di folder dist
```
File `.env.local` sudah berisi konfigurasi Firebase proyek `perpustakaan-5ae00` (tidak ikut ke GitHub). Untuk mencoba tanpa Firebase, isi `VITE_IS_PREVIEW=true` (data contoh hanya di memori).

## Deploy: GitHub lalu Vercel
1. Buat repositori GitHub, unggah isi folder ini (`node_modules`, `dist`, `.env.local` otomatis diabaikan).
2. Vercel > Add New > Project > import repositori. Framework **Vite**, Build `npm run build`, Output `dist`.
3. Settings > Environment Variables, isi `VITE_IS_PREVIEW=false` dan semua `VITE_FIREBASE_*` (nilainya sama dengan `.env.local`). Deploy ulang setiap kali variabel diubah.
4. **Penting:** Firebase Console > Authentication > Settings > Authorized domains > tambahkan domain Vercel (`xxx.vercel.app`) dan domain sendiri. Tanpa ini login admin gagal.
5. Firebase Console > Firestore > Rules: tempel isi `firestore.rules` lalu Publish. Pastikan email admin di fungsi `isAdmin()` sudah benar.

## Anggota GTK
Koleksi Firestore baru: `anggota_gtk` (ID dokumen = No. Anggota) dengan field `noAnggota`, `nip`, `nama`, `jenisPtk` (Guru/Tendik), `jenisKelamin` (L/P), `aktif`, `dibuatPada`. No. Anggota adalah isi barcode kartu dan dikenali pada Peminjaman dan Pengembalian. Data lama (siswa, buku, peminjaman) tidak berubah. No. Anggota tidak boleh sama dengan NIS siswa.

## Catatan
- Mode produksi tidak lagi mengisi data contoh ke Firestore; hanya dokumen `pengaturan/umum` yang dibuat otomatis jika kosong.
- Service worker hanya menyimpan aset statis, bukan data Firestore.
