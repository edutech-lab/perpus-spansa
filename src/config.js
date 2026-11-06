/* Konfigurasi terpusat. Nilai Firebase dibaca dari file .env (lokal) atau Environment Variables (Vercel). */
export const IS_PREVIEW = import.meta.env.VITE_IS_PREVIEW === 'true';
export const firebaseConfig = {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY,
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN,
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID,
  storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET,
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID,
  appId: import.meta.env.VITE_FIREBASE_APP_ID
};
export const FIREBASE_READY = !!(firebaseConfig.apiKey && firebaseConfig.projectId && firebaseConfig.appId);
export const LOGO = 'https://lh3.googleusercontent.com/d/1GDe6aqXyF_RR6kL3jMCy0vWqaO83jgyp';
export const DEF_SET = {
  namaAplikasi: 'Aplikasi Perpustakaan',
  namaSekolah: 'SMP Negeri 1 Bandar Lampung',
  tahunAjaran: '2026/2027',
  logoUrl: LOGO,
  semboyan: 'Religious, Righteous, Victorious',
  lamaPinjamHari: 7,
  dendaPerHari: 0,
  maksPinjam: 3,
  kategori: ['Buku Pelajaran', 'Novel', 'Pengetahuan Populer', 'Cerita Bergambar', 'Pengembangan Diri', 'Kumpulan Latihan Soal', 'Lainnya']
};
