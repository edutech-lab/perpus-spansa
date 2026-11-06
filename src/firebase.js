/* Inisialisasi Firebase (SDK modular). Modul ini hanya dimuat ketika IS_PREVIEW = false. */
import { initializeApp } from 'firebase/app';
import { getAuth, onAuthStateChanged, signInWithEmailAndPassword, signOut } from 'firebase/auth';
import { collection, doc, getDocs, getFirestore, increment, limit, onSnapshot, orderBy, query, where, writeBatch } from 'firebase/firestore';

export function createFirebase(config) {
  const app = initializeApp(config);
  return {
    fb: { collection, doc, getDocs, increment, limit, onSnapshot, orderBy, query, where, writeBatch, onAuthStateChanged, signInWithEmailAndPassword, signOut },
    fauth: getAuth(app),
    fdb: getFirestore(app)
  };
}
