import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore"; // Baris ini otomatis ditambahkan

const firebaseConfig = {
  apiKey: "AIzaSyB8S5PPM_yad9Bu8XHCnEr-NqFbw8NDl_U",
  authDomain: "://firebaseapp.com",
  projectId: "pembukuan-gamis-niki-dziyab",
  storageBucket: "pembukuan-gamis-niki-dziyab.firebasestorage.app",
  messagingSenderId: "106587516728",
  appId: "1:106587516728:web:9bdb28d4c7376aa9509bcd",
  measurementId: "G-LMEFG4BNZQ"
};

// Initialize Firebase
const app = initializeApp(firebaseConfig);

// Ekspor database Firestore agar bisa dipakai di seluruh file React Anda
export const db = getFirestore(app);
