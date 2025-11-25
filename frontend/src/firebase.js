// src/firebase.js
import { initializeApp } from "firebase/app";
import {
  getAuth,
  GoogleAuthProvider,
  setPersistence,
  browserLocalPersistence,
} from "firebase/auth";
import { getFirestore } from "firebase/firestore";

// Your Firebase config (from Firebase dashboard)
const firebaseConfig = {
  apiKey: "AIzaSyCVbacu_7PUSVB6zOh-stJohH4Ea-75cJQ",
  authDomain: "my-doc-app-d787f.firebaseapp.com",
  projectId: "my-doc-app-d787f",
  storageBucket: "my-doc-app-d787f.firebasestorage.app",
  messagingSenderId: "885664964442",
  appId: "1:885664964442:web:62d894a22edfa057509d63",
  measurementId: "G-7TC0T5PDWC"
};

// Initialize Firebase
const app = initializeApp(firebaseConfig);

// Auth instance
export const auth = getAuth(app);

// Always keep the user logged in
setPersistence(auth, browserLocalPersistence);

// Google login provider
export const googleProvider = new GoogleAuthProvider();

// Firestore database
export const db = getFirestore(app);
