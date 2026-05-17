// Firebase 連線設定
import { initializeApp } from "firebase/app";
import { getAuth } from "firebase/auth";
import { getFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyAYkVEtzIRlGZHxrl6XSWM9lMaHJ08LhYQ",
  authDomain: "lk-learning-74f02.firebaseapp.com",
  projectId: "lk-learning-74f02",
  storageBucket: "lk-learning-74f02.firebasestorage.app",
  messagingSenderId: "246666904449",
  appId: "1:246666904449:web:40741ca7a734bd27599e80",
  measurementId: "G-5LEDTKC6V6"
};

const app = initializeApp(firebaseConfig);
export const auth = getAuth(app);
export const db = getFirestore(app);
export default app;
