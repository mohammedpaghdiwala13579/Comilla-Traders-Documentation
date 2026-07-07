import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyD8QK5nJ2TVllAjCeMojOD9wZPBJAeuy4E",
  authDomain: "gifted-gift-rdpgw.firebaseapp.com",
  projectId: "gifted-gift-rdpgw",
  storageBucket: "gifted-gift-rdpgw.firebasestorage.app",
  messagingSenderId: "102803571663",
  appId: "1:102803571663:web:b97a4b94273ec54269c8ba"
};

// Initialize Firebase
const app = initializeApp(firebaseConfig);

// Initialize Firestore with custom database ID
export const db = getFirestore(app, "ai-studio-2c592343-56ab-4d40-a2ac-d15fed703e91");
