import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";
import firebaseConfig from "../../firebase-applet-config.json";

// Initialize Firebase using the project applet configuration
const app = initializeApp(firebaseConfig);

// Initialize Firestore with custom database ID from configuration
export const db = getFirestore(app, firebaseConfig.firestoreDatabaseId || "(default)");
