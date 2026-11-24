// Firebase Configuration (Compat Mode)
const firebaseConfig = {
  apiKey: "AIzaSyDw8DsykEtoYRDLN8WOBvLoj9IyJwhOC3Q",
  authDomain: "letter-management-f1f19.firebaseapp.com",
  projectId: "letter-management-f1f19",
  storageBucket: "letter-management-f1f19.firebasestorage.app",
  messagingSenderId: "48311265428",
  appId: "1:48311265428:web:4a9af5080fc82b6517ec81",
  measurementId: "G-3WC0B2VGDR"
};

// Initialize Firebase using the Compat SDK
if (!firebase.apps.length) {
  firebase.initializeApp(firebaseConfig);
  console.log("🔥 Firebase Cloud Connected Successfully!");
}

const db_cloud = firebase.firestore();
