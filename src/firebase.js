import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyB4hMA-hLM9PYg6-nGt8TtxIXsXyHqzX_U",
  authDomain: "igm-parking.firebaseapp.com",
  projectId: "igm-parking",
  storageBucket: "igm-parking.firebasestorage.app",
  messagingSenderId: "148289453195",
  appId: "1:148289453195:web:0b703fcba76cce0a52288e",
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
