// Firebase Configuration — DEVELOPMENT
const firebaseConfig = {
    apiKey: "AIzaSyC_HizIMah42B6LVWV7ww6dqY5Z8l5vtbU",
    authDomain: "taskrecorder-dev.firebaseapp.com",
    projectId: "taskrecorder-dev",
    storageBucket: "taskrecorder-dev.firebasestorage.app",
    messagingSenderId: "234818156733",
    appId: "1:234818156733:web:62ded6667abb7162e61fd4"
};

firebase.initializeApp(firebaseConfig);
const db = firebase.firestore();
const auth = firebase.auth();
const googleProvider = new firebase.auth.GoogleAuthProvider();

console.log("🔧 DEV Firebase connected — taskrecorder-dev");